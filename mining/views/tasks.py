from django.http import JsonResponse
from django.views.decorators.http import require_POST
from django.core.cache import cache
from decimal import Decimal
from datetime import date
import threading
import time
from ..models import APIData, Payout
from ..api_utils import fetch_all_api_data, get_historical_btc_price


_CACHE_TIMEOUT = 3600  # 1 hour

_bulk_fetch_lock = threading.Lock()
_api_fetch_lock = threading.Lock()


def _get_bulk_fetch_status():
    return cache.get(_BULK_FETCH_CACHE_KEY, {
        'running': False, 'total': 0, 'processed': 0,
        'updated': 0, 'skipped': 0, 'errors': 0,
        'error_details': [], 'message': '',
    })


def _set_bulk_fetch_status(status):
    cache.set(_BULK_FETCH_CACHE_KEY, status, _CACHE_TIMEOUT)


def _get_api_fetch_status():
    return cache.get(_API_FETCH_CACHE_KEY, {
        'running': False, 'message': '', 'success': None,
    })


def _set_api_fetch_status(status):
    cache.set(_API_FETCH_CACHE_KEY, status, _CACHE_TIMEOUT)




def _bulk_fetch_closing_prices_task():
    """Background task: fetch closing prices in sub-batches with delay to respect API rate limits."""
    from .api_utils import get_historical_btc_price as fetch_price

    BATCH_SIZE = 5
    DELAY_BETWEEN_BATCHES = 3  # seconds

    today = date.today()

    payouts = list(
        Payout.objects.filter(payout_date__isnull=False).order_by('payout_date')
    )

    payouts_to_fetch = []
    for p in payouts:
        if p.closing_price_fetched_at is None:
            payouts_to_fetch.append(p)
        elif p.closing_price_fetched_at <= p.payout_date:
            payouts_to_fetch.append(p)

    status = _get_bulk_fetch_status()
    status.update({
        'total': len(payouts_to_fetch), 'processed': 0,
        'updated': 0, 'skipped': 0, 'errors': 0,
        'error_details': [],
        'message': f'Processing {len(payouts_to_fetch)} payouts...',
    })
    _set_bulk_fetch_status(status)

    for i in range(0, len(payouts_to_fetch), BATCH_SIZE):
        batch = payouts_to_fetch[i:i + BATCH_SIZE]

        for payout in batch:
            try:
                historical_price = fetch_price(payout.payout_date)
                payout.closing_price = Decimal(str(historical_price))
                payout.closing_price_fetched_at = today
                payout.save()
                status = _get_bulk_fetch_status()
                status['updated'] += 1
            except Exception as e:
                status = _get_bulk_fetch_status()
                status['errors'] += 1
                status['error_details'].append(
                    f'Payout #{payout.pk} ({payout.payout_date}): {str(e)}'
                )

            status['processed'] += 1
            _set_bulk_fetch_status(status)

        if i + BATCH_SIZE < len(payouts_to_fetch):
            time.sleep(DELAY_BETWEEN_BATCHES)

    status = _get_bulk_fetch_status()
    skipped = len(payouts) - len(payouts_to_fetch)
    status['skipped'] = skipped
    status['message'] = (
        f'Completed: {status["updated"]} updated, '
        f'{skipped} skipped, '
        f'{status["errors"]} errors.'
    )
    status['running'] = False
    _set_bulk_fetch_status(status)




@require_POST
def fetch_closing_price(request, payout_id):
    """Fetch historical BTC price for payout date and update closing_price field"""
    try:
        payout = get_object_or_404(Payout, pk=payout_id)

        if not payout.payout_date:
            return JsonResponse({
                'success': False,
                'error': 'Payout date is required to fetch closing price'
            })

        # Fetch historical BTC price for the payout date
        historical_price = get_historical_btc_price(payout.payout_date)

        # Update the payout's closing_price and fetched_at fields
        payout.closing_price = Decimal(str(historical_price))
        payout.closing_price_fetched_at = date.today()
        payout.save()

        return JsonResponse({
            'success': True,
            'closing_price': float(payout.closing_price),
            'formatted_price': f'${payout.closing_price:,.2f}'
        })

    except Exception as e:
        return JsonResponse({
            'success': False,
            'error': f'Failed to fetch closing price: {str(e)}'
        })


# Background task state via Django cache (shareable across processes)
# For multi-process production, configure a shared cache backend (Redis, Memcached, or database)
# in settings.py. The default LocMemCache works for single-process dev servers.
from django.core.cache import cache

_BULK_FETCH_CACHE_KEY = 'bulk_fetch_closing_prices_status'
_API_FETCH_CACHE_KEY = 'api_fetch_status'
_CACHE_TIMEOUT = 3600  # 1 hour

_bulk_fetch_lock = threading.Lock()
_api_fetch_lock = threading.Lock()




@require_POST
def bulk_fetch_closing_prices(request):
    """Trigger bulk closing price fetch as a background task."""
    with _bulk_fetch_lock:
        status = _get_bulk_fetch_status()
        if status['running']:
            return JsonResponse({
                'success': False,
                'error': 'A bulk fetch is already in progress.'
            })
        status = {
            'running': True, 'total': 0, 'processed': 0,
            'updated': 0, 'skipped': 0, 'errors': 0,
            'error_details': [], 'message': 'Starting...',
        }
        _set_bulk_fetch_status(status)

    thread = threading.Thread(target=_bulk_fetch_closing_prices_task, daemon=True)
    thread.start()

    return JsonResponse({'success': True, 'message': 'Bulk fetch started.'})




def bulk_fetch_closing_prices_status(request):
    """Return the current status of the bulk closing price fetch task."""
    status = _get_bulk_fetch_status()
    return JsonResponse({
        'running': status['running'],
        'total': status['total'],
        'processed': status['processed'],
        'updated': status['updated'],
        'skipped': status['skipped'],
        'errors': status['errors'],
        'message': status['message'],
        'error_details': list(status['error_details']),
    })




def _fetch_api_data_task():
    """Background task: fetch all API data and save to database."""
    try:
        status = _get_api_fetch_status()
        status['message'] = 'Fetching API data...'
        _set_api_fetch_status(status)

        result = fetch_all_api_data()

        if result['success']:
            api_data = APIData.get_api_data()
            api_data.bitcoin_price_usd = result['bitcoin_price_usd']
            api_data.network_hashrate_ehs = result['network_hashrate_ehs']
            api_data.network_difficulty = result['network_difficulty']
            api_data.avg_block_fees_24h = result['avg_block_fees_24h']
            api_data.save()

            status = _get_api_fetch_status()
            status['message'] = result['message']
            status['success'] = True
        else:
            status = _get_api_fetch_status()
            status['message'] = result['message']
            status['success'] = False
    except Exception as e:
        status = _get_api_fetch_status()
        status['message'] = f'Unexpected error: {str(e)}'
        status['success'] = False
    finally:
        status['running'] = False
        _set_api_fetch_status(status)




@require_POST
def trigger_fetch_api_data(request):
    """Trigger API data fetch as a background task."""
    with _api_fetch_lock:
        status = _get_api_fetch_status()
        if status['running']:
            return JsonResponse({
                'success': False,
                'error': 'API fetch is already in progress.'
            })
        _set_api_fetch_status({
            'running': True, 'message': 'Starting...', 'success': None,
        })

    thread = threading.Thread(target=_fetch_api_data_task, daemon=True)
    thread.start()

    return JsonResponse({'success': True, 'message': 'API fetch started.'})




def fetch_api_data_status(request):
    """Return the current status of the API data fetch task."""
    status = _get_api_fetch_status()
    return JsonResponse({
        'running': status['running'],
        'message': status['message'],
        'success': status['success'],
    })


