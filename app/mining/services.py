from django.db.models import Sum, Avg
from django.db.models.functions import TruncMonth
from decimal import Decimal
from .models import RemoteMiningPlatform, Miner, Settings, APIData, Payout, Expense


def get_capex_opex_data():
    """Gather all CAPEX/OPEX dashboard data."""
    total_expenses = Expense.objects.aggregate(total=Sum('expense_amount'))['total'] or Decimal('0')
    total_capex = Expense.objects.filter(category='CAPEX').aggregate(total=Sum('expense_amount'))['total'] or Decimal('0')
    total_opex = Expense.objects.filter(category='OPEX').aggregate(total=Sum('expense_amount'))['total'] or Decimal('0')

    platforms = RemoteMiningPlatform.objects.all()
    platform_expenses = []

    for platform in platforms:
        platform_total = Expense.objects.filter(platform=platform).aggregate(total=Sum('expense_amount'))['total'] or Decimal('0')
        platform_capex = Expense.objects.filter(platform=platform, category='CAPEX').aggregate(total=Sum('expense_amount'))['total'] or Decimal('0')
        platform_opex = Expense.objects.filter(platform=platform, category='OPEX').aggregate(total=Sum('expense_amount'))['total'] or Decimal('0')

        if platform_total > 0:
            platform_expenses.append({
                'platform': platform,
                'total': platform_total,
                'capex': platform_capex,
                'opex': platform_opex
            })

    monthly_capex = Expense.objects.filter(category='CAPEX').annotate(
        month=TruncMonth('expense_date')
    ).values('month').annotate(
        total=Sum('expense_amount')
    ).order_by('month')

    monthly_capex_by_platform = {}
    for platform in platforms:
        platform_monthly_capex = Expense.objects.filter(
            category='CAPEX', platform=platform
        ).annotate(
            month=TruncMonth('expense_date')
        ).values('month').annotate(
            total=Sum('expense_amount')
        ).order_by('month')

        if platform_monthly_capex:
            monthly_capex_by_platform[platform] = platform_monthly_capex

    monthly_opex = Expense.objects.filter(category='OPEX').annotate(
        month=TruncMonth('expense_date')
    ).values('month').annotate(
        total=Sum('expense_amount')
    ).order_by('month')

    monthly_opex_by_platform = {}
    for platform in platforms:
        platform_monthly_opex = Expense.objects.filter(
            category='OPEX', platform=platform
        ).annotate(
            month=TruncMonth('expense_date')
        ).values('month').annotate(
            total=Sum('expense_amount')
        ).order_by('month')

        if platform_monthly_opex:
            monthly_opex_by_platform[platform] = platform_monthly_opex

    all_months = set()
    for item in monthly_capex:
        all_months.add(item['month'])
    for item in monthly_opex:
        all_months.add(item['month'])
    for platform_data in monthly_capex_by_platform.values():
        for item in platform_data:
            all_months.add(item['month'])
    for platform_data in monthly_opex_by_platform.values():
        for item in platform_data:
            all_months.add(item['month'])

    all_months = sorted(list(all_months))

    return {
        'total_expenses': total_expenses,
        'total_capex': total_capex,
        'total_opex': total_opex,
        'platform_expenses': platform_expenses,
        'monthly_capex': monthly_capex,
        'monthly_capex_by_platform': monthly_capex_by_platform,
        'monthly_opex': monthly_opex,
        'monthly_opex_by_platform': monthly_opex_by_platform,
        'all_months': all_months,
    }


def get_income_data():
    """Gather all Income dashboard data."""
    api_data = APIData.get_api_data()
    current_btc_price = float(api_data.bitcoin_price_usd) if api_data.bitcoin_price_usd else 0

    total_income_btc = Payout.objects.aggregate(total=Sum('payout_amount'))['total'] or Decimal('0')
    total_income_usd_then = Payout.objects.aggregate(total=Sum('value_at_payout'))['total'] or Decimal('0')

    total_income_usd_now = Decimal('0')
    if current_btc_price > 0:
        total_income_usd_now = total_income_btc * Decimal(str(current_btc_price))

    platforms = RemoteMiningPlatform.objects.all()
    platform_income = []

    for platform in platforms:
        platform_btc = Payout.objects.filter(platform=platform).aggregate(total=Sum('payout_amount'))['total'] or Decimal('0')
        platform_usd_then = Payout.objects.filter(platform=platform).aggregate(total=Sum('value_at_payout'))['total'] or Decimal('0')
        platform_usd_now = Decimal('0')
        if current_btc_price > 0:
            platform_usd_now = platform_btc * Decimal(str(current_btc_price))

        if platform_btc > 0:
            platform_income.append({
                'platform': platform,
                'total_btc': platform_btc,
                'total_usd_then': platform_usd_then,
                'total_usd_now': platform_usd_now
            })

    monthly_income_btc = Payout.objects.annotate(
        month=TruncMonth('payout_date')
    ).values('month').annotate(
        total_btc=Sum('payout_amount'),
        total_usd_then=Sum('value_at_payout')
    ).order_by('month')

    for item in monthly_income_btc:
        item['total_usd_now'] = item['total_btc'] * Decimal(str(current_btc_price)) if current_btc_price > 0 else Decimal('0')

    monthly_income_by_platform = {}
    for platform in platforms:
        platform_monthly_income = Payout.objects.filter(
            platform=platform
        ).annotate(
            month=TruncMonth('payout_date')
        ).values('month').annotate(
            total_btc=Sum('payout_amount'),
            total_usd_then=Sum('value_at_payout')
        ).order_by('month')

        for item in platform_monthly_income:
            item['total_usd_now'] = item['total_btc'] * Decimal(str(current_btc_price)) if current_btc_price > 0 else Decimal('0')

        if platform_monthly_income:
            monthly_income_by_platform[platform] = platform_monthly_income

    all_months = set()
    for item in monthly_income_btc:
        all_months.add(item['month'])
    for platform_data in monthly_income_by_platform.values():
        for item in platform_data:
            all_months.add(item['month'])

    all_months = sorted(list(all_months))

    return {
        'current_btc_price': current_btc_price,
        'total_income_btc': total_income_btc,
        'total_income_usd_then': total_income_usd_then,
        'total_income_usd_now': total_income_usd_now,
        'platform_income': platform_income,
        'monthly_income_btc': monthly_income_btc,
        'monthly_income_by_platform': monthly_income_by_platform,
        'all_months': all_months,
    }


def resolve_selected_platform(platform_id):
    """Resolve a platform ID string from request.GET into a model instance or None."""
    if platform_id:
        try:
            return RemoteMiningPlatform.objects.get(pk=platform_id)
        except (RemoteMiningPlatform.DoesNotExist, ValueError):
            pass
    return None


def get_overview_data(selected_platform=None):
    """Gather all Overview dashboard data."""
    api_data = APIData.get_api_data()
    platforms = RemoteMiningPlatform.objects.all()

    # NETWORK DATA
    bitcoin_price = api_data.bitcoin_price_usd or 0
    network_hashrate = api_data.network_hashrate_ehs or 0
    network_difficulty = api_data.network_difficulty or 0
    avg_block_fees_24h = api_data.avg_block_fees_24h or 0

    # FLEET DATA
    miners = Miner.objects.select_related('platform').filter(hashrate__isnull=False, power__isnull=False)
    if selected_platform:
        miners = miners.filter(platform=selected_platform)
    miner_count = miners.count()
    total_hashrate = miners.aggregate(total=Sum('hashrate'))['total'] or 0
    total_power = miners.aggregate(total=Sum('power'))['total'] or 0
    total_capex = miners.aggregate(total=Sum('purchase_price'))['total'] or 0

    # EFFICIENCY DATA
    avg_efficiency = miners.aggregate(avg=Avg('efficiency'))['avg'] or 0
    if avg_efficiency:
        avg_efficiency = round(float(avg_efficiency), 2)

    hashrate_weighted_efficiency = 0
    if total_hashrate > 0:
        efficiency_sum = 0
        for miner in miners.filter(efficiency__isnull=False):
            efficiency_sum += float(miner.hashrate) * float(miner.efficiency)
        hashrate_weighted_efficiency = round(efficiency_sum / float(total_hashrate), 2)

    # ENERGY DATA
    miners_with_energy = miners.filter(platform__energy_price__isnull=False)
    avg_energy_cost = miners_with_energy.aggregate(avg=Avg('platform__energy_price'))['avg'] or 0
    if avg_energy_cost:
        avg_energy_cost = round(float(avg_energy_cost), 6)

    hashrate_weighted_energy_cost = 0
    if total_hashrate > 0:
        energy_cost_sum = 0
        total_hashrate_with_energy = 0
        for miner in miners_with_energy:
            energy_cost_sum += float(miner.hashrate) * float(miner.platform.energy_price)
            total_hashrate_with_energy += float(miner.hashrate)
        if total_hashrate_with_energy > 0:
            hashrate_weighted_energy_cost = round(energy_cost_sum / total_hashrate_with_energy, 6)

    # HASHRATE DISTRIBUTION DATA
    hashrate_by_platform = []
    platform_list = [selected_platform] if selected_platform else RemoteMiningPlatform.objects.all()
    for platform in platform_list:
        platform_miners = platform.miners.filter(hashrate__isnull=False, power__isnull=False)
        platform_hashrate = platform_miners.aggregate(total=Sum('hashrate'))['total'] or 0
        if platform_hashrate > 0:
            hashrate_by_platform.append({
                'platform': platform.name,
                'hashrate': float(platform_hashrate)
            })

    hashrate_by_location = []
    locations = miners.values_list('location', flat=True).distinct()
    for location in locations:
        if location:
            location_hashrate = miners.filter(location=location).aggregate(total=Sum('hashrate'))['total'] or 0
            if location_hashrate > 0:
                hashrate_by_location.append({
                    'location': location,
                    'hashrate': float(location_hashrate)
                })

    # REVENUES DATA
    payouts = Payout.objects.all()
    if selected_platform:
        payouts = payouts.filter(platform=selected_platform)
    total_btc_mined = payouts.aggregate(total=Sum('payout_amount'))['total'] or 0
    current_gross_value = float(total_btc_mined) * float(bitcoin_price) if total_btc_mined and bitcoin_price else 0
    total_payouts = payouts.count()

    gross_value_at_payout = payouts.aggregate(total=Sum('value_at_payout'))['total'] or 0
    gross_value_at_payout = float(gross_value_at_payout)

    appreciation = current_gross_value - gross_value_at_payout

    expenses = Expense.objects.filter(category='OPEX')
    if selected_platform:
        expenses = expenses.filter(platform=selected_platform)
    total_opex = expenses.aggregate(total=Sum('expense_amount'))['total'] or Decimal('0')
    total_opex = float(total_opex)

    current_net_value = current_gross_value - total_opex

    # REVENUES DISTRIBUTION DATA
    revenue_by_platform = []
    rev_platform_list = [selected_platform] if selected_platform else RemoteMiningPlatform.objects.all()
    for platform in rev_platform_list:
        platform_btc = platform.payouts.aggregate(total=Sum('payout_amount'))['total'] or 0
        platform_payouts = platform.payouts.count()
        if platform_btc > 0:
            platform_value = float(platform_btc) * float(bitcoin_price) if bitcoin_price else 0
            platform_gross_value_at_payout = platform.payouts.aggregate(total=Sum('value_at_payout'))['total'] or 0
            platform_gross_value_at_payout = float(platform_gross_value_at_payout)
            revenue_by_platform.append({
                'platform': platform.name,
                'btc_mined': float(platform_btc),
                'gross_value': platform_value,
                'gross_value_at_payout': platform_gross_value_at_payout,
                'payout_count': platform_payouts
            })

    return {
        'platforms': platforms,
        'selected_platform': selected_platform,
        'bitcoin_price': bitcoin_price,
        'network_hashrate': network_hashrate,
        'network_difficulty': network_difficulty,
        'avg_block_fees_24h': avg_block_fees_24h,
        'miner_count': miner_count,
        'total_hashrate': total_hashrate,
        'total_power': round(float(total_power), 2),
        'total_capex': total_capex,
        'avg_efficiency': avg_efficiency,
        'hashrate_weighted_efficiency': hashrate_weighted_efficiency,
        'avg_energy_cost': avg_energy_cost,
        'hashrate_weighted_energy_cost': hashrate_weighted_energy_cost,
        'hashrate_by_platform': hashrate_by_platform,
        'hashrate_by_location': hashrate_by_location,
        'total_btc_mined': total_btc_mined,
        'current_gross_value': current_gross_value,
        'gross_value_at_payout': gross_value_at_payout,
        'appreciation': appreciation,
        'total_opex': total_opex,
        'current_net_value': current_net_value,
        'total_payouts': total_payouts,
        'revenue_by_platform': revenue_by_platform,
    }


def get_forecasting_data(selected_platform=None):
    """Gather all Forecasting dashboard data."""
    api_data = APIData.get_api_data()
    settings = Settings.get_settings()
    platforms = RemoteMiningPlatform.objects.all()

    total_miner_count = Miner.objects.count()
    miners = Miner.objects.select_related('platform').filter(hashrate__isnull=False, power__isnull=False, is_active=True)
    if selected_platform:
        miners = miners.filter(platform=selected_platform)

    total_hashrate = miners.aggregate(total=Sum('hashrate'))['total'] or Decimal('0')
    miner_count = miners.count()
    total_capex = miners.aggregate(total=Sum('purchase_price'))['total'] or Decimal('0')

    hashrate_weighted_efficiency = Decimal('0')
    if total_hashrate > 0:
        total_weighted = Decimal('0')
        for miner in miners.filter(efficiency__isnull=False):
            total_weighted += miner.hashrate * miner.efficiency
        hashrate_weighted_efficiency = total_weighted / total_hashrate if total_weighted > 0 else Decimal('0')

    hashrate_weighted_energy_cost = Decimal('0')
    if total_hashrate > 0:
        total_weighted = Decimal('0')
        total_hashrate_with_energy = Decimal('0')
        for miner in miners.filter(platform__energy_price__isnull=False):
            total_weighted += miner.hashrate * miner.platform.energy_price
            total_hashrate_with_energy += miner.hashrate
        if total_hashrate_with_energy > 0:
            hashrate_weighted_energy_cost = total_weighted / total_hashrate_with_energy

    avg_efficiency = miners.filter(efficiency__isnull=False).aggregate(avg=Avg('efficiency'))['avg'] or Decimal('0')
    if avg_efficiency:
        avg_efficiency = round(float(avg_efficiency), 2)

    miners_with_energy = miners.filter(platform__energy_price__isnull=False)
    avg_energy_cost = miners_with_energy.aggregate(avg=Avg('platform__energy_price'))['avg'] or Decimal('0')
    if avg_energy_cost:
        avg_energy_cost = round(float(avg_energy_cost), 6)

    network_difficulty = api_data.network_difficulty or 0
    network_hashrate_ehs = float(api_data.network_hashrate_ehs or Decimal('0'))
    avg_tx_fees = float(api_data.avg_block_fees_24h or Decimal('0'))
    pool_fee = float(settings.pool_fee_percentage)
    btc_price_usd = float(api_data.bitcoin_price_usd or Decimal('0'))
    price_per_kwh = float(hashrate_weighted_energy_cost)
    efficiency_w_th = float(hashrate_weighted_efficiency)
    hardware_cost_usd = float(total_capex)

    # Profitability calculations
    results = None
    if total_hashrate > 0 and network_difficulty > 0 and btc_price_usd > 0:
        miner_hashrate_hs = float(total_hashrate) * 1e12

        network_hashrate_hs = network_hashrate_ehs * 1e18 if network_hashrate_ehs > 0 else 0
        hashrate_share_percent = (miner_hashrate_hs / network_hashrate_hs * 100) if network_hashrate_hs > 0 else 0

        block_reward = float(settings.block_reward) + avg_tx_fees
        expected_blocks_per_day = (miner_hashrate_hs * 86400) / (network_difficulty * 2**32)
        daily_btc_gross_before_fee = expected_blocks_per_day * block_reward
        pool_fee_btc = daily_btc_gross_before_fee * (pool_fee / 100)
        daily_btc_after_fee = daily_btc_gross_before_fee - pool_fee_btc

        miner_hashrate_ths = miner_hashrate_hs / 1e12
        power_watts = miner_hashrate_ths * efficiency_w_th
        daily_energy_kwh = (power_watts * 24) / 1000
        daily_electricity_cost_usd = daily_energy_kwh * price_per_kwh
        daily_electricity_cost_btc = daily_electricity_cost_usd / btc_price_usd if btc_price_usd > 0 else 0

        energy_cost_percentage = (daily_electricity_cost_btc / daily_btc_after_fee * 100) if daily_btc_after_fee > 0 else 0

        daily_usd_gross = daily_btc_gross_before_fee * btc_price_usd
        daily_usd_after_fee = daily_btc_after_fee * btc_price_usd
        daily_usd_net = daily_usd_after_fee - daily_electricity_cost_usd
        daily_btc_net = daily_usd_net / btc_price_usd if btc_price_usd > 0 else 0

        total_cost_usd = (pool_fee_btc * btc_price_usd) + daily_electricity_cost_usd
        cost_basis_usd_per_btc = total_cost_usd / daily_btc_after_fee if daily_btc_after_fee > 0 else 0
        discount_vs_market_pct = -1 * ((btc_price_usd - cost_basis_usd_per_btc) / btc_price_usd * 100) if btc_price_usd > 0 else 0

        margin = (daily_usd_net / daily_usd_gross * 100) if daily_usd_gross > 0 else 0

        days_to_mine_1_btc = 1 / daily_btc_after_fee if daily_btc_after_fee > 0 else float('inf')
        days_to_mine_small_btc = 0.005 / daily_btc_after_fee if daily_btc_after_fee > 0 else float('inf')

        roi_data = None
        if hardware_cost_usd > 0:
            if daily_usd_net > 0:
                days_to_roi = hardware_cost_usd / daily_usd_net
                years_roi = days_to_roi / 365
                months_roi = (years_roi - int(years_roi)) * 12
                days_roi = (months_roi - int(months_roi)) * 30
                roi_data = {
                    'days_to_roi': days_to_roi,
                    'time_breakdown': {
                        'years': int(years_roi),
                        'months': int(months_roi),
                        'days': int(days_roi)
                    }
                }
            else:
                roi_data = {
                    'days_to_roi': float('inf'),
                    'time_breakdown': {
                        'years': 0,
                        'months': 0,
                        'days': 0
                    }
                }

        def _time_breakdown(days_total):
            if days_total == float('inf'):
                return {'years': 0, 'months': 0, 'days': 0}
            return {
                'years': int(days_total / 365),
                'months': int((days_total % 365) / 30),
                'days': int(days_total % 30)
            }

        results = {
            'network_hashrate_ehs': network_hashrate_ehs,
            'hashrate_share_percent': hashrate_share_percent,
            'power_consumption_watts': power_watts,
            'power_consumption_kw': power_watts / 1000,
            'daily_energy_kwh': daily_energy_kwh,
            'energy_cost_percentage': energy_cost_percentage,
            'margin': margin,
            'days_to_mine_1_btc': days_to_mine_1_btc,
            'time_to_mine_1_btc': _time_breakdown(days_to_mine_1_btc),
            'days_to_mine_small_btc': days_to_mine_small_btc,
            'time_to_mine_small_btc': _time_breakdown(days_to_mine_small_btc),
            'roi_data': roi_data,
            'daily': {
                'btc_gross': daily_btc_gross_before_fee,
                'btc_fee': pool_fee_btc,
                'btc_after_fee': daily_btc_after_fee,
                'btc_net': daily_btc_net,
                'usd_gross': daily_usd_gross,
                'usd_fee': daily_usd_gross - daily_usd_after_fee,
                'usd_after_fee': daily_usd_after_fee,
                'usd_net': daily_usd_net,
                'electricity_cost_usd': daily_electricity_cost_usd,
                'electricity_cost_btc': daily_electricity_cost_btc
            },
            'monthly': {
                'btc_gross': daily_btc_gross_before_fee * 30,
                'btc_fee': pool_fee_btc * 30,
                'btc_after_fee': daily_btc_after_fee * 30,
                'btc_net': daily_btc_net * 30,
                'usd_gross': daily_usd_gross * 30,
                'usd_fee': (daily_usd_gross - daily_usd_after_fee) * 30,
                'usd_after_fee': daily_usd_after_fee * 30,
                'usd_net': daily_usd_net * 30,
                'electricity_cost_usd': daily_electricity_cost_usd * 30,
                'electricity_cost_btc': daily_electricity_cost_btc * 30
            },
            'yearly': {
                'btc_gross': daily_btc_gross_before_fee * 365,
                'btc_fee': pool_fee_btc * 365,
                'btc_after_fee': daily_btc_after_fee * 365,
                'btc_net': daily_btc_net * 365,
                'usd_gross': daily_usd_gross * 365,
                'usd_fee': (daily_usd_gross - daily_usd_after_fee) * 365,
                'usd_after_fee': daily_usd_after_fee * 365,
                'usd_net': daily_usd_net * 365,
                'electricity_cost_usd': daily_electricity_cost_usd * 365,
                'electricity_cost_btc': daily_electricity_cost_btc * 365
            },
            'cost_basis': {
                'cost_basis_usd_per_btc': cost_basis_usd_per_btc,
                'discount_vs_market_pct': discount_vs_market_pct
            }
        }

    return {
        'platforms': platforms,
        'selected_platform': selected_platform,
        'total_hashrate': total_hashrate,
        'miner_count': miner_count,
        'total_miner_count': total_miner_count,
        'network_difficulty': network_difficulty,
        'network_hashrate_ehs': network_hashrate_ehs,
        'avg_tx_fees': avg_tx_fees,
        'pool_fee': pool_fee,
        'btc_price_usd': btc_price_usd,
        'price_per_kwh': price_per_kwh,
        'efficiency_w_th': efficiency_w_th,
        'hardware_cost_usd': hardware_cost_usd,
        'avg_efficiency': avg_efficiency,
        'hashrate_weighted_efficiency': round(float(hashrate_weighted_efficiency), 2),
        'avg_energy_cost': avg_energy_cost,
        'hashrate_weighted_energy_cost': round(float(hashrate_weighted_energy_cost), 6),
        'results': results,
    }
