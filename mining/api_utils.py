import requests
import json
import math
from datetime import datetime
from .models import Settings

# (connect timeout, read timeout) in seconds
REQUEST_TIMEOUT = (10, 30)
REQUEST_HEADERS = {'User-Agent': 'MiningDashboard/1.0'}


def fetch_cmc_data(endpoint):
    """Fetch data from CoinMarketCap API"""
    settings = Settings.get_settings()
    api_key = settings.coinmarketcap_api_key

    if not api_key:
        raise ValueError("CoinMarketCap API key not configured in settings")

    url = f'https://pro-api.coinmarketcap.com/v1/{endpoint}'
    headers = {
        **REQUEST_HEADERS,
        'X-CMC_PRO_API_KEY': api_key,
    }

    response = requests.get(url, headers=headers, timeout=REQUEST_TIMEOUT)
    response.raise_for_status()

    return response.json()


def get_btc_price():
    """Get BTC price in USD from CoinMarketCap API"""
    data = fetch_cmc_data('cryptocurrency/quotes/latest?symbol=BTC&convert=USD')
    return data['data']['BTC']['quote']['USD']['price']


def get_historical_btc_price(date):
    """Get historical BTC price for a specific date using CryptoCompare API
    
    Args:
        date: datetime.date object for the date to fetch price for
        
    Returns:
        float: BTC price in USD for the specified date
    """
    # Convert date to unix timestamp
    timestamp = int(datetime.combine(date, datetime.min.time()).timestamp())
    
    # Get BTC price from CryptoCompare (free, no key needed)
    url = f"https://min-api.cryptocompare.com/data/pricehistorical?fsym=BTC&tsyms=USD&ts={timestamp}"
    
    response = requests.get(url, headers=REQUEST_HEADERS, timeout=REQUEST_TIMEOUT)
    response.raise_for_status()

    data = response.json()
    price = data['BTC']['USD']

    return float(price)


def get_bitcoin_hashrate_and_difficulty():
    """
    Fetches the current Bitcoin network hashrate and difficulty from mempool.space
    in a single API call.

    Returns:
        tuple: (hashrate_ehs, difficulty)
            - hashrate_ehs (int): Network hashrate in EH/s, rounded up.
            - difficulty (int): Current network difficulty.
    """
    url = 'https://mempool.space/api/v1/mining/hashrate/3d'

    response = requests.get(url, headers=REQUEST_HEADERS, timeout=REQUEST_TIMEOUT)
    response.raise_for_status()

    data = response.json()

    hashrate_ehs = math.ceil(data['currentHashrate'] / 1e18)
    difficulty = data['currentDifficulty']

    return hashrate_ehs, difficulty


def get_24h_avg_block_fees():
    """
    Fetches the 24h average block fees from mempool.space.
    
    Returns:
        float: The 24h average block fees in BTC, rounded to 8 decimal places.
    """
    url = 'https://mempool.space/api/v1/mining/blocks/fees/24h'

    response = requests.get(url, headers=REQUEST_HEADERS, timeout=REQUEST_TIMEOUT)
    response.raise_for_status()
    
    data = response.json()
    
    # Calculate average fees per block in BTC
    avg_btc = sum(block['avgFees'] for block in data) / len(data) / 100_000_000
    
    return round(avg_btc, 8)


def fetch_all_api_data():
    """
    Fetch all API data and return as dictionary.
    
    Returns:
        dict: Dictionary containing all API data
    """
    try:
        # Fetch CoinMarketCap data
        btc_price = get_btc_price()
        
        # Fetch mempool.space data
        hashrate, difficulty = get_bitcoin_hashrate_and_difficulty()
        avg_block_fees = get_24h_avg_block_fees()
        
        return {
            'bitcoin_price_usd': btc_price,
            'network_hashrate_ehs': hashrate,
            'network_difficulty': difficulty,
            'avg_block_fees_24h': avg_block_fees,
            'success': True,
            'message': 'API data fetched successfully!'
        }
        
    except requests.RequestException as e:
        return {
            'success': False,
            'message': f'Network error: {str(e)}'
        }
    except KeyError as e:
        return {
            'success': False,
            'message': f'API response format error: {str(e)}'
        }
    except ValueError as e:
        return {
            'success': False,
            'message': str(e)
        }
    except Exception as e:
        return {
            'success': False,
            'message': f'Unexpected error: {str(e)}'
        }
