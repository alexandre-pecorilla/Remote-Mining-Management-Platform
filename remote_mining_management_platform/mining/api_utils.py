import requests
import json
import math
from .models import Settings


def fetch_cmc_data(endpoint):
    """Fetch data from CoinMarketCap API"""
    settings = Settings.get_settings()
    api_key = settings.coinmarketcap_api_key
    
    if not api_key:
        raise ValueError("CoinMarketCap API key not configured in settings")
    
    url = f'https://pro-api.coinmarketcap.com/v1/{endpoint}'
    headers = {
        'X-CMC_PRO_API_KEY': api_key
    }
    
    response = requests.get(url, headers=headers)
    response.raise_for_status()  # Raise an exception for bad status codes
    
    return response.json()


def get_btc_price():
    """Get BTC price in USD from CoinMarketCap API"""
    data = fetch_cmc_data('cryptocurrency/quotes/latest?symbol=BTC&convert=USD')
    return data['data']['BTC']['quote']['USD']['price']


def get_bitcoin_hashrate_in_ehs():
    """
    Fetches the current Bitcoin network hashrate from mempool.space
    and returns it in exahashes per second (EH/s), rounded up to the nearest integer.
    
    Returns:
        int: The current network hashrate in EH/s, rounded up.
    """
    url = 'https://mempool.space/api/v1/mining/hashrate/3d'
    
    response = requests.get(url)
    response.raise_for_status()
    
    data = response.json()
    
    # Extract the current hashrate in EH/s and round up to the nearest integer
    hashrate_ehs = math.ceil(data['currentHashrate'] / 1e18)
    
    return hashrate_ehs


def get_bitcoin_difficulty():
    """
    Fetches the current Bitcoin network difficulty from mempool.space.
    
    Returns:
        int: The current network difficulty as an integer.
    """
    url = 'https://mempool.space/api/v1/mining/hashrate/3d'
    
    response = requests.get(url)
    response.raise_for_status()
    
    data = response.json()
    
    # Extract the current difficulty
    current_difficulty = data['currentDifficulty']
    
    return current_difficulty


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
        hashrate = get_bitcoin_hashrate_in_ehs()
        difficulty = get_bitcoin_difficulty()
        
        return {
            'bitcoin_price_usd': btc_price,
            'network_hashrate_ehs': hashrate,
            'network_difficulty': difficulty,
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
