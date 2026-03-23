from django.conf import settings as django_settings
from django.shortcuts import render, redirect
from django.contrib import messages
from decimal import Decimal
from ..models import Settings, APIData
from ..forms import SettingsForm
from ..services import (
    get_capex_opex_data, get_income_data, get_overview_data,
    get_forecasting_data, resolve_selected_platform,
)


def home_view(request):
    """Home page with navigation to all sections of the application"""
    return render(request, 'mining/home.html')


# CAPEX/OPEX Dashboard View


def capex_opex_dashboard(request):
    """Dashboard view for CAPEX/OPEX analysis"""
    data = get_capex_opex_data()
    return render(request, 'mining/capex_opex_dashboard.html', data)




def income_dashboard(request):
    """Dashboard view for Income analysis"""
    data = get_income_data()
    current_btc_price = data['current_btc_price']
    platform_income = data['platform_income']
    monthly_income_btc = data['monthly_income_btc']
    monthly_income_by_platform = data['monthly_income_by_platform']
    all_months = data['all_months']

    # Prepare monthly data for template (dashboard-specific pivot)
    monthly_btc_data = []
    monthly_usd_then_data = []
    monthly_usd_now_data = []
    
    for month in all_months:
        if month:
            # BTC data
            btc_row = {'month': month, 'total': Decimal('0'), 'platforms': {}}
            usd_then_row = {'month': month, 'total': Decimal('0'), 'platforms': {}}
            usd_now_row = {'month': month, 'total': Decimal('0'), 'platforms': {}}
            
            # Find total for this month
            for item in monthly_income_btc:
                if item['month'] == month:
                    btc_row['total'] = item['total_btc']
                    usd_then_row['total'] = item['total_usd_then']
                    usd_now_row['total'] = item['total_usd_now']
                    break
            
            # Add platform data
            for platform, platform_data in monthly_income_by_platform.items():
                btc_row['platforms'][platform] = Decimal('0')
                usd_then_row['platforms'][platform] = Decimal('0')
                usd_now_row['platforms'][platform] = Decimal('0')
                
                for item in platform_data:
                    if item['month'] == month:
                        btc_row['platforms'][platform] = item['total_btc']
                        usd_then_row['platforms'][platform] = item['total_usd_then']
                        usd_now_row['platforms'][platform] = item['total_usd_now']
                        break
            
            monthly_btc_data.append(btc_row)
            monthly_usd_then_data.append(usd_then_row)
            monthly_usd_now_data.append(usd_now_row)
    
    context = {
        'total_income_btc': data['total_income_btc'],
        'total_income_usd_then': data['total_income_usd_then'],
        'total_income_usd_now': data['total_income_usd_now'],
        'platform_income': platform_income,
        'monthly_btc_data': monthly_btc_data,
        'monthly_usd_then_data': monthly_usd_then_data,
        'monthly_usd_now_data': monthly_usd_now_data,
        'platforms_with_income': [item['platform'] for item in platform_income],
        'current_btc_price': current_btc_price,
    }

    return render(request, 'mining/income_dashboard.html', context)




def overview_dashboard(request):
    """Overview Dashboard with comprehensive mining analytics"""
    selected_platform = resolve_selected_platform(request.GET.get('platform', ''))
    data = get_overview_data(selected_platform)
    return render(request, 'mining/overview_dashboard.html', data)


# Import Template Download Views


def forecasting_dashboard(request):
    """Forecasting Dashboard with BTC mining profitability calculations"""
    selected_platform = resolve_selected_platform(request.GET.get('platform', ''))
    data = get_forecasting_data(selected_platform)
    return render(request, 'mining/forecasting_dashboard.html', data)




def api_data_view(request):
    """API Data page view"""
    api_data = APIData.get_api_data()
    return render(request, 'mining/api_data.html', {'api_data': api_data})




def settings_view(request):
    """Settings page view"""
    settings = Settings.get_settings()

    if request.method == 'POST':
        form = SettingsForm(request.POST, instance=settings)
        if form.is_valid():
            form.save()
            messages.success(request, 'Settings saved successfully!')
            return redirect('settings')
    else:
        form = SettingsForm(instance=settings)

    cmc_key = django_settings.COINMARKETCAP_API_KEY

    return render(request, 'mining/settings.html', {
        'form': form,
        'settings': settings,
        'cmc_api_key': cmc_key,
    })


def app_login(request):
    """Password protection login page"""
    error = None
    if request.method == 'POST':
        password = request.POST.get('password', '')
        if password == django_settings.APP_PASSWORD:
            request.session['app_authenticated'] = True
            next_url = request.GET.get('next', '/')
            return redirect(next_url)
        error = 'Incorrect password.'

    return render(request, 'mining/login.html', {'error': error})


def app_logout(request):
    """Clear the password protection session and redirect to login"""
    request.session.flush()
    return redirect('app_login')


# Expense Views
