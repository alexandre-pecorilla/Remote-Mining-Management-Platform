from django.shortcuts import render, get_object_or_404, redirect
from django.contrib import messages
from django.http import HttpResponse, JsonResponse
from django.urls import reverse_lazy, reverse
from django.views.generic import ListView, DetailView, CreateView, UpdateView, DeleteView
from django.db.models import Sum, Q
from django.db.models.functions import TruncMonth
import xlwt
import xlrd
from decimal import Decimal
from datetime import datetime
import json
from .models import RemoteMiningPlatform, Miner, Settings, APIData, Payout, Expense, TopUp
from .forms import RemoteMiningPlatformForm, MinerForm, SettingsForm, PayoutForm, ExpenseForm, TopUpForm
from .api_utils import fetch_all_api_data, get_historical_btc_price


# Home Page View
def home_view(request):
    """Home page with navigation to all sections of the application"""
    return render(request, 'mining/home.html')


# CAPEX/OPEX Dashboard View
def capex_opex_dashboard(request):
    """Dashboard view for CAPEX/OPEX analysis"""
    
    # Total Expenses calculations
    total_expenses = Expense.objects.aggregate(total=Sum('expense_amount'))['total'] or Decimal('0')
    total_capex = Expense.objects.filter(category='CAPEX').aggregate(total=Sum('expense_amount'))['total'] or Decimal('0')
    total_opex = Expense.objects.filter(category='OPEX').aggregate(total=Sum('expense_amount'))['total'] or Decimal('0')
    
    # Total Expenses by Platform
    platforms = RemoteMiningPlatform.objects.all()
    platform_expenses = []
    
    for platform in platforms:
        platform_total = Expense.objects.filter(platform=platform).aggregate(total=Sum('expense_amount'))['total'] or Decimal('0')
        platform_capex = Expense.objects.filter(platform=platform, category='CAPEX').aggregate(total=Sum('expense_amount'))['total'] or Decimal('0')
        platform_opex = Expense.objects.filter(platform=platform, category='OPEX').aggregate(total=Sum('expense_amount'))['total'] or Decimal('0')
        
        if platform_total > 0:  # Only include platforms with expenses
            platform_expenses.append({
                'platform': platform,
                'total': platform_total,
                'capex': platform_capex,
                'opex': platform_opex
            })
    
    # Monthly CAPEX calculations
    monthly_capex = Expense.objects.filter(category='CAPEX').annotate(
        month=TruncMonth('expense_date')
    ).values('month').annotate(
        total=Sum('expense_amount')
    ).order_by('month')
    
    # Monthly CAPEX by platform
    monthly_capex_by_platform = {}
    for platform in platforms:
        platform_monthly_capex = Expense.objects.filter(
            category='CAPEX', platform=platform
        ).annotate(
            month=TruncMonth('expense_date')
        ).values('month').annotate(
            total=Sum('expense_amount')
        ).order_by('month')
        
        if platform_monthly_capex:  # Only include platforms with CAPEX expenses
            monthly_capex_by_platform[platform] = platform_monthly_capex
    
    # Monthly OPEX calculations
    monthly_opex = Expense.objects.filter(category='OPEX').annotate(
        month=TruncMonth('expense_date')
    ).values('month').annotate(
        total=Sum('expense_amount')
    ).order_by('month')
    
    # Monthly OPEX by platform
    monthly_opex_by_platform = {}
    for platform in platforms:
        platform_monthly_opex = Expense.objects.filter(
            category='OPEX', platform=platform
        ).annotate(
            month=TruncMonth('expense_date')
        ).values('month').annotate(
            total=Sum('expense_amount')
        ).order_by('month')
        
        if platform_monthly_opex:  # Only include platforms with OPEX expenses
            monthly_opex_by_platform[platform] = platform_monthly_opex
    
    # Get all unique months for table structure
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
    
    context = {
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
    
    return render(request, 'mining/capex_opex_dashboard.html', context)


def export_capex_opex_data(request):
    """Export CAPEX/OPEX dashboard data to Excel file"""
    
    wb = xlwt.Workbook()
    
    # EXACT COPY of capex_opex_dashboard calculations
    # Total Expenses calculations
    total_expenses = Expense.objects.aggregate(total=Sum('expense_amount'))['total'] or Decimal('0')
    total_capex = Expense.objects.filter(category='CAPEX').aggregate(total=Sum('expense_amount'))['total'] or Decimal('0')
    total_opex = Expense.objects.filter(category='OPEX').aggregate(total=Sum('expense_amount'))['total'] or Decimal('0')
    
    # Total Expenses by Platform
    platforms = RemoteMiningPlatform.objects.all()
    platform_expenses = []
    
    for platform in platforms:
        platform_total = Expense.objects.filter(platform=platform).aggregate(total=Sum('expense_amount'))['total'] or Decimal('0')
        platform_capex = Expense.objects.filter(platform=platform, category='CAPEX').aggregate(total=Sum('expense_amount'))['total'] or Decimal('0')
        platform_opex = Expense.objects.filter(platform=platform, category='OPEX').aggregate(total=Sum('expense_amount'))['total'] or Decimal('0')
        
        if platform_total > 0:  # Only include platforms with expenses
            platform_expenses.append({
                'platform': platform,
                'total': platform_total,
                'capex': platform_capex,
                'opex': platform_opex
            })
    
    # Monthly CAPEX calculations
    monthly_capex = Expense.objects.filter(category='CAPEX').annotate(
        month=TruncMonth('expense_date')
    ).values('month').annotate(
        total=Sum('expense_amount')
    ).order_by('month')
    
    # Monthly CAPEX by platform
    monthly_capex_by_platform = {}
    for platform in platforms:
        platform_monthly_capex = Expense.objects.filter(
            category='CAPEX', platform=platform
        ).annotate(
            month=TruncMonth('expense_date')
        ).values('month').annotate(
            total=Sum('expense_amount')
        ).order_by('month')
        
        if platform_monthly_capex:  # Only include platforms with CAPEX expenses
            monthly_capex_by_platform[platform] = platform_monthly_capex
    
    # Monthly OPEX calculations
    monthly_opex = Expense.objects.filter(category='OPEX').annotate(
        month=TruncMonth('expense_date')
    ).values('month').annotate(
        total=Sum('expense_amount')
    ).order_by('month')
    
    # Monthly OPEX by platform
    monthly_opex_by_platform = {}
    for platform in platforms:
        platform_monthly_opex = Expense.objects.filter(
            category='OPEX', platform=platform
        ).annotate(
            month=TruncMonth('expense_date')
        ).values('month').annotate(
            total=Sum('expense_amount')
        ).order_by('month')
        
        if platform_monthly_opex:  # Only include platforms with OPEX expenses
            monthly_opex_by_platform[platform] = platform_monthly_opex
    
    # Get all unique months for table structure
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
    
    # Sheet 1: Total Expenses Summary
    ws_summary = wb.add_sheet('Total Expenses Summary')
    
    # Headers
    ws_summary.write(0, 0, 'Expense Type')
    ws_summary.write(0, 1, 'Amount (USD)')
    
    # Data rows
    ws_summary.write(1, 0, 'Total Expenses')
    ws_summary.write(1, 1, float(total_expenses))
    
    ws_summary.write(2, 0, 'Total CAPEX')
    ws_summary.write(2, 1, float(total_capex))
    
    ws_summary.write(3, 0, 'Total OPEX')
    ws_summary.write(3, 1, float(total_opex))
    
    # Sheet 2: Expenses by Platform
    if platform_expenses:
        ws_platform = wb.add_sheet('Expenses by Platform')
        
        # Headers
        ws_platform.write(0, 0, 'Platform')
        ws_platform.write(0, 1, 'Total Expenses (USD)')
        ws_platform.write(0, 2, 'CAPEX (USD)')
        ws_platform.write(0, 3, 'OPEX (USD)')
        
        # Data rows
        for row, item in enumerate(platform_expenses, start=1):
            ws_platform.write(row, 0, item['platform'].name)
            ws_platform.write(row, 1, float(item['total']))
            ws_platform.write(row, 2, float(item['capex']))
            ws_platform.write(row, 3, float(item['opex']))
    
    # Sheet 3: Monthly CAPEX
    if monthly_capex and all_months:
        ws_monthly_capex = wb.add_sheet('Monthly CAPEX')
        
        # Headers
        ws_monthly_capex.write(0, 0, 'Month')
        ws_monthly_capex.write(0, 1, 'Total CAPEX (USD)')
        
        # Platform headers
        col = 2
        platform_cols = {}
        for platform in monthly_capex_by_platform.keys():
            ws_monthly_capex.write(0, col, f'{platform.name} CAPEX (USD)')
            platform_cols[platform] = col
            col += 1
        
        # Data rows
        for row, month in enumerate(all_months, start=1):
            if month:
                ws_monthly_capex.write(row, 0, month.strftime('%Y-%m'))
                
                # Total CAPEX for this month
                month_total = Decimal('0')
                for item in monthly_capex:
                    if item['month'] == month:
                        month_total = item['total']
                        break
                ws_monthly_capex.write(row, 1, float(month_total))
                
                # Platform CAPEX for this month
                for platform, platform_data in monthly_capex_by_platform.items():
                    platform_month_total = Decimal('0')
                    for item in platform_data:
                        if item['month'] == month:
                            platform_month_total = item['total']
                            break
                    ws_monthly_capex.write(row, platform_cols[platform], float(platform_month_total))
    
    # Sheet 4: Monthly OPEX
    if monthly_opex and all_months:
        ws_monthly_opex = wb.add_sheet('Monthly OPEX')
        
        # Headers
        ws_monthly_opex.write(0, 0, 'Month')
        ws_monthly_opex.write(0, 1, 'Total OPEX (USD)')
        
        # Platform headers
        col = 2
        platform_cols = {}
        for platform in monthly_opex_by_platform.keys():
            ws_monthly_opex.write(0, col, f'{platform.name} OPEX (USD)')
            platform_cols[platform] = col
            col += 1
        
        # Data rows
        for row, month in enumerate(all_months, start=1):
            if month:
                ws_monthly_opex.write(row, 0, month.strftime('%Y-%m'))
                
                # Total OPEX for this month
                month_total = Decimal('0')
                for item in monthly_opex:
                    if item['month'] == month:
                        month_total = item['total']
                        break
                ws_monthly_opex.write(row, 1, float(month_total))
                
                # Platform OPEX for this month
                for platform, platform_data in monthly_opex_by_platform.items():
                    platform_month_total = Decimal('0')
                    for item in platform_data:
                        if item['month'] == month:
                            platform_month_total = item['total']
                            break
                    ws_monthly_opex.write(row, platform_cols[platform], float(platform_month_total))
    
    # Generate response
    response = HttpResponse(content_type='application/vnd.ms-excel')
    response['Content-Disposition'] = f'attachment; filename="capex_opex_dashboard_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xls"'
    
    wb.save(response)
    return response


# Income Dashboard View
def income_dashboard(request):
    """Dashboard view for Income analysis"""
    
    # Get API data for current market value calculations
    api_data = APIData.get_api_data()
    current_btc_price = float(api_data.bitcoin_price_usd) if api_data.bitcoin_price_usd else 0
    
    # Total Income calculations
    total_income_btc = Payout.objects.aggregate(total=Sum('payout_amount'))['total'] or Decimal('0')
    total_income_usd_then = Payout.objects.aggregate(total=Sum('value_at_payout'))['total'] or Decimal('0')
    
    # Calculate total income USD now (current market value)
    total_income_usd_now = Decimal('0')
    if current_btc_price > 0:
        total_income_usd_now = total_income_btc * Decimal(str(current_btc_price))
    
    # Total Income by Platform
    platforms = RemoteMiningPlatform.objects.all()
    platform_income = []
    
    for platform in platforms:
        platform_btc = Payout.objects.filter(platform=platform).aggregate(total=Sum('payout_amount'))['total'] or Decimal('0')
        platform_usd_then = Payout.objects.filter(platform=platform).aggregate(total=Sum('value_at_payout'))['total'] or Decimal('0')
        platform_usd_now = Decimal('0')
        if current_btc_price > 0:
            platform_usd_now = platform_btc * Decimal(str(current_btc_price))
        
        if platform_btc > 0:  # Only include platforms with income
            platform_income.append({
                'platform': platform,
                'total_btc': platform_btc,
                'total_usd_then': platform_usd_then,
                'total_usd_now': platform_usd_now
            })
    
    # Monthly Income BTC calculations
    monthly_income_btc = Payout.objects.annotate(
        month=TruncMonth('payout_date')
    ).values('month').annotate(
        total_btc=Sum('payout_amount'),
        total_usd_then=Sum('value_at_payout')
    ).order_by('month')
    
    # Add current market value to monthly income
    for item in monthly_income_btc:
        item['total_usd_now'] = item['total_btc'] * Decimal(str(current_btc_price)) if current_btc_price > 0 else Decimal('0')
    
    # Monthly Income by platform
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
        
        # Add current market value
        for item in platform_monthly_income:
            item['total_usd_now'] = item['total_btc'] * Decimal(str(current_btc_price)) if current_btc_price > 0 else Decimal('0')
        
        if platform_monthly_income:  # Only include platforms with income
            monthly_income_by_platform[platform] = platform_monthly_income
    
    # Get all unique months for table structure
    all_months = set()
    for item in monthly_income_btc:
        all_months.add(item['month'])
    for platform_data in monthly_income_by_platform.values():
        for item in platform_data:
            all_months.add(item['month'])
    
    all_months = sorted(list(all_months))
    
    # Prepare monthly data for template
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
        'total_income_btc': total_income_btc,
        'total_income_usd_then': total_income_usd_then,
        'total_income_usd_now': total_income_usd_now,
        'platform_income': platform_income,
        'monthly_btc_data': monthly_btc_data,
        'monthly_usd_then_data': monthly_usd_then_data,
        'monthly_usd_now_data': monthly_usd_now_data,
        'platforms_with_income': [item['platform'] for item in platform_income],
        'current_btc_price': current_btc_price,
    }
    
    return render(request, 'mining/income_dashboard.html', context)


def export_income_data(request):
    """Export Income dashboard data to Excel file"""
    
    wb = xlwt.Workbook()
    
    # Get API data for current market value calculations
    api_data = APIData.get_api_data()
    current_btc_price = float(api_data.bitcoin_price_usd) if api_data.bitcoin_price_usd else 0
    
    # EXACT COPY of income_dashboard calculations
    # Total Income calculations
    total_income_btc = Payout.objects.aggregate(total=Sum('payout_amount'))['total'] or Decimal('0')
    total_income_usd_then = Payout.objects.aggregate(total=Sum('value_at_payout'))['total'] or Decimal('0')
    
    # Calculate total income USD now (current market value)
    total_income_usd_now = Decimal('0')
    if current_btc_price > 0:
        total_income_usd_now = total_income_btc * Decimal(str(current_btc_price))
    
    # Total Income by Platform
    platforms = RemoteMiningPlatform.objects.all()
    platform_income = []
    
    for platform in platforms:
        platform_btc = Payout.objects.filter(platform=platform).aggregate(total=Sum('payout_amount'))['total'] or Decimal('0')
        platform_usd_then = Payout.objects.filter(platform=platform).aggregate(total=Sum('value_at_payout'))['total'] or Decimal('0')
        platform_usd_now = Decimal('0')
        if current_btc_price > 0:
            platform_usd_now = platform_btc * Decimal(str(current_btc_price))
        
        if platform_btc > 0:  # Only include platforms with income
            platform_income.append({
                'platform': platform,
                'total_btc': platform_btc,
                'total_usd_then': platform_usd_then,
                'total_usd_now': platform_usd_now
            })
    
    # Monthly Income calculations
    monthly_income_btc = Payout.objects.annotate(
        month=TruncMonth('payout_date')
    ).values('month').annotate(
        total_btc=Sum('payout_amount'),
        total_usd_then=Sum('value_at_payout')
    ).order_by('month')
    
    # Add current market value to monthly income
    for item in monthly_income_btc:
        item['total_usd_now'] = item['total_btc'] * Decimal(str(current_btc_price)) if current_btc_price > 0 else Decimal('0')
    
    # Monthly Income by platform
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
        
        # Add current market value
        for item in platform_monthly_income:
            item['total_usd_now'] = item['total_btc'] * Decimal(str(current_btc_price)) if current_btc_price > 0 else Decimal('0')
        
        if platform_monthly_income:  # Only include platforms with income
            monthly_income_by_platform[platform] = platform_monthly_income
    
    # Get all unique months for table structure
    all_months = set()
    for item in monthly_income_btc:
        all_months.add(item['month'])
    for platform_data in monthly_income_by_platform.values():
        for item in platform_data:
            all_months.add(item['month'])
    
    all_months = sorted(list(all_months))
    
    # Sheet 1: Total Income Summary
    ws_summary = wb.add_sheet('Total Income Summary')
    
    # Headers
    ws_summary.write(0, 0, 'Income Type')
    ws_summary.write(0, 1, 'Amount')
    
    # Data rows
    ws_summary.write(1, 0, 'Total Income BTC')
    ws_summary.write(1, 1, float(total_income_btc))
    
    ws_summary.write(2, 0, 'Total Income USD (then)')
    ws_summary.write(2, 1, float(total_income_usd_then))
    
    ws_summary.write(3, 0, 'Total Income USD (now)')
    ws_summary.write(3, 1, float(total_income_usd_now))
    
    # Sheet 2: Income by Platform
    if platform_income:
        ws_platform = wb.add_sheet('Income by Platform')
        
        # Headers
        ws_platform.write(0, 0, 'Platform')
        ws_platform.write(0, 1, 'Total Income BTC')
        ws_platform.write(0, 2, 'Total Income USD (then)')
        ws_platform.write(0, 3, 'Total Income USD (now)')
        
        # Data rows
        for row, item in enumerate(platform_income, start=1):
            ws_platform.write(row, 0, item['platform'].name)
            ws_platform.write(row, 1, float(item['total_btc']))
            ws_platform.write(row, 2, float(item['total_usd_then']))
            ws_platform.write(row, 3, float(item['total_usd_now']))
    
    # Sheet 3: Monthly Income BTC
    if monthly_income_btc and all_months:
        ws_monthly_btc = wb.add_sheet('Monthly Income BTC')
        
        # Headers
        ws_monthly_btc.write(0, 0, 'Month')
        ws_monthly_btc.write(0, 1, 'Total Income BTC')
        
        # Platform headers
        col = 2
        platform_cols = {}
        for platform in monthly_income_by_platform.keys():
            ws_monthly_btc.write(0, col, f'{platform.name} BTC')
            platform_cols[platform] = col
            col += 1
        
        # Data rows
        for row, month in enumerate(all_months, start=1):
            if month:
                ws_monthly_btc.write(row, 0, month.strftime('%Y-%m'))
                
                # Total BTC for this month
                month_total = Decimal('0')
                for item in monthly_income_btc:
                    if item['month'] == month:
                        month_total = item['total_btc']
                        break
                ws_monthly_btc.write(row, 1, float(month_total))
                
                # Platform BTC for this month
                for platform, platform_data in monthly_income_by_platform.items():
                    platform_month_total = Decimal('0')
                    for item in platform_data:
                        if item['month'] == month:
                            platform_month_total = item['total_btc']
                            break
                    ws_monthly_btc.write(row, platform_cols[platform], float(platform_month_total))
    
    # Sheet 4: Monthly Income USD (then)
    if monthly_income_btc and all_months:
        ws_monthly_usd_then = wb.add_sheet('Monthly Income USD then')
        
        # Headers
        ws_monthly_usd_then.write(0, 0, 'Month')
        ws_monthly_usd_then.write(0, 1, 'Total Income USD (then)')
        
        # Platform headers
        col = 2
        platform_cols = {}
        for platform in monthly_income_by_platform.keys():
            ws_monthly_usd_then.write(0, col, f'{platform.name} USD (then)')
            platform_cols[platform] = col
            col += 1
        
        # Data rows
        for row, month in enumerate(all_months, start=1):
            if month:
                ws_monthly_usd_then.write(row, 0, month.strftime('%Y-%m'))
                
                # Total USD then for this month
                month_total = Decimal('0')
                for item in monthly_income_btc:
                    if item['month'] == month:
                        month_total = item['total_usd_then']
                        break
                ws_monthly_usd_then.write(row, 1, float(month_total))
                
                # Platform USD then for this month
                for platform, platform_data in monthly_income_by_platform.items():
                    platform_month_total = Decimal('0')
                    for item in platform_data:
                        if item['month'] == month:
                            platform_month_total = item['total_usd_then']
                            break
                    ws_monthly_usd_then.write(row, platform_cols[platform], float(platform_month_total))
    
    # Sheet 5: Monthly Income USD (now)
    if monthly_income_btc and all_months:
        ws_monthly_usd_now = wb.add_sheet('Monthly Income USD now')
        
        # Headers
        ws_monthly_usd_now.write(0, 0, 'Month')
        ws_monthly_usd_now.write(0, 1, 'Total Income USD (now)')
        
        # Platform headers
        col = 2
        platform_cols = {}
        for platform in monthly_income_by_platform.keys():
            ws_monthly_usd_now.write(0, col, f'{platform.name} USD (now)')
            platform_cols[platform] = col
            col += 1
        
        # Data rows
        for row, month in enumerate(all_months, start=1):
            if month:
                ws_monthly_usd_now.write(row, 0, month.strftime('%Y-%m'))
                
                # Total USD now for this month
                month_total = Decimal('0')
                for item in monthly_income_btc:
                    if item['month'] == month:
                        month_total = item['total_usd_now']
                        break
                ws_monthly_usd_now.write(row, 1, float(month_total))
                
                # Platform USD now for this month
                for platform, platform_data in monthly_income_by_platform.items():
                    platform_month_total = Decimal('0')
                    for item in platform_data:
                        if item['month'] == month:
                            platform_month_total = item['total_usd_now']
                            break
                    ws_monthly_usd_now.write(row, platform_cols[platform], float(platform_month_total))
    
    # Generate response
    response = HttpResponse(content_type='application/vnd.ms-excel')
    response['Content-Disposition'] = f'attachment; filename="income_dashboard_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xls"'
    
    wb.save(response)
    return response


class PlatformListView(ListView):
    model = RemoteMiningPlatform
    template_name = 'mining/platform_list.html'
    context_object_name = 'platforms'
    paginate_by = 10


class PlatformDetailView(DetailView):
    model = RemoteMiningPlatform
    template_name = 'mining/platform_detail.html'
    context_object_name = 'platform'
    
    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        current_platform = self.get_object()
        
        # Get previous platform (lower ID)
        previous_platform = RemoteMiningPlatform.objects.filter(
            id__lt=current_platform.id
        ).order_by('-id').first()
        
        # Get next platform (higher ID)
        next_platform = RemoteMiningPlatform.objects.filter(
            id__gt=current_platform.id
        ).order_by('id').first()
        
        context['previous_platform'] = previous_platform
        context['next_platform'] = next_platform
        return context


class PlatformCreateView(CreateView):
    model = RemoteMiningPlatform
    form_class = RemoteMiningPlatformForm
    template_name = 'mining/platform_form.html'
    
    def get_success_url(self):
        return reverse_lazy('platform_detail', kwargs={'pk': self.object.pk})
    
    def form_valid(self, form):
        messages.success(self.request, 'Platform created successfully.')
        return super().form_valid(form)


class PlatformUpdateView(UpdateView):
    model = RemoteMiningPlatform
    form_class = RemoteMiningPlatformForm
    template_name = 'mining/platform_form.html'
    
    def get_success_url(self):
        return reverse_lazy('platform_detail', kwargs={'pk': self.object.pk})
    
    def form_valid(self, form):
        messages.success(self.request, 'Platform updated successfully.')
        return super().form_valid(form)


class PlatformDeleteView(DeleteView):
    model = RemoteMiningPlatform
    template_name = 'mining/platform_confirm_delete.html'
    success_url = reverse_lazy('platform_list')
    
    def delete(self, request, *args, **kwargs):
        messages.success(request, "Platform deleted successfully!")
        return super().delete(request, *args, **kwargs)


# Miner Views
class MinerListView(ListView):
    model = Miner
    template_name = 'mining/miner_list.html'
    context_object_name = 'miners'
    paginate_by = 12


class MinerDetailView(DetailView):
    model = Miner
    template_name = 'mining/miner_detail.html'
    context_object_name = 'miner'
    
    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        current_miner = self.get_object()
        
        # Get previous miner (lower ID)
        previous_miner = Miner.objects.filter(
            id__lt=current_miner.id
        ).order_by('-id').first()
        
        # Get next miner (higher ID)
        next_miner = Miner.objects.filter(
            id__gt=current_miner.id
        ).order_by('id').first()
        
        context['previous_miner'] = previous_miner
        context['next_miner'] = next_miner
        return context


class MinerCreateView(CreateView):
    model = Miner
    form_class = MinerForm
    template_name = 'mining/miner_form.html'
    
    def get_success_url(self):
        return reverse_lazy('miner_detail', kwargs={'pk': self.object.pk})

    def form_valid(self, form):
        messages.success(self.request, "Miner created successfully!")
        return super().form_valid(form)


class MinerUpdateView(UpdateView):
    model = Miner
    form_class = MinerForm
    template_name = 'mining/miner_form.html'
    
    def get_success_url(self):
        return reverse_lazy('miner_detail', kwargs={'pk': self.object.pk})

    def form_valid(self, form):
        messages.success(self.request, "Miner updated successfully!")
        return super().form_valid(form)


class MinerDeleteView(DeleteView):
    model = Miner
    template_name = 'mining/miner_confirm_delete.html'
    success_url = reverse_lazy('miner_list')
    context_object_name = 'miner'

    def delete(self, request, *args, **kwargs):
        messages.success(request, "Miner deleted successfully!")
        return super().delete(request, *args, **kwargs)


def toggle_miner_active(request, pk):
    """Toggle a miner's is_active status (on/off)"""
    if request.method == 'POST':
        miner = get_object_or_404(Miner, pk=pk)
        miner.is_active = not miner.is_active
        miner.save(update_fields=['is_active'])
        status = "ON" if miner.is_active else "OFF"
        messages.success(request, f"{miner.model} turned {status}.")
    return redirect('miner_detail', pk=pk)


# Payout Views
class PayoutListView(ListView):
    model = Payout
    template_name = 'mining/payout_list.html'
    context_object_name = 'payouts'
    paginate_by = 50


class PayoutDetailView(DetailView):
    model = Payout
    template_name = 'mining/payout_detail.html'
    context_object_name = 'payout'
    
    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        current_payout = self.get_object()
        
        # Get previous payout (lower ID)
        previous_payout = Payout.objects.filter(
            id__lt=current_payout.id
        ).order_by('-id').first()
        
        # Get next payout (higher ID)
        next_payout = Payout.objects.filter(
            id__gt=current_payout.id
        ).order_by('id').first()
        
        context['previous_payout'] = previous_payout
        context['next_payout'] = next_payout
        return context


class PayoutCreateView(CreateView):
    model = Payout
    form_class = PayoutForm
    template_name = 'mining/payout_form.html'
    
    def get_success_url(self):
        return reverse_lazy('payout_detail', kwargs={'pk': self.object.pk})
    
    def form_valid(self, form):
        messages.success(self.request, 'Payout added successfully!')
        return super().form_valid(form)


class PayoutUpdateView(UpdateView):
    model = Payout
    form_class = PayoutForm
    template_name = 'mining/payout_form.html'
    
    def get_success_url(self):
        return reverse_lazy('payout_detail', kwargs={'pk': self.object.pk})
    
    def form_valid(self, form):
        messages.success(self.request, 'Payout updated successfully!')
        return super().form_valid(form)


class PayoutDeleteView(DeleteView):
    model = Payout
    template_name = 'mining/payout_confirm_delete.html'
    success_url = reverse_lazy('payout_list')
    context_object_name = 'payout'
    
    def delete(self, request, *args, **kwargs):
        messages.success(self.request, 'Payout deleted successfully!')
        return super().delete(request, *args, **kwargs)


def fetch_closing_price(request, payout_id):
    """Fetch historical BTC price for payout date and update closing_price field"""
    if request.method == 'POST':
        try:
            payout = get_object_or_404(Payout, pk=payout_id)
            
            if not payout.payout_date:
                return JsonResponse({
                    'success': False,
                    'error': 'Payout date is required to fetch closing price'
                })
            
            # Fetch historical BTC price for the payout date
            historical_price = get_historical_btc_price(payout.payout_date)
            
            # Update the payout's closing_price field
            payout.closing_price = Decimal(str(historical_price))
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
    
    return JsonResponse({'success': False, 'error': 'Invalid request method'})


def api_data_view(request):
    """API Data page view"""
    api_data = APIData.get_api_data()
    
    if request.method == 'POST':
        # Fetch API data when button is clicked
        result = fetch_all_api_data()
        
        if result['success']:
            # Update the API data in database
            api_data.bitcoin_price_usd = result['bitcoin_price_usd']
            api_data.network_hashrate_ehs = result['network_hashrate_ehs']
            api_data.network_difficulty = result['network_difficulty']
            api_data.avg_block_fees_24h = result['avg_block_fees_24h']
            api_data.save()
            
            messages.success(request, result['message'])
        else:
            messages.error(request, result['message'])
            
        return redirect('api_data')
    
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
    
    return render(request, 'mining/settings.html', {'form': form, 'settings': settings})


# Expense Views
class ExpenseListView(ListView):
    model = Expense
    template_name = 'mining/expense_list.html'
    context_object_name = 'expenses'
    paginate_by = 50


class ExpenseDetailView(DetailView):
    model = Expense
    template_name = 'mining/expense_detail.html'
    context_object_name = 'expense'
    
    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        current_expense = self.get_object()
        
        # Get previous expense (lower ID)
        previous_expense = Expense.objects.filter(
            id__lt=current_expense.id
        ).order_by('-id').first()
        
        # Get next expense (higher ID)
        next_expense = Expense.objects.filter(
            id__gt=current_expense.id
        ).order_by('id').first()
        
        context['previous_expense'] = previous_expense
        context['next_expense'] = next_expense
        return context


class ExpenseCreateView(CreateView):
    model = Expense
    form_class = ExpenseForm
    template_name = 'mining/expense_form.html'
    
    def get_success_url(self):
        return reverse_lazy('expense_detail', kwargs={'pk': self.object.pk})
    
    def form_valid(self, form):
        messages.success(self.request, 'Expense created successfully!')
        return super().form_valid(form)


class ExpenseUpdateView(UpdateView):
    model = Expense
    form_class = ExpenseForm
    template_name = 'mining/expense_form.html'
    
    def get_success_url(self):
        return reverse_lazy('expense_detail', kwargs={'pk': self.object.pk})
    
    def form_valid(self, form):
        messages.success(self.request, 'Expense updated successfully!')
        return super().form_valid(form)


class ExpenseDeleteView(DeleteView):
    model = Expense
    template_name = 'mining/expense_confirm_delete.html'
    success_url = reverse_lazy('expense_list')
    context_object_name = 'expense'
    
    def delete(self, request, *args, **kwargs):
        messages.success(self.request, 'Expense deleted successfully!')
        return super().delete(request, *args, **kwargs)


# Dashboard Views
def overview_dashboard(request):
    """Overview Dashboard with comprehensive mining analytics"""
    from django.db.models import Sum, Avg, Count
    from decimal import Decimal
    
    # Get API data
    api_data = APIData.get_api_data()
    
    # Platform filter
    platforms = RemoteMiningPlatform.objects.all()
    selected_platform_id = request.GET.get('platform', '')
    selected_platform = None
    if selected_platform_id:
        try:
            selected_platform = RemoteMiningPlatform.objects.get(pk=selected_platform_id)
        except (RemoteMiningPlatform.DoesNotExist, ValueError):
            selected_platform = None
    
    # NETWORK DATA
    bitcoin_price = api_data.bitcoin_price_usd or 0
    network_hashrate = api_data.network_hashrate_ehs or 0
    network_difficulty = api_data.network_difficulty or 0
    avg_block_fees_24h = api_data.avg_block_fees_24h or 0
    
    # FLEET DATA
    miners = Miner.objects.filter(hashrate__isnull=False, power__isnull=False)
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
    
    # Hashrate weighted average efficiency
    hashrate_weighted_efficiency = 0
    if total_hashrate > 0:
        efficiency_sum = 0
        for miner in miners.filter(efficiency__isnull=False):
            efficiency_sum += float(miner.hashrate) * float(miner.efficiency)
        hashrate_weighted_efficiency = round(efficiency_sum / float(total_hashrate), 2)
    
    # ENERGY DATA
    # Get miners with platforms that have energy prices
    miners_with_energy = miners.filter(platform__energy_price__isnull=False)
    avg_energy_cost = miners_with_energy.aggregate(avg=Avg('platform__energy_price'))['avg'] or 0
    if avg_energy_cost:
        avg_energy_cost = round(float(avg_energy_cost), 6)
    
    # Hashrate weighted average energy cost
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
    
    # Hashrate by location
    hashrate_by_location = []
    locations = miners.values_list('location', flat=True).distinct()
    for location in locations:
        if location:  # Skip empty locations
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
    
    # Calculate Gross Value at Payout (sum of value_at_payout field)
    gross_value_at_payout = payouts.aggregate(total=Sum('value_at_payout'))['total'] or 0
    gross_value_at_payout = float(gross_value_at_payout)
    
    # Calculate Appreciation (Current Gross Value - Gross Value at Payout)
    appreciation = current_gross_value - gross_value_at_payout
    
    # Calculate Total OPEX (sum of all OPEX expenses)
    expenses = Expense.objects.filter(category='OPEX')
    if selected_platform:
        expenses = expenses.filter(platform=selected_platform)
    total_opex = expenses.aggregate(total=Sum('expense_amount'))['total'] or Decimal('0')
    total_opex = float(total_opex)
    
    # Calculate Current Net Value (Current Gross Value - Total OPEX)
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
    
    context = {
        # Platform filter
        'platforms': platforms,
        'selected_platform': selected_platform,
        
        # Network Data
        'bitcoin_price': bitcoin_price,
        'network_hashrate': network_hashrate,
        'network_difficulty': network_difficulty,
        'avg_block_fees_24h': avg_block_fees_24h,
        
        # Fleet Data
        'miner_count': miner_count,
        'total_hashrate': total_hashrate,
        'total_power': round(float(total_power), 2),  # Power already stored in kW in database
        'total_capex': total_capex,
        
        # Efficiency Data
        'avg_efficiency': avg_efficiency,
        'hashrate_weighted_efficiency': hashrate_weighted_efficiency,
        
        # Energy Data
        'avg_energy_cost': avg_energy_cost,
        'hashrate_weighted_energy_cost': hashrate_weighted_energy_cost,
        
        # Hashrate Distribution
        'hashrate_by_platform': hashrate_by_platform,
        'hashrate_by_location': hashrate_by_location,
        
        # Revenues Data
        'total_btc_mined': total_btc_mined,
        'current_gross_value': current_gross_value,
        'gross_value_at_payout': gross_value_at_payout,
        'appreciation': appreciation,
        'total_opex': total_opex,
        'current_net_value': current_net_value,
        'total_payouts': total_payouts,
        
        # Revenue Distribution
        'revenue_by_platform': revenue_by_platform,
    }
    
    return render(request, 'mining/overview_dashboard.html', context)


# Import Template Download Views
def download_platform_template(request):
    """Download import template for Remote Mining Platforms"""
    wb = xlwt.Workbook()
    ws = wb.add_sheet('Platform Import Template')
    
    # Add headers based on form fields
    headers = ['name', 'website_link', 'portal_url', 'point_of_contact_name', 
               'point_of_contact_email', 'point_of_contact_phone', 'point_of_contact_telegram', 'energy_price']
    
    for col, header in enumerate(headers):
        ws.write(0, col, header)
    
    response = HttpResponse(
        content_type='application/vnd.ms-excel'
    )
    response['Content-Disposition'] = 'attachment; filename="platform_import_template.xls"'
    wb.save(response)
    return response


def download_miner_template(request):
    """Download import template for Miners"""
    wb = xlwt.Workbook()
    ws = wb.add_sheet('Miner Import Template')
    
    # Add headers based on form fields (excluding image field for import)
    headers = ['model', 'manufacturer', 'product_link', 'serial_number', 
               'platform', 'platform_internal_id', 'hashrate', 'power', 'efficiency', 
               'purchase_price', 'purchase_date', 'start_date', 'location']
    
    for col, header in enumerate(headers):
        ws.write(0, col, header)
    
    response = HttpResponse(
        content_type='application/vnd.ms-excel'
    )
    response['Content-Disposition'] = 'attachment; filename="miner_import_template.xls"'
    wb.save(response)
    return response


def download_payout_template(request):
    """Download import template for Payouts"""
    wb = xlwt.Workbook()
    ws = wb.add_sheet('Payout Import Template')
    
    # Add headers based on form fields
    headers = ['payout_date', 'payout_amount', 'platform', 'transaction_id', 'closing_price', 'value_at_payout (read-only)']
    
    for col, header in enumerate(headers):
        ws.write(0, col, header)
    
    response = HttpResponse(
        content_type='application/vnd.ms-excel'
    )
    response['Content-Disposition'] = 'attachment; filename="payout_import_template.xls"'
    wb.save(response)
    return response


# Data Export Views
def export_platform_data(request):
    """Export all platform data to Excel file"""
    wb = xlwt.Workbook()
    ws = wb.add_sheet('Platform Data')
    
    # Add headers
    headers = ['name', 'website_link', 'portal_url', 'point_of_contact_name', 
               'point_of_contact_email', 'point_of_contact_phone', 
               'point_of_contact_telegram', 'energy_price']
    
    for col, header in enumerate(headers):
        ws.write(0, col, header)
    
    # Add data rows
    platforms = RemoteMiningPlatform.objects.all()
    for row, platform in enumerate(platforms, start=1):
        ws.write(row, 0, platform.name or '')
        ws.write(row, 1, platform.website_link or '')
        ws.write(row, 2, platform.portal_url or '')
        ws.write(row, 3, platform.point_of_contact_name or '')
        ws.write(row, 4, platform.point_of_contact_email or '')
        ws.write(row, 5, platform.point_of_contact_phone or '')
        ws.write(row, 6, platform.point_of_contact_telegram or '')
        ws.write(row, 7, float(platform.energy_price) if platform.energy_price else '')
    
    response = HttpResponse(
        content_type='application/vnd.ms-excel'
    )
    response['Content-Disposition'] = 'attachment; filename="platform_data_export.xls"'
    wb.save(response)
    return response


def export_miner_data(request):
    """Export all miner data to Excel file"""
    wb = xlwt.Workbook()
    ws = wb.add_sheet('Miner Data')
    
    # Add headers
    headers = ['model', 'manufacturer', 'product_link', 'serial_number', 
               'platform', 'platform_internal_id', 'hashrate', 'power', 'efficiency', 
               'purchase_price', 'purchase_date', 'start_date', 'location']
    
    for col, header in enumerate(headers):
        ws.write(0, col, header)
    
    # Add data rows
    miners = Miner.objects.all()
    for row, miner in enumerate(miners, start=1):
        ws.write(row, 0, miner.model or '')
        ws.write(row, 1, miner.manufacturer or '')
        ws.write(row, 2, miner.product_link or '')
        ws.write(row, 3, miner.serial_number or '')
        ws.write(row, 4, miner.platform.pk if miner.platform else '')
        ws.write(row, 5, miner.platform_internal_id or '')
        ws.write(row, 6, float(miner.hashrate) if miner.hashrate else '')
        ws.write(row, 7, float(miner.power) if miner.power else '')
        ws.write(row, 8, float(miner.efficiency) if miner.efficiency else '')
        ws.write(row, 9, float(miner.purchase_price) if miner.purchase_price else '')
        ws.write(row, 10, miner.purchase_date.strftime('%Y-%m-%d') if miner.purchase_date else '')
        ws.write(row, 11, miner.start_date.strftime('%Y-%m-%d') if miner.start_date else '')
        ws.write(row, 12, miner.location or '')
    
    response = HttpResponse(
        content_type='application/vnd.ms-excel'
    )
    response['Content-Disposition'] = 'attachment; filename="miner_data_export.xls"'
    wb.save(response)
    return response


def export_payout_data(request):
    """Export all payout data to Excel file"""
    wb = xlwt.Workbook()
    ws = wb.add_sheet('Payout Data')
    
    # Add headers
    headers = ['payout_date', 'payout_amount', 'platform', 'transaction_id', 'closing_price', 'value_at_payout']
    
    for col, header in enumerate(headers):
        ws.write(0, col, header)
    
    # Add data rows
    payouts = Payout.objects.all()
    for row, payout in enumerate(payouts, start=1):
        ws.write(row, 0, payout.payout_date.strftime('%Y-%m-%d'))
        ws.write(row, 1, float(payout.payout_amount))
        ws.write(row, 2, payout.platform.pk if payout.platform else '')
        ws.write(row, 3, payout.transaction_id or '')
        ws.write(row, 4, float(payout.closing_price) if payout.closing_price else '')
        ws.write(row, 5, float(payout.value_at_payout) if payout.value_at_payout else '')
    
    response = HttpResponse(
        content_type='application/vnd.ms-excel'
    )
    response['Content-Disposition'] = 'attachment; filename="payout_data_export.xls"'
    wb.save(response)
    return response


def export_overview_data(request):
    """Export overview dashboard data to Excel file"""
    from django.db.models import Sum, Avg, Count
    from decimal import Decimal
    
    wb = xlwt.Workbook()
    
    # Get API data
    api_data = APIData.get_api_data()
    
    # Platform filter
    selected_platform_id = request.GET.get('platform', '')
    selected_platform = None
    selected_platform_name = 'All Platforms'
    if selected_platform_id:
        try:
            selected_platform = RemoteMiningPlatform.objects.get(pk=selected_platform_id)
            selected_platform_name = selected_platform.name
        except (RemoteMiningPlatform.DoesNotExist, ValueError):
            selected_platform = None
    
    # NETWORK DATA
    bitcoin_price = api_data.bitcoin_price_usd or 0
    network_hashrate = api_data.network_hashrate_ehs or 0
    network_difficulty = api_data.network_difficulty or 0
    avg_block_fees_24h = api_data.avg_block_fees_24h or 0
    
    # FLEET DATA
    miners = Miner.objects.filter(hashrate__isnull=False, power__isnull=False)
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
    
    # Hashrate weighted average efficiency
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
    
    # Hashrate weighted average energy cost
    hashrate_weighted_energy_cost = 0
    if total_hashrate > 0:
        energy_cost_sum = 0
        total_hashrate_with_energy = 0
        for miner in miners_with_energy:
            energy_cost_sum += float(miner.hashrate) * float(miner.platform.energy_price)
            total_hashrate_with_energy += float(miner.hashrate)
        if total_hashrate_with_energy > 0:
            hashrate_weighted_energy_cost = round(energy_cost_sum / total_hashrate_with_energy, 6)
    
    # REVENUES DATA
    payouts = Payout.objects.all()
    if selected_platform:
        payouts = payouts.filter(platform=selected_platform)
    total_btc_mined = payouts.aggregate(total=Sum('payout_amount'))['total'] or 0
    current_gross_value = float(total_btc_mined) * float(bitcoin_price) if total_btc_mined and bitcoin_price else 0
    total_payouts = payouts.count()
    
    # Calculate Gross Value at Payout (sum of value_at_payout field)
    gross_value_at_payout = payouts.aggregate(total=Sum('value_at_payout'))['total'] or 0
    gross_value_at_payout = float(gross_value_at_payout)
    
    # Calculate Appreciation (Current Gross Value - Gross Value at Payout)
    appreciation = current_gross_value - gross_value_at_payout
    
    # Calculate Total OPEX (sum of all OPEX expenses)
    expenses = Expense.objects.filter(category='OPEX')
    if selected_platform:
        expenses = expenses.filter(platform=selected_platform)
    total_opex = expenses.aggregate(total=Sum('expense_amount'))['total'] or Decimal('0')
    total_opex = float(total_opex)
    
    # Calculate Current Net Value (Current Gross Value - Total OPEX)
    current_net_value = current_gross_value - total_opex
    
    # Sheet 1: Overview Summary
    ws_summary = wb.add_sheet('Overview Summary')
    
    # Headers
    ws_summary.write(0, 0, 'Metric Category')
    ws_summary.write(0, 1, 'Metric Name')
    ws_summary.write(0, 2, 'Value')
    ws_summary.write(0, 3, 'Unit')
    
    # Data rows
    row = 1
    
    # Platform Filter
    ws_summary.write(row, 0, 'Filter')
    ws_summary.write(row, 1, 'Platform')
    ws_summary.write(row, 2, selected_platform_name)
    ws_summary.write(row, 3, '')
    row += 1
    
    # Network Data
    ws_summary.write(row, 0, 'Network Data')
    ws_summary.write(row, 1, 'Bitcoin Spot Price')
    ws_summary.write(row, 2, float(bitcoin_price))
    ws_summary.write(row, 3, 'USD')
    row += 1
    
    ws_summary.write(row, 0, 'Network Data')
    ws_summary.write(row, 1, 'Total Network Hashrate')
    ws_summary.write(row, 2, float(network_hashrate))
    ws_summary.write(row, 3, 'EH/s')
    row += 1
    
    ws_summary.write(row, 0, 'Network Data')
    ws_summary.write(row, 1, 'Network Difficulty')
    ws_summary.write(row, 2, float(network_difficulty))
    ws_summary.write(row, 3, '')
    row += 1
    
    ws_summary.write(row, 0, 'Network Data')
    ws_summary.write(row, 1, '24h Avg Block Fees')
    ws_summary.write(row, 2, float(avg_block_fees_24h))
    ws_summary.write(row, 3, 'BTC')
    row += 1
    
    # Fleet Data
    ws_summary.write(row, 0, 'Fleet Data')
    ws_summary.write(row, 1, 'Miner Count')
    ws_summary.write(row, 2, miner_count)
    ws_summary.write(row, 3, 'units')
    row += 1
    
    ws_summary.write(row, 0, 'Fleet Data')
    ws_summary.write(row, 1, 'Total Hashrate')
    ws_summary.write(row, 2, float(total_hashrate))
    ws_summary.write(row, 3, 'TH/s')
    row += 1
    
    ws_summary.write(row, 0, 'Fleet Data')
    ws_summary.write(row, 1, 'Total Power')
    ws_summary.write(row, 2, round(float(total_power), 2))  # Power already stored in kW in database
    ws_summary.write(row, 3, 'kW')
    row += 1
    
    ws_summary.write(row, 0, 'Fleet Data')
    ws_summary.write(row, 1, 'Total Hardware Cost')
    ws_summary.write(row, 2, float(total_capex))
    ws_summary.write(row, 3, 'USD')
    row += 1
    
    # Efficiency Data
    ws_summary.write(row, 0, 'Efficiency Data')
    ws_summary.write(row, 1, 'Average Efficiency')
    ws_summary.write(row, 2, float(avg_efficiency))
    ws_summary.write(row, 3, 'W/TH')
    row += 1
    
    ws_summary.write(row, 0, 'Efficiency Data')
    ws_summary.write(row, 1, 'Hashrate Weighted Avg Efficiency')
    ws_summary.write(row, 2, float(hashrate_weighted_efficiency))
    ws_summary.write(row, 3, 'W/TH')
    row += 1
    
    # Energy Data
    ws_summary.write(row, 0, 'Energy Data')
    ws_summary.write(row, 1, 'Average Energy Cost')
    ws_summary.write(row, 2, float(avg_energy_cost))
    ws_summary.write(row, 3, '$/kWh')
    row += 1
    
    ws_summary.write(row, 0, 'Energy Data')
    ws_summary.write(row, 1, 'Hashrate Weighted Avg Energy Cost')
    ws_summary.write(row, 2, float(hashrate_weighted_energy_cost))
    ws_summary.write(row, 3, '$/kWh')
    row += 1
    
    # Revenue Data
    ws_summary.write(row, 0, 'Revenue Data')
    ws_summary.write(row, 1, 'Total BTC Mined')
    ws_summary.write(row, 2, float(total_btc_mined))
    ws_summary.write(row, 3, 'BTC')
    row += 1
    
    ws_summary.write(row, 0, 'Revenue Data')
    ws_summary.write(row, 1, 'Current Gross Value')
    ws_summary.write(row, 2, round(float(current_gross_value), 2))
    ws_summary.write(row, 3, 'USD')
    row += 1
    
    ws_summary.write(row, 0, 'Revenue Data')
    ws_summary.write(row, 1, 'Gross Value at Payout')
    ws_summary.write(row, 2, round(float(gross_value_at_payout), 2))
    ws_summary.write(row, 3, 'USD')
    row += 1
    
    ws_summary.write(row, 0, 'Revenue Data')
    ws_summary.write(row, 1, 'Appreciation')
    ws_summary.write(row, 2, round(float(appreciation), 2))
    ws_summary.write(row, 3, 'USD')
    row += 1
    
    ws_summary.write(row, 0, 'Revenue Data')
    ws_summary.write(row, 1, 'Total Payouts')
    ws_summary.write(row, 2, total_payouts)
    ws_summary.write(row, 3, 'count')
    row += 1
    
    ws_summary.write(row, 0, 'Revenue Data')
    ws_summary.write(row, 1, 'Total OPEX')
    ws_summary.write(row, 2, round(float(total_opex), 2))
    ws_summary.write(row, 3, 'USD')
    row += 1
    
    ws_summary.write(row, 0, 'Revenue Data')
    ws_summary.write(row, 1, 'Current Net Value')
    ws_summary.write(row, 2, round(float(current_net_value), 2))
    ws_summary.write(row, 3, 'USD')
    row += 1
    
    # Sheet 2: Hashrate by Platform
    ws_hashrate_platform = wb.add_sheet('Hashrate by Platform')
    ws_hashrate_platform.write(0, 0, 'Platform')
    ws_hashrate_platform.write(0, 1, 'Hashrate (TH/s)')
    
    platform_row = 1
    export_platform_list = [selected_platform] if selected_platform else RemoteMiningPlatform.objects.all()
    for platform in export_platform_list:
        platform_miners = platform.miners.filter(hashrate__isnull=False, power__isnull=False)
        platform_hashrate = platform_miners.aggregate(total=Sum('hashrate'))['total'] or 0
        if platform_hashrate > 0:
            ws_hashrate_platform.write(platform_row, 0, platform.name)
            ws_hashrate_platform.write(platform_row, 1, float(platform_hashrate))
            platform_row += 1
    
    # Sheet 3: Hashrate by Location
    ws_hashrate_location = wb.add_sheet('Hashrate by Location')
    ws_hashrate_location.write(0, 0, 'Location')
    ws_hashrate_location.write(0, 1, 'Hashrate (TH/s)')
    
    location_row = 1
    locations = miners.values_list('location', flat=True).distinct()
    for location in locations:
        if location:
            location_hashrate = miners.filter(location=location).aggregate(total=Sum('hashrate'))['total'] or 0
            if location_hashrate > 0:
                ws_hashrate_location.write(location_row, 0, location)
                ws_hashrate_location.write(location_row, 1, float(location_hashrate))
                location_row += 1
    
    # Sheet 4: Revenue by Platform
    ws_revenue_platform = wb.add_sheet('Revenue by Platform')
    ws_revenue_platform.write(0, 0, 'Platform')
    ws_revenue_platform.write(0, 1, 'BTC Mined')
    ws_revenue_platform.write(0, 2, 'Gross Value (USD)')
    ws_revenue_platform.write(0, 3, 'Gross Value at Payout (USD)')
    ws_revenue_platform.write(0, 4, 'Payout Count')
    
    revenue_row = 1
    for platform in export_platform_list:
        platform_btc = platform.payouts.aggregate(total=Sum('payout_amount'))['total'] or 0
        platform_payouts = platform.payouts.count()
        if platform_btc > 0:
            platform_value = float(platform_btc) * float(bitcoin_price) if bitcoin_price else 0
            platform_gross_value_at_payout = platform.payouts.aggregate(total=Sum('value_at_payout'))['total'] or 0
            platform_gross_value_at_payout = float(platform_gross_value_at_payout)
            ws_revenue_platform.write(revenue_row, 0, platform.name)
            ws_revenue_platform.write(revenue_row, 1, float(platform_btc))
            ws_revenue_platform.write(revenue_row, 2, round(float(platform_value), 2))
            ws_revenue_platform.write(revenue_row, 3, round(float(platform_gross_value_at_payout), 2))
            ws_revenue_platform.write(revenue_row, 4, platform_payouts)
            revenue_row += 1
    
    response = HttpResponse(
        content_type='application/vnd.ms-excel'
    )
    platform_suffix = f'_{selected_platform_name.replace(" ", "_")}' if selected_platform else ''
    response['Content-Disposition'] = f'attachment; filename="overview_dashboard{platform_suffix}_export.xls"'
    wb.save(response)
    return response


# Data Import Views
def import_platform_data(request):
    """Import platform data from uploaded Excel file"""
    if request.method == 'POST' and request.FILES.get('import_file'):
        try:
            file = request.FILES['import_file']
            wb = xlrd.open_workbook(file_contents=file.read())
            ws = wb.sheet_by_index(0)
            
            # Get headers from first row
            headers = []
            for col in range(ws.ncols):
                header = ws.cell_value(0, col)
                if header:
                    headers.append(str(header).strip())
            
            # Process data rows
            imported_count = 0
            for row in range(1, ws.nrows):
                row_data = {}
                for col, header in enumerate(headers):
                    if col < ws.ncols:
                        cell_value = ws.cell_value(row, col)
                        cell_type = ws.cell_type(row, col)
                        if cell_value:
                            # Handle different cell types
                            if cell_type == 3:  # Date type
                                # Convert Excel date serial number to Python date
                                date_tuple = xlrd.xldate_as_tuple(cell_value, wb.datemode)
                                row_data[header] = datetime(*date_tuple).date()
                            elif cell_type == 2:  # Number type
                                row_data[header] = cell_value
                            else:  # Text type
                                row_data[header] = str(cell_value).strip()
                
                if row_data:  # Skip empty rows
                    # Create platform instance
                    platform_data = {}
                    for field, value in row_data.items():
                        if hasattr(RemoteMiningPlatform, field) and value:
                            if field == 'energy_price' and value:
                                platform_data[field] = Decimal(str(value))
                            else:
                                platform_data[field] = value
                    
                    if platform_data:
                        RemoteMiningPlatform.objects.create(**platform_data)
                        imported_count += 1
            
            messages.success(request, f'Successfully imported {imported_count} platforms!')
            return redirect('platform_list')
            
        except Exception as e:
            messages.error(request, 'Wrong import file format or data. Please check your file and try again.')
            return redirect('platform_list')
    
    return redirect('platform_list')


def import_miner_data(request):
    """Import miner data from uploaded Excel file"""
    if request.method == 'POST' and request.FILES.get('import_file'):
        try:
            file = request.FILES['import_file']
            wb = xlrd.open_workbook(file_contents=file.read())
            ws = wb.sheet_by_index(0)
            
            # Get headers from first row
            headers = []
            for col in range(ws.ncols):
                header = ws.cell_value(0, col)
                if header:
                    headers.append(str(header).strip())
            
            # Process data rows
            imported_count = 0
            for row in range(1, ws.nrows):
                row_data = {}
                for col, header in enumerate(headers):
                    if col < ws.ncols:
                        cell_value = ws.cell_value(row, col)
                        cell_type = ws.cell_type(row, col)
                        if cell_value:
                            # Handle different cell types
                            if cell_type == 3:  # Date type
                                # Convert Excel date serial number to Python date
                                date_tuple = xlrd.xldate_as_tuple(cell_value, wb.datemode)
                                row_data[header] = datetime(*date_tuple).date()
                            elif cell_type == 2:  # Number type
                                row_data[header] = cell_value
                            else:  # Text type
                                row_data[header] = str(cell_value).strip()
                
                if row_data:  # Skip empty rows
                    # Create miner instance
                    miner_data = {}
                    for field, value in row_data.items():
                        if hasattr(Miner, field) and value:
                            if field == 'platform':
                                # Handle foreign key - expect platform ID
                                try:
                                    platform = RemoteMiningPlatform.objects.get(pk=int(float(value)))
                                    miner_data[field] = platform
                                except:
                                    continue
                            elif field in ['hashrate', 'power', 'efficiency', 'purchase_price'] and value:
                                miner_data[field] = Decimal(str(value))
                            elif field in ['purchase_date', 'start_date'] and value:
                                if isinstance(value, str):
                                    miner_data[field] = datetime.strptime(value, '%Y-%m-%d').date()
                                else:
                                    miner_data[field] = value
                            else:
                                miner_data[field] = value
                    
                    if miner_data:
                        Miner.objects.create(**miner_data)
                        imported_count += 1
            
            messages.success(request, f'Successfully imported {imported_count} miners!')
            return redirect('miner_list')
            
        except Exception as e:
            messages.error(request, 'Wrong import file format or data. Please check your file and try again.')
            return redirect('miner_list')
    
    return redirect('miner_list')


def import_payout_data(request):
    """Import payout data from uploaded Excel file"""
    if request.method == 'POST' and request.FILES.get('import_file'):
        try:
            file = request.FILES['import_file']
            wb = xlrd.open_workbook(file_contents=file.read())
            ws = wb.sheet_by_index(0)
            
            # Get headers from first row
            headers = []
            for col in range(ws.ncols):
                header = ws.cell_value(0, col)
                if header:
                    headers.append(str(header).strip())
            
            # Process data rows
            imported_count = 0
            for row in range(1, ws.nrows):
                row_data = {}
                for col, header in enumerate(headers):
                    if col < ws.ncols:
                        cell_value = ws.cell_value(row, col)
                        cell_type = ws.cell_type(row, col)
                        if cell_value:
                            # Handle different cell types
                            if cell_type == 3:  # Date type
                                # Convert Excel date serial number to Python date
                                date_tuple = xlrd.xldate_as_tuple(cell_value, wb.datemode)
                                row_data[header] = datetime(*date_tuple).date()
                            elif cell_type == 2:  # Number type
                                row_data[header] = cell_value
                            else:  # Text type
                                row_data[header] = str(cell_value).strip()
                
                if row_data:  # Skip empty rows
                    # Create payout instance
                    payout_data = {}
                    for field, value in row_data.items():
                        if hasattr(Payout, field) and value:
                            if field == 'platform':
                                # Handle foreign key - expect platform ID
                                try:
                                    platform = RemoteMiningPlatform.objects.get(pk=int(float(value)))
                                    payout_data[field] = platform
                                except:
                                    continue
                            elif field == 'payout_amount' and value:
                                payout_data[field] = Decimal(str(value))
                            elif field == 'closing_price' and value:
                                payout_data[field] = Decimal(str(value))
                            elif field == 'payout_date' and value:
                                if isinstance(value, str):
                                    payout_data[field] = datetime.strptime(value, '%Y-%m-%d').date()
                                else:
                                    payout_data[field] = value
                            else:
                                payout_data[field] = value
                    
                    if payout_data:
                        Payout.objects.create(**payout_data)
                        imported_count += 1
            
            messages.success(request, f'Successfully imported {imported_count} payouts!')
            return redirect('payout_list')
            
        except Exception as e:
            messages.error(request, 'Wrong import file format or data. Please check your file and try again.')
            return redirect('payout_list')
    
    return redirect('payout_list')


# Expense Import/Export Functions
def download_expense_template(request):
    """Download import template for Expenses"""
    wb = xlwt.Workbook()
    ws = wb.add_sheet('Expense Import Template')
    
    # Add headers based on form fields
    headers = ['expense_date', 'platform', 'category', 'description', 'expense_amount', 'invoice_link', 'receipt_link', 'notes']
    
    for col, header in enumerate(headers):
        ws.write(0, col, header)
    
    response = HttpResponse(
        content_type='application/vnd.ms-excel'
    )
    response['Content-Disposition'] = 'attachment; filename="expense_import_template.xls"'
    wb.save(response)
    return response


def export_expense_data(request):
    """Export all expense data to Excel file"""
    wb = xlwt.Workbook()
    ws = wb.add_sheet('Expense Data')
    
    # Add headers
    headers = ['expense_date', 'platform', 'category', 'description', 'expense_amount', 'invoice_link', 'receipt_link', 'notes']
    
    for col, header in enumerate(headers):
        ws.write(0, col, header)
    
    # Add data rows
    expenses = Expense.objects.all().order_by('-expense_date')
    for row, expense in enumerate(expenses, start=1):
        ws.write(row, 0, expense.expense_date.strftime('%Y-%m-%d') if expense.expense_date else '')
        ws.write(row, 1, expense.platform.id if expense.platform else '')
        ws.write(row, 2, expense.category or '')
        ws.write(row, 3, expense.description or '')
        ws.write(row, 4, float(expense.expense_amount) if expense.expense_amount else '')
        ws.write(row, 5, expense.invoice_link or '')
        ws.write(row, 6, expense.receipt_link or '')
        ws.write(row, 7, expense.notes or '')
    
    response = HttpResponse(
        content_type='application/vnd.ms-excel'
    )
    response['Content-Disposition'] = 'attachment; filename="expense_data_export.xls"'
    wb.save(response)
    return response


def import_expense_data(request):
    """Import expense data from uploaded Excel file"""
    if request.method == 'POST' and request.FILES.get('import_file'):
        try:
            file = request.FILES['import_file']
            wb = xlrd.open_workbook(file_contents=file.read())
            ws = wb.sheet_by_index(0)
            
            # Get headers from first row
            headers = []
            for col in range(ws.ncols):
                header = ws.cell_value(0, col)
                headers.append(header.lower().strip())
            
            # Process data rows
            imported_count = 0
            for row in range(1, ws.nrows):
                expense_data = {}
                
                for col, header in enumerate(headers):
                    if col >= ws.ncols:
                        continue
                        
                    cell_value = ws.cell_value(row, col)
                    
                    if header == 'expense_date' and cell_value:
                        try:
                            if isinstance(cell_value, float):
                                # Excel date as float
                                from datetime import date
                                import xldate
                                expense_data['expense_date'] = xldate.xldate_as_datetime(cell_value, wb.datemode).date()
                            else:
                                # String date
                                expense_data['expense_date'] = datetime.strptime(str(cell_value), '%Y-%m-%d').date()
                        except:
                            continue
                    elif header == 'platform' and cell_value:
                        try:
                            platform_id = int(float(cell_value))
                            platform = RemoteMiningPlatform.objects.get(pk=platform_id)
                            expense_data['platform'] = platform
                        except:
                            pass
                    elif header == 'category' and cell_value:
                        category_value = str(cell_value).upper().strip()
                        if category_value in ['CAPEX', 'OPEX']:
                            expense_data['category'] = category_value
                    elif header == 'description' and cell_value:
                        expense_data['description'] = str(cell_value)
                    elif header == 'expense_amount' and cell_value:
                        try:
                            expense_data['expense_amount'] = Decimal(str(cell_value))
                        except:
                            pass
                    elif header == 'invoice_link' and cell_value:
                        expense_data['invoice_link'] = str(cell_value)
                    elif header == 'receipt_link' and cell_value:
                        expense_data['receipt_link'] = str(cell_value)
                    elif header == 'notes' and cell_value:
                        expense_data['notes'] = str(cell_value)
                
                # Create expense if we have required fields
                if 'expense_date' in expense_data and 'category' in expense_data and 'expense_amount' in expense_data:
                    Expense.objects.create(**expense_data)
                    imported_count += 1
            
            messages.success(request, f'Successfully imported {imported_count} expenses!')
            return redirect('expense_list')
            
        except Exception as e:
            messages.error(request, 'Wrong import file format or data. Please check your file and try again.')
            return redirect('expense_list')
    
    return redirect('expense_list')


# ===== TOP-UP VIEWS =====

class TopUpListView(ListView):
    model = TopUp
    template_name = 'mining/topup_list.html'
    context_object_name = 'topups'
    paginate_by = 50


class TopUpDetailView(DetailView):
    model = TopUp
    template_name = 'mining/topup_detail.html'
    context_object_name = 'topup'
    
    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        
        # Add previous and next navigation
        topup = self.get_object()
        context['previous_topup'] = TopUp.objects.filter(
            id__lt=topup.id
        ).order_by('-id').first()
        
        context['next_topup'] = TopUp.objects.filter(
            id__gt=topup.id
        ).order_by('id').first()
        
        return context


class TopUpCreateView(CreateView):
    model = TopUp
    form_class = TopUpForm
    template_name = 'mining/topup_form.html'
    
    def form_valid(self, form):
        response = super().form_valid(form)
        messages.success(self.request, 'Top-Up created successfully!')
        return response
    
    def get_success_url(self):
        return reverse('topup_detail', kwargs={'pk': self.object.pk})


class TopUpUpdateView(UpdateView):
    model = TopUp
    form_class = TopUpForm
    template_name = 'mining/topup_form.html'
    
    def form_valid(self, form):
        response = super().form_valid(form)
        messages.success(self.request, 'Top-Up updated successfully!')
        return response
    
    def get_success_url(self):
        return reverse('topup_detail', kwargs={'pk': self.object.pk})


class TopUpDeleteView(DeleteView):
    model = TopUp
    template_name = 'mining/topup_confirm_delete.html'
    
    def delete(self, request, *args, **kwargs):
        response = super().delete(request, *args, **kwargs)
        messages.success(request, 'Top-Up deleted successfully!')
        return response
    
    def get_success_url(self):
        return reverse('topup_list')


def download_topup_template(request):
    """Download Excel template for Top-Up import"""
    wb = xlwt.Workbook()
    ws = wb.add_sheet('TopUp Template')
    
    # Add headers
    headers = ['topup_date', 'platform', 'topup_amount', 'description', 'receipt_link']
    
    for col, header in enumerate(headers):
        ws.write(0, col, header)
    
    response = HttpResponse(
        content_type='application/vnd.ms-excel'
    )
    response['Content-Disposition'] = 'attachment; filename="topup_import_template.xls"'
    wb.save(response)
    return response


def export_topup_data(request):
    """Export all top-up data to Excel file"""
    wb = xlwt.Workbook()
    ws = wb.add_sheet('TopUp Data')
    
    # Add headers
    headers = ['topup_date', 'platform', 'topup_amount', 'description', 'receipt_link']
    
    for col, header in enumerate(headers):
        ws.write(0, col, header)
    
    # Add data rows
    topups = TopUp.objects.all().order_by('-topup_date')
    for row, topup in enumerate(topups, start=1):
        ws.write(row, 0, str(topup.topup_date) if topup.topup_date else '')
        ws.write(row, 1, topup.platform.name if topup.platform else '')
        ws.write(row, 2, float(topup.topup_amount) if topup.topup_amount else '')
        ws.write(row, 3, topup.description or '')
        ws.write(row, 4, topup.receipt_link or '')
    
    response = HttpResponse(
        content_type='application/vnd.ms-excel'
    )
    response['Content-Disposition'] = 'attachment; filename="topup_data_export.xls"'
    wb.save(response)
    return response


def import_topup_data(request):
    """Import top-up data from uploaded Excel file"""
    if request.method == 'POST' and request.FILES.get('import_file'):
        try:
            file = request.FILES['import_file']
            wb = xlrd.open_workbook(file_contents=file.read())
            ws = wb.sheet_by_index(0)
            
            # Get headers from first row
            headers = [str(ws.cell_value(0, col)).lower().strip() for col in range(ws.ncols)]
            
            imported_count = 0
            
            # Process each row (skip header row)
            for row in range(1, ws.nrows):
                topup_data = {}
                
                # Process each column
                for col, header in enumerate(headers):
                    if col >= ws.ncols:
                        break
                        
                    cell_value = ws.cell_value(row, col)
                    
                    if header == 'topup_date' and cell_value:
                        # Handle date conversion
                        if isinstance(cell_value, float):
                            try:
                                from datetime import datetime, date
                                dt = xlrd.xldate_as_datetime(cell_value, wb.datemode)
                                topup_data['topup_date'] = dt.date()
                            except:
                                continue
                        else:
                            try:
                                from datetime import datetime
                                topup_data['topup_date'] = datetime.strptime(str(cell_value), '%Y-%m-%d').date()
                            except:
                                continue
                    elif header == 'platform' and cell_value:
                        try:
                            # Try to find platform by ID or name
                            if isinstance(cell_value, float):
                                platform = RemoteMiningPlatform.objects.get(id=int(cell_value))
                            else:
                                platform = RemoteMiningPlatform.objects.get(name=str(cell_value))
                            topup_data['platform'] = platform
                        except RemoteMiningPlatform.DoesNotExist:
                            continue
                    elif header == 'topup_amount' and cell_value:
                        try:
                            topup_data['topup_amount'] = float(cell_value)
                        except (ValueError, TypeError):
                            continue
                    elif header == 'description' and cell_value:
                        topup_data['description'] = str(cell_value)
                    elif header == 'receipt_link' and cell_value:
                        topup_data['receipt_link'] = str(cell_value)
                
                # Create top-up if we have required fields
                if 'topup_date' in topup_data and 'platform' in topup_data and 'topup_amount' in topup_data:
                    TopUp.objects.create(**topup_data)
                    imported_count += 1
            
            messages.success(request, f'Successfully imported {imported_count} top-ups!')
            return redirect('topup_list')
            
        except Exception as e:
            messages.error(request, 'Wrong import file format or data. Please check your file and try again.')
            return redirect('topup_list')
    
    return redirect('topup_list')


def forecasting_dashboard(request):
    """Forecasting Dashboard with BTC mining profitability calculations"""
    from decimal import Decimal
    from django.db.models import Sum, Avg
    
    # Gather all required data from database models
    api_data = APIData.get_api_data()
    settings = Settings.get_settings()
    
    # Platform filter
    platforms = RemoteMiningPlatform.objects.all()
    selected_platform_id = request.GET.get('platform', '')
    selected_platform = None
    if selected_platform_id:
        try:
            selected_platform = RemoteMiningPlatform.objects.get(pk=selected_platform_id)
        except (RemoteMiningPlatform.DoesNotExist, ValueError):
            selected_platform = None
    
    # Get miners with valid hashrate and power data, only active miners
    total_miner_count = Miner.objects.count()
    miners = Miner.objects.filter(hashrate__isnull=False, power__isnull=False, is_active=True)
    if selected_platform:
        miners = miners.filter(platform=selected_platform)
    
    # Get total hashrate
    total_hashrate = miners.aggregate(total=Sum('hashrate'))['total'] or Decimal('0')
    
    # Get miner count (accounted for on dashboard)
    miner_count = miners.count()
    
    # Get total hardware cost (sum of miner purchase prices)
    total_capex = miners.aggregate(total=Sum('purchase_price'))['total'] or Decimal('0')
    
    # Calculate hashrate weighted average efficiency
    hashrate_weighted_efficiency = Decimal('0')
    if total_hashrate > 0:
        total_weighted = Decimal('0')
        for miner in miners.filter(efficiency__isnull=False):
            total_weighted += miner.hashrate * miner.efficiency
        hashrate_weighted_efficiency = total_weighted / total_hashrate if total_weighted > 0 else Decimal('0')
    
    # Calculate hashrate weighted average energy cost (denominator = only miners with energy prices)
    hashrate_weighted_energy_cost = Decimal('0')
    if total_hashrate > 0:
        total_weighted = Decimal('0')
        total_hashrate_with_energy = Decimal('0')
        for miner in miners.filter(platform__energy_price__isnull=False):
            total_weighted += miner.hashrate * miner.platform.energy_price
            total_hashrate_with_energy += miner.hashrate
        if total_hashrate_with_energy > 0:
            hashrate_weighted_energy_cost = total_weighted / total_hashrate_with_energy
    
    # Simple average efficiency
    avg_efficiency = miners.filter(efficiency__isnull=False).aggregate(avg=Avg('efficiency'))['avg'] or Decimal('0')
    if avg_efficiency:
        avg_efficiency = round(float(avg_efficiency), 2)
    
    # Simple average energy cost
    miners_with_energy = miners.filter(platform__energy_price__isnull=False)
    avg_energy_cost = miners_with_energy.aggregate(avg=Avg('platform__energy_price'))['avg'] or Decimal('0')
    if avg_energy_cost:
        avg_energy_cost = round(float(avg_energy_cost), 6)
    
    # Get data from API and settings
    network_difficulty = api_data.network_difficulty or 0
    network_hashrate_ehs = float(api_data.network_hashrate_ehs or Decimal('0'))
    avg_tx_fees = float(api_data.avg_block_fees_24h or Decimal('0'))
    pool_fee = float(settings.pool_fee_percentage)
    btc_price_usd = float(api_data.bitcoin_price_usd or Decimal('0'))
    price_per_kwh = float(hashrate_weighted_energy_cost)
    efficiency_w_th = float(hashrate_weighted_efficiency)
    hardware_cost_usd = float(total_capex)
    
    # Perform calculations using difficulty-based formula
    results = None
    if total_hashrate > 0 and network_difficulty > 0 and btc_price_usd > 0:
        # Convert hashrate to H/s
        miner_hashrate_hs = float(total_hashrate) * 1e12  # TH/s to H/s
        
        # Hashrate share for display purposes only
        network_hashrate_hs = network_hashrate_ehs * 1e18 if network_hashrate_ehs > 0 else 0  # EH/s to H/s
        hashrate_share_percent = (miner_hashrate_hs / network_hashrate_hs * 100) if network_hashrate_hs > 0 else 0
        
        # Daily mining calculations using difficulty-based formula
        # Expected blocks per day = (hashrate * 86400) / (difficulty * 2^32)
        block_reward = float(settings.block_reward) + avg_tx_fees
        expected_blocks_per_day = (miner_hashrate_hs * 86400) / (network_difficulty * 2**32)
        daily_btc_gross_before_fee = expected_blocks_per_day * block_reward
        pool_fee_btc = daily_btc_gross_before_fee * (pool_fee / 100)
        daily_btc_after_fee = daily_btc_gross_before_fee - pool_fee_btc
        
        # Power and energy calculations
        miner_hashrate_ths = miner_hashrate_hs / 1e12
        power_watts = miner_hashrate_ths * efficiency_w_th
        daily_energy_kwh = (power_watts * 24) / 1000
        daily_electricity_cost_usd = daily_energy_kwh * price_per_kwh
        daily_electricity_cost_btc = daily_electricity_cost_usd / btc_price_usd if btc_price_usd > 0 else 0
        
        # Calculate electricity cost as percentage of mined BTC
        energy_cost_percentage = (daily_electricity_cost_btc / daily_btc_after_fee * 100) if daily_btc_after_fee > 0 else 0
        
        # USD calculations
        daily_usd_gross = daily_btc_gross_before_fee * btc_price_usd
        daily_usd_after_fee = daily_btc_after_fee * btc_price_usd
        daily_usd_net = daily_usd_after_fee - daily_electricity_cost_usd
        daily_btc_net = daily_usd_net / btc_price_usd if btc_price_usd > 0 else 0
        
        # Cost basis calculation
        total_cost_usd = (pool_fee_btc * btc_price_usd) + daily_electricity_cost_usd
        cost_basis_usd_per_btc = total_cost_usd / daily_btc_after_fee if daily_btc_after_fee > 0 else 0
        discount_vs_market_pct = -1 * ((btc_price_usd - cost_basis_usd_per_btc) / btc_price_usd * 100) if btc_price_usd > 0 else 0
        
        # Net profit margin
        margin = (daily_usd_net / daily_usd_gross * 100) if daily_usd_gross > 0 else 0
        
        # Time calculations (using BTC after pool fee but before electricity, matching original script)
        days_to_mine_1_btc = 1 / daily_btc_after_fee if daily_btc_after_fee > 0 else float('inf')
        days_to_mine_small_btc = 0.005 / daily_btc_after_fee if daily_btc_after_fee > 0 else float('inf')
        
        # ROI calculation
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
        
        results = {
            'network_hashrate_ehs': network_hashrate_ehs,
            'hashrate_share_percent': hashrate_share_percent,
            'power_consumption_watts': power_watts,
            'power_consumption_kw': power_watts / 1000,
            'daily_energy_kwh': daily_energy_kwh,
            'energy_cost_percentage': energy_cost_percentage,
            'margin': margin,
            'days_to_mine_1_btc': days_to_mine_1_btc,
            'time_to_mine_1_btc': {
                'years': int(days_to_mine_1_btc / 365) if days_to_mine_1_btc != float('inf') else 0,
                'months': int((days_to_mine_1_btc % 365) / 30) if days_to_mine_1_btc != float('inf') else 0,
                'days': int(days_to_mine_1_btc % 30) if days_to_mine_1_btc != float('inf') else 0
            },
            'days_to_mine_small_btc': days_to_mine_small_btc,
            'time_to_mine_small_btc': {
                'years': int(days_to_mine_small_btc / 365) if days_to_mine_small_btc != float('inf') else 0,
                'months': int((days_to_mine_small_btc % 365) / 30) if days_to_mine_small_btc != float('inf') else 0,
                'days': int(days_to_mine_small_btc % 30) if days_to_mine_small_btc != float('inf') else 0
            },
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
    
    context = {
        # Platform filter
        'platforms': platforms,
        'selected_platform': selected_platform,
        # Input parameters
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
        # Efficiency & Energy KPIs
        'avg_efficiency': avg_efficiency,
        'hashrate_weighted_efficiency': round(float(hashrate_weighted_efficiency), 2),
        'avg_energy_cost': avg_energy_cost,
        'hashrate_weighted_energy_cost': round(float(hashrate_weighted_energy_cost), 6),
        # Calculation results
        'results': results,
    }
    
    return render(request, 'mining/forecasting_dashboard.html', context)


def export_forecasting_data(request):
    """Export forecasting dashboard data to Excel file - EXACT COPY of dashboard calculations"""
    from decimal import Decimal
    from django.db.models import Sum, Avg
    import xlwt
    
    wb = xlwt.Workbook()
    
    # EXACT COPY of forecasting_dashboard function logic
    api_data = APIData.get_api_data()
    settings = Settings.get_settings()
    
    # Platform filter
    selected_platform_id = request.GET.get('platform', '')
    selected_platform = None
    selected_platform_name = 'All Platforms'
    if selected_platform_id:
        try:
            selected_platform = RemoteMiningPlatform.objects.get(pk=selected_platform_id)
            selected_platform_name = selected_platform.name
        except (RemoteMiningPlatform.DoesNotExist, ValueError):
            selected_platform = None
    
    # Get miners with valid hashrate and power data, only active miners
    total_miner_count = Miner.objects.count()
    miners = Miner.objects.filter(hashrate__isnull=False, power__isnull=False, is_active=True)
    if selected_platform:
        miners = miners.filter(platform=selected_platform)
    
    # Get total hashrate
    total_hashrate = miners.aggregate(total=Sum('hashrate'))['total'] or Decimal('0')
    
    # Get miner count (accounted for on dashboard)
    miner_count = miners.count()
    
    # Get total hardware cost (sum of miner purchase prices)
    total_capex = miners.aggregate(total=Sum('purchase_price'))['total'] or Decimal('0')
    
    # Calculate hashrate weighted average efficiency
    hashrate_weighted_efficiency = Decimal('0')
    if total_hashrate > 0:
        total_weighted = Decimal('0')
        for miner in miners.filter(efficiency__isnull=False):
            total_weighted += miner.hashrate * miner.efficiency
        hashrate_weighted_efficiency = total_weighted / total_hashrate if total_weighted > 0 else Decimal('0')
    
    # Calculate hashrate weighted average energy cost (denominator = only miners with energy prices)
    hashrate_weighted_energy_cost = Decimal('0')
    if total_hashrate > 0:
        total_weighted = Decimal('0')
        total_hashrate_with_energy = Decimal('0')
        for miner in miners.filter(platform__energy_price__isnull=False):
            total_weighted += miner.hashrate * miner.platform.energy_price
            total_hashrate_with_energy += miner.hashrate
        if total_hashrate_with_energy > 0:
            hashrate_weighted_energy_cost = total_weighted / total_hashrate_with_energy
    
    # Simple average efficiency
    avg_efficiency = miners.filter(efficiency__isnull=False).aggregate(avg=Avg('efficiency'))['avg'] or Decimal('0')
    if avg_efficiency:
        avg_efficiency = round(float(avg_efficiency), 2)
    
    # Simple average energy cost
    miners_with_energy = miners.filter(platform__energy_price__isnull=False)
    avg_energy_cost = miners_with_energy.aggregate(avg=Avg('platform__energy_price'))['avg'] or Decimal('0')
    if avg_energy_cost:
        avg_energy_cost = round(float(avg_energy_cost), 6)
    
    # Get data from API and settings
    network_difficulty = api_data.network_difficulty or 0
    network_hashrate_ehs = float(api_data.network_hashrate_ehs or Decimal('0'))
    avg_tx_fees = float(api_data.avg_block_fees_24h or Decimal('0'))
    pool_fee = float(settings.pool_fee_percentage)
    btc_price_usd = float(api_data.bitcoin_price_usd or Decimal('0'))
    price_per_kwh = float(hashrate_weighted_energy_cost)
    efficiency_w_th = float(hashrate_weighted_efficiency)
    hardware_cost_usd = float(total_capex)
    
    # Check if we have required data
    if not api_data:
        ws = wb.add_sheet('Error')
        ws.write(0, 0, 'Error: API data not available')
    elif total_hashrate == 0:
        ws = wb.add_sheet('Error')
        ws.write(0, 0, 'Error: No miners with hashrate data available')
    elif network_difficulty <= 0 or btc_price_usd <= 0:
        ws = wb.add_sheet('Error')
        ws.write(0, 0, 'Error: Invalid network difficulty or price data')
    else:
        # EXACT SAME CALCULATIONS AS DASHBOARD
        # Convert hashrate to H/s
        miner_hashrate_hs = float(total_hashrate) * 1e12  # TH/s to H/s
        
        # Hashrate share for display purposes only
        network_hashrate_hs = network_hashrate_ehs * 1e18 if network_hashrate_ehs > 0 else 0  # EH/s to H/s
        hashrate_share_percent = (miner_hashrate_hs / network_hashrate_hs * 100) if network_hashrate_hs > 0 else 0
        
        # Daily mining calculations using difficulty-based formula
        # Expected blocks per day = (hashrate * 86400) / (difficulty * 2^32)
        block_reward = float(settings.block_reward) + avg_tx_fees
        expected_blocks_per_day = (miner_hashrate_hs * 86400) / (network_difficulty * 2**32)
        daily_btc_gross_before_fee = expected_blocks_per_day * block_reward
        pool_fee_btc = daily_btc_gross_before_fee * (pool_fee / 100)
        daily_btc_after_fee = daily_btc_gross_before_fee - pool_fee_btc
        
        # Power and energy calculations
        miner_hashrate_ths = miner_hashrate_hs / 1e12
        power_watts = miner_hashrate_ths * efficiency_w_th
        daily_energy_kwh = (power_watts * 24) / 1000
        daily_electricity_cost_usd = daily_energy_kwh * price_per_kwh
        daily_electricity_cost_btc = daily_electricity_cost_usd / btc_price_usd if btc_price_usd > 0 else 0
        
        # Calculate electricity cost as percentage of mined BTC
        energy_cost_percentage = (daily_electricity_cost_btc / daily_btc_after_fee * 100) if daily_btc_after_fee > 0 else 0
        
        # USD calculations
        daily_usd_gross = daily_btc_gross_before_fee * btc_price_usd
        daily_usd_after_fee = daily_btc_after_fee * btc_price_usd
        daily_usd_net = daily_usd_after_fee - daily_electricity_cost_usd
        daily_btc_net = daily_usd_net / btc_price_usd if btc_price_usd > 0 else 0
        
        # Cost basis calculation
        total_cost_usd = (pool_fee_btc * btc_price_usd) + daily_electricity_cost_usd
        cost_basis_usd_per_btc = total_cost_usd / daily_btc_after_fee if daily_btc_after_fee > 0 else 0
        discount_vs_market_pct = -1 * ((btc_price_usd - cost_basis_usd_per_btc) / btc_price_usd * 100) if btc_price_usd > 0 else 0
        
        # Net profit margin
        margin = (daily_usd_net / daily_usd_gross * 100) if daily_usd_gross > 0 else 0
        
        # Time calculations (using BTC after pool fee but before electricity, matching original script)
        days_to_mine_1_btc = 1 / daily_btc_after_fee if daily_btc_after_fee > 0 else float('inf')
        days_to_mine_small_btc = 0.005 / daily_btc_after_fee if daily_btc_after_fee > 0 else float('inf')
        
        # ROI calculation
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
        
        # Time formatting function
        def format_time_breakdown(days_total):
            if days_total == float('inf'):
                return 'Never (not profitable)'
            years = int(days_total / 365)
            months = int((days_total % 365) / 30)
            remaining_days = int(days_total % 30)
            return f"{years} years, {months} months, {remaining_days} days"
        
        time_to_mine_1_btc = format_time_breakdown(days_to_mine_1_btc)
        time_to_mine_small_btc = format_time_breakdown(days_to_mine_small_btc)
        
        # Sheet 1: Forecasting Summary
        ws_summary = wb.add_sheet('Forecasting Summary')
        ws_summary.write(0, 0, 'Metric Category')
        ws_summary.write(0, 1, 'Metric Name')
        ws_summary.write(0, 2, 'Value')
        ws_summary.write(0, 3, 'Unit')
        
        row = 1
        
        # Platform Filter
        ws_summary.write(row, 0, 'Filter')
        ws_summary.write(row, 1, 'Platform')
        ws_summary.write(row, 2, selected_platform_name)
        ws_summary.write(row, 3, '')
        row += 1
        
        # Network Overview
        ws_summary.write(row, 0, 'Network Overview')
        ws_summary.write(row, 1, 'Network Difficulty')
        ws_summary.write(row, 2, network_difficulty)
        ws_summary.write(row, 3, '')
        row += 1
        
        ws_summary.write(row, 0, 'Network Overview')
        ws_summary.write(row, 1, 'Network Hashrate')
        ws_summary.write(row, 2, network_hashrate_ehs)
        ws_summary.write(row, 3, 'EH/s')
        row += 1
        
        ws_summary.write(row, 0, 'Network Overview')
        ws_summary.write(row, 1, 'Avg Block Fees (24h)')
        ws_summary.write(row, 2, round(avg_tx_fees, 8))
        ws_summary.write(row, 3, 'BTC')
        row += 1
        
        ws_summary.write(row, 0, 'Network Overview')
        ws_summary.write(row, 1, 'BTC Price')
        ws_summary.write(row, 2, round(btc_price_usd, 2))
        ws_summary.write(row, 3, 'USD')
        row += 1
        
        # My Fleet Overview
        ws_summary.write(row, 0, 'My Fleet Overview')
        ws_summary.write(row, 1, 'Miners Accounted For')
        ws_summary.write(row, 2, miner_count)
        ws_summary.write(row, 3, f'of {total_miner_count}')
        row += 1
        
        ws_summary.write(row, 0, 'My Fleet Overview')
        ws_summary.write(row, 1, 'My Hashrate')
        ws_summary.write(row, 2, float(total_hashrate))
        ws_summary.write(row, 3, 'TH/s')
        row += 1
        
        ws_summary.write(row, 0, 'My Fleet Overview')
        ws_summary.write(row, 1, 'Pool Fee')
        ws_summary.write(row, 2, pool_fee)
        ws_summary.write(row, 3, '%')
        row += 1
        
        ws_summary.write(row, 0, 'My Fleet Overview')
        ws_summary.write(row, 1, 'Power Consumption')
        ws_summary.write(row, 2, round(power_watts / 1000, 2))
        ws_summary.write(row, 3, 'kW')
        row += 1
        
        ws_summary.write(row, 0, 'My Fleet Overview')
        ws_summary.write(row, 1, 'Net Profit Margin')
        ws_summary.write(row, 2, round(margin, 2))
        ws_summary.write(row, 3, '%')
        row += 1
        
        # Efficiency Data
        ws_summary.write(row, 0, 'Efficiency Data')
        ws_summary.write(row, 1, 'Average Efficiency')
        ws_summary.write(row, 2, float(avg_efficiency))
        ws_summary.write(row, 3, 'W/TH')
        row += 1
        
        ws_summary.write(row, 0, 'Efficiency Data')
        ws_summary.write(row, 1, 'Hashrate Weighted Avg Efficiency')
        ws_summary.write(row, 2, round(float(hashrate_weighted_efficiency), 2))
        ws_summary.write(row, 3, 'W/TH')
        row += 1
        
        # Energy Data
        ws_summary.write(row, 0, 'Energy Data')
        ws_summary.write(row, 1, 'Average Energy Cost')
        ws_summary.write(row, 2, float(avg_energy_cost))
        ws_summary.write(row, 3, '$/kWh')
        row += 1
        
        ws_summary.write(row, 0, 'Energy Data')
        ws_summary.write(row, 1, 'Hashrate Weighted Avg Energy Cost')
        ws_summary.write(row, 2, round(float(hashrate_weighted_energy_cost), 6))
        ws_summary.write(row, 3, '$/kWh')
        row += 1
        
        # Key Metrics
        ws_summary.write(row, 0, 'Key Metrics')
        ws_summary.write(row, 1, 'Time to mine 1 BTC')
        ws_summary.write(row, 2, time_to_mine_1_btc)
        ws_summary.write(row, 3, '')
        row += 1
        
        ws_summary.write(row, 0, 'Key Metrics')
        ws_summary.write(row, 1, 'Time to mine 0.005 BTC')
        ws_summary.write(row, 2, time_to_mine_small_btc)
        ws_summary.write(row, 3, '')
        row += 1
        
        if roi_data:
            roi_display = f"{roi_data['time_breakdown']['years']} years, {roi_data['time_breakdown']['months']} months, {roi_data['time_breakdown']['days']} days" if roi_data['days_to_roi'] != float('inf') else "Never (not profitable)"
            ws_summary.write(row, 0, 'Key Metrics')
            ws_summary.write(row, 1, 'ROI')
            ws_summary.write(row, 2, roi_display)
            ws_summary.write(row, 3, '')
            row += 1
        
        # Sheet 2: Daily Projections
        ws_daily = wb.add_sheet('Daily Projections')
        ws_daily.write(0, 0, 'Projection Type')
        ws_daily.write(0, 1, 'BTC Amount')
        ws_daily.write(0, 2, 'USD Amount')
        
        row = 1
        ws_daily.write(row, 0, 'Gross Payout (before pool fee)')
        ws_daily.write(row, 1, round(daily_btc_gross_before_fee, 8))
        ws_daily.write(row, 2, round(daily_usd_gross, 2))
        row += 1
        
        ws_daily.write(row, 0, 'Pool Fee')
        ws_daily.write(row, 1, round(pool_fee_btc, 8))
        ws_daily.write(row, 2, round(pool_fee_btc * btc_price_usd, 2))
        row += 1
        
        ws_daily.write(row, 0, 'After Pool Fee')
        ws_daily.write(row, 1, round(daily_btc_after_fee, 8))
        ws_daily.write(row, 2, round(daily_usd_after_fee, 2))
        row += 1
        
        ws_daily.write(row, 0, 'Electricity Cost')
        ws_daily.write(row, 1, round(daily_electricity_cost_usd / btc_price_usd, 8) if btc_price_usd > 0 else 0)
        ws_daily.write(row, 2, round(daily_electricity_cost_usd, 2))
        row += 1
        
        ws_daily.write(row, 0, 'Net Profit')
        ws_daily.write(row, 1, round(daily_btc_net, 8))
        ws_daily.write(row, 2, round(daily_usd_net, 2))
        row += 1
        
        # Sheet 3: Monthly Projections
        ws_monthly = wb.add_sheet('Monthly Projections')
        ws_monthly.write(0, 0, 'Projection Type')
        ws_monthly.write(0, 1, 'BTC Amount')
        ws_monthly.write(0, 2, 'USD Amount')
        
        row = 1
        ws_monthly.write(row, 0, 'Gross Payout (before pool fee)')
        ws_monthly.write(row, 1, round(daily_btc_gross_before_fee * 30, 8))
        ws_monthly.write(row, 2, round(daily_usd_gross * 30, 2))
        row += 1
        
        ws_monthly.write(row, 0, 'Pool Fee')
        ws_monthly.write(row, 1, round(pool_fee_btc * 30, 8))
        ws_monthly.write(row, 2, round(pool_fee_btc * btc_price_usd * 30, 2))
        row += 1
        
        ws_monthly.write(row, 0, 'After Pool Fee')
        ws_monthly.write(row, 1, round(daily_btc_after_fee * 30, 8))
        ws_monthly.write(row, 2, round(daily_usd_after_fee * 30, 2))
        row += 1
        
        ws_monthly.write(row, 0, 'Electricity Cost')
        ws_monthly.write(row, 1, round(daily_electricity_cost_usd / btc_price_usd * 30, 8) if btc_price_usd > 0 else 0)
        ws_monthly.write(row, 2, round(daily_electricity_cost_usd * 30, 2))
        row += 1
        
        ws_monthly.write(row, 0, 'Net Profit')
        ws_monthly.write(row, 1, round(daily_btc_net * 30, 8))
        ws_monthly.write(row, 2, round(daily_usd_net * 30, 2))
        row += 1
        
        # Sheet 4: Yearly Projections
        ws_yearly = wb.add_sheet('Yearly Projections')
        ws_yearly.write(0, 0, 'Projection Type')
        ws_yearly.write(0, 1, 'BTC Amount')
        ws_yearly.write(0, 2, 'USD Amount')
        
        row = 1
        ws_yearly.write(row, 0, 'Gross Payout (before pool fee)')
        ws_yearly.write(row, 1, round(daily_btc_gross_before_fee * 365, 8))
        ws_yearly.write(row, 2, round(daily_usd_gross * 365, 2))
        row += 1
        
        ws_yearly.write(row, 0, 'Pool Fee')
        ws_yearly.write(row, 1, round(pool_fee_btc * 365, 8))
        ws_yearly.write(row, 2, round(pool_fee_btc * btc_price_usd * 365, 2))
        row += 1
        
        ws_yearly.write(row, 0, 'After Pool Fee')
        ws_yearly.write(row, 1, round(daily_btc_after_fee * 365, 8))
        ws_yearly.write(row, 2, round(daily_usd_after_fee * 365, 2))
        row += 1
        
        ws_yearly.write(row, 0, 'Electricity Cost')
        ws_yearly.write(row, 1, round(daily_electricity_cost_usd / btc_price_usd * 365, 8) if btc_price_usd > 0 else 0)
        ws_yearly.write(row, 2, round(daily_electricity_cost_usd * 365, 2))
        row += 1
        
        ws_yearly.write(row, 0, 'Net Profit')
        ws_yearly.write(row, 1, round(daily_btc_net * 365, 8))
        ws_yearly.write(row, 2, round(daily_usd_net * 365, 2))
        row += 1
        
        # Sheet 5: Cost Basis Analysis
        ws_cost = wb.add_sheet('Cost Basis Analysis')
        ws_cost.write(0, 0, 'Cost Analysis')
        ws_cost.write(0, 1, 'Value')
        ws_cost.write(0, 2, 'Unit')
        
        row = 1
        ws_cost.write(row, 0, 'Market Price per BTC')
        ws_cost.write(row, 1, round(btc_price_usd, 2))
        ws_cost.write(row, 2, 'USD')
        row += 1
        
        ws_cost.write(row, 0, 'My Cost Basis per BTC')
        ws_cost.write(row, 1, round(cost_basis_usd_per_btc, 2))
        ws_cost.write(row, 2, 'USD')
        row += 1
        
        ws_cost.write(row, 0, 'Discount vs Market')
        ws_cost.write(row, 1, round(discount_vs_market_pct, 2))
        ws_cost.write(row, 2, '%')
        row += 1
        
        # Sheet 6: Energy Metrics
        ws_energy = wb.add_sheet('Energy Metrics')
        ws_energy.write(0, 0, 'Energy Metric')
        ws_energy.write(0, 1, 'Value')
        ws_energy.write(0, 2, 'Unit')
        
        row = 1
        ws_energy.write(row, 0, 'Power Consumption')
        ws_energy.write(row, 1, round(power_watts / 1000, 2))
        ws_energy.write(row, 2, 'kW')
        row += 1
        
        ws_energy.write(row, 0, 'Daily Energy Usage')
        ws_energy.write(row, 1, round(daily_energy_kwh, 2))
        ws_energy.write(row, 2, 'kWh')
        row += 1
        
        ws_energy.write(row, 0, 'Electricity Price')
        ws_energy.write(row, 1, round(price_per_kwh, 5))
        ws_energy.write(row, 2, '$/kWh')
        row += 1
        
        ws_energy.write(row, 0, 'Mining Efficiency')
        ws_energy.write(row, 1, round(efficiency_w_th, 2))
        ws_energy.write(row, 2, 'W/TH')
        row += 1
        
        ws_energy.write(row, 0, 'Energy to Mining Ratio')
        ws_energy.write(row, 1, round(energy_cost_percentage, 2))
        ws_energy.write(row, 2, '%')
        row += 1
    
    response = HttpResponse(content_type='application/vnd.ms-excel')
    platform_suffix = f'_{selected_platform_name.replace(" ", "_")}' if selected_platform else ''
    response['Content-Disposition'] = f'attachment; filename="forecasting_dashboard{platform_suffix}_export.xls"'
    wb.save(response)
    return response
