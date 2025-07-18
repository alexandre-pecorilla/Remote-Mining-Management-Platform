from django.shortcuts import render, get_object_or_404, redirect
from django.contrib import messages
from django.http import HttpResponse
from django.urls import reverse_lazy
from django.views.generic import ListView, DetailView, CreateView, UpdateView, DeleteView
import xlwt
import xlrd
from decimal import Decimal
from datetime import datetime
import json
from .models import RemoteMiningPlatform, Miner, Settings, APIData, Payout
from .forms import RemoteMiningPlatformForm, MinerForm, SettingsForm, PayoutForm
from .api_utils import fetch_all_api_data


# Home Page View
def home_view(request):
    """Home page with navigation to all sections of the application"""
    return render(request, 'mining/home.html')


class PlatformListView(ListView):
    model = RemoteMiningPlatform
    template_name = 'mining/platform_list.html'
    context_object_name = 'platforms'
    paginate_by = 10


class PlatformDetailView(DetailView):
    model = RemoteMiningPlatform
    template_name = 'mining/platform_detail.html'
    context_object_name = 'platform'


class PlatformCreateView(CreateView):
    model = RemoteMiningPlatform
    form_class = RemoteMiningPlatformForm
    template_name = 'mining/platform_form.html'
    success_url = reverse_lazy('platform_list')
    
    def form_valid(self, form):
        messages.success(self.request, 'Platform created successfully.')
        return super().form_valid(form)


class PlatformUpdateView(UpdateView):
    model = RemoteMiningPlatform
    form_class = RemoteMiningPlatformForm
    template_name = 'mining/platform_form.html'
    success_url = reverse_lazy('platform_list')
    
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


class MinerCreateView(CreateView):
    model = Miner
    form_class = MinerForm
    template_name = 'mining/miner_form.html'
    success_url = reverse_lazy('miner_list')

    def form_valid(self, form):
        messages.success(self.request, "Miner created successfully!")
        return super().form_valid(form)


class MinerUpdateView(UpdateView):
    model = Miner
    form_class = MinerForm
    template_name = 'mining/miner_form.html'
    success_url = reverse_lazy('miner_list')

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


# Payout Views
class PayoutListView(ListView):
    model = Payout
    template_name = 'mining/payout_list.html'
    context_object_name = 'payouts'
    paginate_by = 10


class PayoutDetailView(DetailView):
    model = Payout
    template_name = 'mining/payout_detail.html'
    context_object_name = 'payout'


class PayoutCreateView(CreateView):
    model = Payout
    form_class = PayoutForm
    template_name = 'mining/payout_form.html'
    success_url = reverse_lazy('payout_list')
    
    def form_valid(self, form):
        messages.success(self.request, 'Payout added successfully!')
        return super().form_valid(form)


class PayoutUpdateView(UpdateView):
    model = Payout
    form_class = PayoutForm
    template_name = 'mining/payout_form.html'
    success_url = reverse_lazy('payout_list')
    
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


# Dashboard Views
def overview_dashboard(request):
    """Overview Dashboard with comprehensive mining analytics"""
    from django.db.models import Sum, Avg, Count
    from decimal import Decimal
    
    # Get API data
    api_data = APIData.get_api_data()
    
    # NETWORK DATA
    bitcoin_price = api_data.bitcoin_price_usd or 0
    network_hashrate = api_data.network_hashrate_ehs or 0
    network_difficulty = api_data.network_difficulty or 0
    avg_block_fees_24h = api_data.avg_block_fees_24h or 0
    
    # FLEET DATA
    miners = Miner.objects.filter(hashrate__isnull=False, power__isnull=False)
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
    for platform in RemoteMiningPlatform.objects.all():
        platform_hashrate = platform.miners.aggregate(total=Sum('hashrate'))['total'] or 0
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
    total_btc_mined = payouts.aggregate(total=Sum('payout_amount'))['total'] or 0
    current_gross_value = float(total_btc_mined) * float(bitcoin_price) if total_btc_mined and bitcoin_price else 0
    total_payouts = payouts.count()
    
    # REVENUES DISTRIBUTION DATA
    revenue_by_platform = []
    for platform in RemoteMiningPlatform.objects.all():
        platform_btc = platform.payouts.aggregate(total=Sum('payout_amount'))['total'] or 0
        platform_payouts = platform.payouts.count()
        if platform_btc > 0:
            platform_value = float(platform_btc) * float(bitcoin_price) if bitcoin_price else 0
            revenue_by_platform.append({
                'platform': platform.name,
                'btc_mined': float(platform_btc),
                'gross_value': platform_value,
                'payout_count': platform_payouts
            })
    
    context = {
        # Network Data
        'bitcoin_price': bitcoin_price,
        'network_hashrate': network_hashrate,
        'network_difficulty': network_difficulty,
        'avg_block_fees_24h': avg_block_fees_24h,
        
        # Fleet Data
        'miner_count': miner_count,
        'total_hashrate': total_hashrate,
        'total_power': total_power,
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
    headers = ['payout_date', 'payout_amount', 'platform', 'transaction_id']
    
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
    headers = ['payout_date', 'payout_amount', 'platform', 'transaction_id']
    
    for col, header in enumerate(headers):
        ws.write(0, col, header)
    
    # Add data rows
    payouts = Payout.objects.all()
    for row, payout in enumerate(payouts, start=1):
        ws.write(row, 0, payout.payout_date.strftime('%Y-%m-%d') if payout.payout_date else '')
        ws.write(row, 1, float(payout.payout_amount) if payout.payout_amount else '')
        ws.write(row, 2, payout.platform.pk if payout.platform else '')
        ws.write(row, 3, payout.transaction_id or '')
    
    response = HttpResponse(
        content_type='application/vnd.ms-excel'
    )
    response['Content-Disposition'] = 'attachment; filename="payout_data_export.xls"'
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
