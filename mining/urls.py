from django.urls import path
from . import views

urlpatterns = [
    # Home Page
    path('', views.home_view, name='home'),
    
    # CAPEX/OPEX Dashboard
    path('dashboard/capex-opex/', views.capex_opex_dashboard, name='capex_opex_dashboard'),
    path('export-data/capex-opex/', views.export_capex_opex_data, name='export_capex_opex_data'),
    
    # Platform URLs - /data/platforms
    path('data/platforms/', views.PlatformListView.as_view(), name='platform_list'),
    path('data/platforms/<int:pk>/', views.PlatformDetailView.as_view(), name='platform_detail'),
    path('data/platforms/create/', views.PlatformCreateView.as_view(), name='platform_create'),
    path('data/platforms/<int:pk>/edit/', views.PlatformUpdateView.as_view(), name='platform_edit'),
    path('data/platforms/<int:pk>/delete/', views.PlatformDeleteView.as_view(), name='platform_delete'),
    
    # Miner URLs - /data/miners
    path('data/miners/', views.MinerListView.as_view(), name='miner_list'),
    path('data/miners/add/', views.MinerCreateView.as_view(), name='miner_create'),
    path('data/miners/<int:pk>/', views.MinerDetailView.as_view(), name='miner_detail'),
    path('data/miners/<int:pk>/edit/', views.MinerUpdateView.as_view(), name='miner_edit'),
    path('data/miners/<int:pk>/delete/', views.MinerDeleteView.as_view(), name='miner_delete'),
    
    # Payout URLs - /data/payouts
    path('data/payouts/', views.PayoutListView.as_view(), name='payout_list'),
    path('data/payouts/<int:pk>/', views.PayoutDetailView.as_view(), name='payout_detail'),
    path('data/payouts/add/', views.PayoutCreateView.as_view(), name='payout_add'),
    path('data/payouts/<int:pk>/edit/', views.PayoutUpdateView.as_view(), name='payout_edit'),
    path('data/payouts/<int:pk>/delete/', views.PayoutDeleteView.as_view(), name='payout_delete'),
    path('data/payouts/<int:payout_id>/fetch-closing-price/', views.fetch_closing_price, name='fetch_closing_price'),
    
    # Expenses - /data/expenses
    path('data/expenses/', views.ExpenseListView.as_view(), name='expense_list'),
    path('data/expenses/<int:pk>/', views.ExpenseDetailView.as_view(), name='expense_detail'),
    path('data/expenses/add/', views.ExpenseCreateView.as_view(), name='expense_create'),
    path('data/expenses/<int:pk>/edit/', views.ExpenseUpdateView.as_view(), name='expense_edit'),
    path('data/expenses/<int:pk>/delete/', views.ExpenseDeleteView.as_view(), name='expense_delete'),
    
    # Top-Ups - /data/topups
    path('data/topups/', views.TopUpListView.as_view(), name='topup_list'),
    path('data/topups/<int:pk>/', views.TopUpDetailView.as_view(), name='topup_detail'),
    path('data/topups/add/', views.TopUpCreateView.as_view(), name='topup_create'),
    path('data/topups/<int:pk>/edit/', views.TopUpUpdateView.as_view(), name='topup_edit'),
    path('data/topups/<int:pk>/delete/', views.TopUpDeleteView.as_view(), name='topup_delete'),
    
    # API Data - /data/api-data
    path('data/api-data/', views.api_data_view, name='api_data'),
    
    # Settings
    path('settings/', views.settings_view, name='settings'),
    
    # Dashboards
    path('dashboard/overview/', views.overview_dashboard, name='overview_dashboard'),
    path('dashboard/forecasting/', views.forecasting_dashboard, name='forecasting_dashboard'),
    
    # Import Template Downloads
    path('download-templates/platform/', views.download_platform_template, name='download_platform_template'),
    path('download-templates/miner/', views.download_miner_template, name='download_miner_template'),
    path('download-templates/payout/', views.download_payout_template, name='download_payout_template'),
    path('download-templates/expense/', views.download_expense_template, name='download_expense_template'),
    path('download-templates/topup/', views.download_topup_template, name='download_topup_template'),
    
    # Data Import
    path('import-data/platform/', views.import_platform_data, name='import_platform_data'),
    path('import-data/miner/', views.import_miner_data, name='import_miner_data'),
    path('import-data/payout/', views.import_payout_data, name='import_payout_data'),
    path('import-data/expense/', views.import_expense_data, name='import_expense_data'),
    path('import-data/topup/', views.import_topup_data, name='import_topup_data'),
    
    # Data Export
    path('export-data/platform/', views.export_platform_data, name='export_platform_data'),
    path('export-data/miner/', views.export_miner_data, name='export_miner_data'),
    path('export-data/payout/', views.export_payout_data, name='export_payout_data'),
    path('export-data/expense/', views.export_expense_data, name='export_expense_data'),
    path('export-data/topup/', views.export_topup_data, name='export_topup_data'),
    path('export-data/overview/', views.export_overview_data, name='export_overview_data'),
    path('export-data/forecasting/', views.export_forecasting_data, name='export_forecasting_data'),
]
