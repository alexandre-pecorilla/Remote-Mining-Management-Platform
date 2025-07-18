from django.urls import path
from . import views

urlpatterns = [
    path('', views.PlatformListView.as_view(), name='platform_list'),
    path('platform/<int:pk>/', views.PlatformDetailView.as_view(), name='platform_detail'),
    path('platform/create/', views.PlatformCreateView.as_view(), name='platform_create'),
    path('platforms/<int:pk>/edit/', views.PlatformUpdateView.as_view(), name='platform_edit'),
    path('platforms/<int:pk>/delete/', views.PlatformDeleteView.as_view(), name='platform_delete'),
    
    # Miner URLs
    path('miners/', views.MinerListView.as_view(), name='miner_list'),
    path('miners/add/', views.MinerCreateView.as_view(), name='miner_create'),
    path('miners/<int:pk>/', views.MinerDetailView.as_view(), name='miner_detail'),
    path('miners/<int:pk>/edit/', views.MinerUpdateView.as_view(), name='miner_edit'),
    path('miners/<int:pk>/delete/', views.MinerDeleteView.as_view(), name='miner_delete'),
    
    # Payout URLs
    path('payouts/', views.PayoutListView.as_view(), name='payout_list'),
    path('payouts/<int:pk>/', views.PayoutDetailView.as_view(), name='payout_detail'),
    path('payouts/add/', views.PayoutCreateView.as_view(), name='payout_add'),
    path('payouts/<int:pk>/edit/', views.PayoutUpdateView.as_view(), name='payout_edit'),
    path('payouts/<int:pk>/delete/', views.PayoutDeleteView.as_view(), name='payout_delete'),
    
    # API Data
    path('api-data/', views.api_data_view, name='api_data'),
    
    # Settings
    path('settings/', views.settings_view, name='settings'),
    
    # Dashboards
    path('dashboard/overview/', views.overview_dashboard, name='overview_dashboard'),
    
    # Import Template Downloads
    path('download-templates/platform/', views.download_platform_template, name='download_platform_template'),
    path('download-templates/miner/', views.download_miner_template, name='download_miner_template'),
    path('download-templates/payout/', views.download_payout_template, name='download_payout_template'),
    
    # Data Import
    path('import-data/platform/', views.import_platform_data, name='import_platform_data'),
    path('import-data/miner/', views.import_miner_data, name='import_miner_data'),
    path('import-data/payout/', views.import_payout_data, name='import_payout_data'),
    
    # Data Export
    path('export-data/platform/', views.export_platform_data, name='export_platform_data'),
    path('export-data/miner/', views.export_miner_data, name='export_miner_data'),
    path('export-data/payout/', views.export_payout_data, name='export_payout_data'),
]
