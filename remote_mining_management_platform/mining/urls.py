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
    
    # API Data
    path('api-data/', views.api_data_view, name='api_data'),
    
    # Settings
    path('settings/', views.settings_view, name='settings'),
]
