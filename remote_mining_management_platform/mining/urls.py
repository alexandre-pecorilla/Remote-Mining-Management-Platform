from django.urls import path
from . import views

urlpatterns = [
    path('', views.PlatformListView.as_view(), name='platform_list'),
    path('platform/<int:pk>/', views.PlatformDetailView.as_view(), name='platform_detail'),
    path('platform/create/', views.PlatformCreateView.as_view(), name='platform_create'),
    path('platform/<int:pk>/edit/', views.PlatformUpdateView.as_view(), name='platform_edit'),
    path('platform/<int:pk>/delete/', views.PlatformDeleteView.as_view(), name='platform_delete'),
]
