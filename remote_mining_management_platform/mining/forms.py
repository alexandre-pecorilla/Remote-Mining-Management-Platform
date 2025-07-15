from django import forms
from .models import RemoteMiningPlatform


class RemoteMiningPlatformForm(forms.ModelForm):
    class Meta:
        model = RemoteMiningPlatform
        fields = ['name', 'website_link', 'portal_url', 'logo', 'point_of_contact_name', 
                 'point_of_contact_email', 'point_of_contact_phone', 'point_of_contact_telegram', 'energy_price']
        widgets = {
            'name': forms.TextInput(attrs={'class': 'form-control', 'required': True}),
            'website_link': forms.URLInput(attrs={'class': 'form-control'}),
            'portal_url': forms.URLInput(attrs={'class': 'form-control'}),
            'logo': forms.FileInput(attrs={'class': 'form-control'}),
            'point_of_contact_name': forms.TextInput(attrs={'class': 'form-control'}),
            'point_of_contact_email': forms.EmailInput(attrs={'class': 'form-control'}),
            'point_of_contact_phone': forms.TextInput(attrs={'class': 'form-control'}),
            'point_of_contact_telegram': forms.TextInput(attrs={'class': 'form-control'}),
            'energy_price': forms.NumberInput(attrs={
                'class': 'form-control', 
                'step': '0.0001'
            }),
        }
        labels = {
            'name': 'Platform Name',
            'website_link': 'Website URL',
            'portal_url': 'Portal',
            'logo': 'Platform Logo',
            'point_of_contact_name': 'Contact Name',
            'point_of_contact_email': 'Contact Email',
            'point_of_contact_phone': 'Contact Phone',
            'point_of_contact_telegram': 'Telegram Username',
            'energy_price': 'Energy Price ($/kWh)',
        }
