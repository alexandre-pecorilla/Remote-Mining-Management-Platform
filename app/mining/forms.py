from django import forms
from .models import RemoteMiningPlatform, Miner, Settings, Payout, Expense, TopUp


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


class MinerForm(forms.ModelForm):
    class Meta:
        model = Miner
        fields = ['model', 'image', 'manufacturer', 'product_link', 'serial_number', 
                 'platform', 'platform_internal_id', 'hashrate', 'power', 'efficiency', 
                 'purchase_price', 'purchase_date', 'start_date', 'location']
        widgets = {
            'model': forms.TextInput(attrs={'class': 'form-control', 'required': True}),
            'image': forms.FileInput(attrs={'class': 'form-control'}),
            'manufacturer': forms.TextInput(attrs={'class': 'form-control'}),
            'product_link': forms.URLInput(attrs={'class': 'form-control'}),
            'serial_number': forms.TextInput(attrs={'class': 'form-control'}),
            'platform': forms.Select(attrs={'class': 'form-control'}),
            'platform_internal_id': forms.TextInput(attrs={'class': 'form-control'}),
            'hashrate': forms.NumberInput(attrs={'class': 'form-control', 'step': '0.001'}),
            'power': forms.NumberInput(attrs={'class': 'form-control', 'step': '0.001'}),
            'efficiency': forms.NumberInput(attrs={'class': 'form-control', 'step': '0.01'}),
            'purchase_price': forms.NumberInput(attrs={'class': 'form-control', 'step': '0.01'}),
            'purchase_date': forms.DateInput(attrs={'class': 'form-control', 'type': 'date'}),
            'start_date': forms.DateInput(attrs={'class': 'form-control', 'type': 'date'}),
            'location': forms.TextInput(attrs={'class': 'form-control'}),
        }
        labels = {
            'model': 'Miner Model',
            'image': 'Miner Image',
            'manufacturer': 'Manufacturer',
            'product_link': 'Product Link',
            'serial_number': 'Serial Number',
            'platform': 'Platform',
            'platform_internal_id': 'Platform ID',
            'hashrate': 'Hashrate (TH/s)',
            'power': 'Power (kW)',
            'efficiency': 'Efficiency (W/TH)',
            'purchase_price': 'Purchase Price ($)',
            'purchase_date': 'Purchase Date',
            'start_date': 'Start Date',
            'location': 'Location',
        }


class PayoutForm(forms.ModelForm):
    class Meta:
        model = Payout
        fields = ['payout_date', 'payout_amount', 'platform', 'transaction_id', 'closing_price']
        widgets = {
            'payout_date': forms.DateInput(attrs={
                'type': 'date',
                'class': 'form-control'
            }),
            'payout_amount': forms.NumberInput(attrs={
                'class': 'form-control',
                'step': '0.00000001',
                'min': '0'
            }),
            'platform': forms.Select(attrs={
                'class': 'form-control'
            }),
            'transaction_id': forms.TextInput(attrs={
                'class': 'form-control',
                'maxlength': '100'
            }),
            'closing_price': forms.NumberInput(attrs={
                'class': 'form-control',
                'step': '0.01',
                'min': '0'
            }),
        }
        labels = {
            'payout_date': 'Payout Date',
            'payout_amount': 'Payout Amount (BTC)',
            'platform': 'Platform',
            'transaction_id': 'Transaction ID',
            'closing_price': 'Closing Price ($)',
        }


class ExpenseForm(forms.ModelForm):
    class Meta:
        model = Expense
        fields = ['expense_date', 'platform', 'category', 'description', 'expense_amount', 'invoice_link', 'receipt_link', 'notes']
        widgets = {
            'expense_date': forms.DateInput(attrs={
                'type': 'date',
                'class': 'form-control'
            }),
            'platform': forms.Select(attrs={
                'class': 'form-control'
            }),
            'category': forms.Select(attrs={
                'class': 'form-control'
            }),
            'description': forms.TextInput(attrs={
                'class': 'form-control',
                'maxlength': '200'
            }),
            'expense_amount': forms.NumberInput(attrs={
                'class': 'form-control',
                'step': '0.01',
                'min': '0'
            }),
            'invoice_link': forms.URLInput(attrs={
                'class': 'form-control'
            }),
            'receipt_link': forms.URLInput(attrs={
                'class': 'form-control'
            }),
            'notes': forms.Textarea(attrs={
                'class': 'form-control',
                'rows': '4'
            }),
        }
        labels = {
            'expense_date': 'Expense Date',
            'platform': 'Platform',
            'category': 'Category',
            'description': 'Description',
            'expense_amount': 'Expense Amount ($)',
            'invoice_link': 'Invoice Link',
            'receipt_link': 'Receipt Link',
            'notes': 'Notes',
        }


class TopUpForm(forms.ModelForm):
    class Meta:
        model = TopUp
        fields = ['topup_date', 'platform', 'topup_amount', 'description', 'receipt_link']
        widgets = {
            'topup_date': forms.DateInput(attrs={
                'type': 'date',
                'class': 'form-control'
            }),
            'platform': forms.Select(attrs={
                'class': 'form-control'
            }),
            'topup_amount': forms.NumberInput(attrs={
                'class': 'form-control',
                'step': '0.01',
                'min': '0'
            }),
            'description': forms.Textarea(attrs={
                'class': 'form-control',
                'rows': '4'
            }),
            'receipt_link': forms.URLInput(attrs={
                'class': 'form-control'
            }),
        }
        labels = {
            'topup_date': 'Top-Up Date',
            'platform': 'Platform',
            'topup_amount': 'Top-Up Amount ($)',
            'description': 'Description',
            'receipt_link': 'Receipt Link',
        }


class SettingsForm(forms.ModelForm):
    class Meta:
        model = Settings
        fields = ['developer_mode', 'pool_fee_percentage', 'block_reward']
        widgets = {
            'developer_mode': forms.CheckboxInput(attrs={
                'class': 'form-check-input'
            }),
            'pool_fee_percentage': forms.NumberInput(attrs={
                'class': 'form-control',
                'step': '0.01',
                'min': '0',
                'max': '100'
            }),
            'block_reward': forms.NumberInput(attrs={
                'class': 'form-control',
                'step': '0.00000001',
                'min': '0',
                'max': '50'
            }),
        }
        labels = {
            'developer_mode': 'Developer Mode',
            'pool_fee_percentage': 'Pool Fee Percentage (%)',
            'block_reward': 'Block Reward (BTC)',
        }
