from django.db import models
from django.urls import reverse


class RemoteMiningPlatform(models.Model):
    name = models.CharField(max_length=200)
    website_link = models.URLField(blank=True, null=True)
    portal_url = models.URLField(blank=True, null=True)
    logo = models.ImageField(upload_to='platform_logos/', blank=True, null=True)
    point_of_contact_name = models.CharField(max_length=100, blank=True, null=True)
    point_of_contact_email = models.EmailField(blank=True, null=True)
    point_of_contact_phone = models.CharField(max_length=20, blank=True, null=True)
    point_of_contact_telegram = models.CharField(max_length=50, blank=True, null=True)
    energy_price = models.DecimalField(max_digits=6, decimal_places=4, blank=True, null=True, help_text="Energy price in $/kWh")
    
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    
    class Meta:
        ordering = ['name']
        verbose_name = 'Remote Mining Platform'
        verbose_name_plural = 'Remote Mining Platforms'
    
    def __str__(self):
        return self.name
    
    def get_absolute_url(self):
        return reverse('platform_detail', kwargs={'pk': self.pk})
    
    def formatted_energy_price(self):
        if self.energy_price:
            return f"${self.energy_price:.4f}/kWh"
        return "Not specified"


class Miner(models.Model):
    model = models.CharField(max_length=200)
    image = models.ImageField(upload_to='miner_images/', blank=True, null=True)
    manufacturer = models.CharField(max_length=100, blank=True, null=True)
    product_link = models.URLField(blank=True, null=True)
    serial_number = models.CharField(max_length=100, blank=True, null=True)
    platform = models.ForeignKey(RemoteMiningPlatform, on_delete=models.SET_NULL, blank=True, null=True, related_name='miners')
    platform_internal_id = models.CharField(max_length=100, blank=True, null=True, help_text="Internal platform ID")
    hashrate = models.DecimalField(max_digits=10, decimal_places=3, blank=True, null=True, help_text="TH/s")
    power = models.DecimalField(max_digits=8, decimal_places=3, blank=True, null=True, help_text="kW")
    efficiency = models.DecimalField(max_digits=8, decimal_places=2, blank=True, null=True, help_text="W/TH")
    purchase_price = models.DecimalField(max_digits=12, decimal_places=2, blank=True, null=True, help_text="USD")
    purchase_date = models.DateField(blank=True, null=True)
    start_date = models.DateField(blank=True, null=True)
    location = models.CharField(max_length=200, blank=True, null=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        return self.model

    @property
    def energy_price(self):
        """Get energy price from associated platform"""
        return self.platform.energy_price if self.platform else None

    @property
    def formatted_energy_price(self):
        """Format energy price from platform"""
        return f"${self.energy_price:.4f}/kWh" if self.energy_price else None

    @property
    def formatted_purchase_price(self):
        """Format purchase price"""
        return f"${self.purchase_price:,.2f}" if self.purchase_price else None

    @property
    def formatted_hashrate(self):
        """Format hashrate to show integers without decimals, but keep decimals when needed"""
        if self.hashrate:
            # If the hashrate is a whole number, display as integer
            if self.hashrate == int(self.hashrate):
                return f"{int(self.hashrate)}"
            else:
                return f"{self.hashrate:.3f}".rstrip('0').rstrip('.')
        return None


class Payout(models.Model):
    """Mining payout record"""
    payout_date = models.DateField(help_text="Date of the payout")
    payout_amount = models.DecimalField(max_digits=12, decimal_places=8, help_text="Payout amount in BTC")
    platform = models.ForeignKey(RemoteMiningPlatform, on_delete=models.SET_NULL, blank=True, null=True, related_name='payouts')
    transaction_id = models.CharField(max_length=100, blank=True, null=True, help_text="Bitcoin transaction ID")
    closing_price = models.DecimalField(max_digits=12, decimal_places=2, blank=True, null=True, help_text="Bitcoin closing price in USD at payout date")
    value_at_payout = models.DecimalField(max_digits=15, decimal_places=2, blank=True, null=True, help_text="USD value at payout (payout_amount * closing_price)")
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Payout"
        verbose_name_plural = "Payouts"
        ordering = ['-payout_date']

    def __str__(self):
        return f"Payout {self.payout_amount} BTC on {self.payout_date}"

    def save(self, *args, **kwargs):
        """Override save to automatically calculate value_at_payout when closing_price exists"""
        if self.closing_price and self.payout_amount:
            self.value_at_payout = self.payout_amount * self.closing_price
        elif not self.closing_price:
            self.value_at_payout = None
        super().save(*args, **kwargs)

    @property
    def current_market_value(self):
        """Calculate current market value in USD using Bitcoin price from API data"""
        from .models import APIData
        api_data = APIData.get_api_data()
        if api_data.bitcoin_price_usd:
            return float(self.payout_amount) * float(api_data.bitcoin_price_usd)
        return None

    @property
    def mempool_link(self):
        """Generate mempool.space link for transaction ID"""
        return f"https://mempool.space/tx/{self.transaction_id}"


class APIData(models.Model):
    """API data from external sources - singleton model"""
    bitcoin_price_usd = models.DecimalField(max_digits=12, decimal_places=2, blank=True, null=True, help_text="Bitcoin Price in USD")
    network_hashrate_ehs = models.DecimalField(max_digits=15, decimal_places=3, blank=True, null=True, help_text="Network Hashrate in EH/s")
    network_difficulty = models.BigIntegerField(blank=True, null=True, help_text="Network Difficulty")
    avg_block_fees_24h = models.DecimalField(max_digits=12, decimal_places=8, blank=True, null=True, help_text="24h Average Block Fees in BTC")
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "API Data"
        verbose_name_plural = "API Data"

    def __str__(self):
        return "API Data"

    @classmethod
    def get_api_data(cls):
        """Get or create singleton API data instance"""
        api_data, created = cls.objects.get_or_create(pk=1)
        return api_data


class Settings(models.Model):
    """Application settings - singleton model"""
    coinmarketcap_api_key = models.CharField(max_length=255, blank=True, null=True, help_text="CoinMarketCap API Key for Bitcoin price data")
    dark_mode = models.BooleanField(default=False, help_text="Enable dark mode theme")
    developer_mode = models.BooleanField(default=False, help_text="Enable developer mode to show database IDs")
    pool_fee_percentage = models.DecimalField(max_digits=5, decimal_places=2, default=2.5, help_text="Pool fee percentage (e.g. 2.5 for 2.5%)")
    block_reward = models.DecimalField(max_digits=10, decimal_places=8, default=3.125, help_text="Bitcoin block reward (BTC per block)")
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Settings"
        verbose_name_plural = "Settings"

    def __str__(self):
        return "Application Settings"

    @classmethod
    def get_settings(cls):
        """Get or create singleton settings instance"""
        settings, created = cls.objects.get_or_create(pk=1)
        return settings
