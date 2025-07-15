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
