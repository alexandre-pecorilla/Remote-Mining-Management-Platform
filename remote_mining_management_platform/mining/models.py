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
