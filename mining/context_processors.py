import logging
from django.conf import settings as django_settings
from .models import Settings

logger = logging.getLogger(__name__)

def settings_context(request):
    """Add settings to template context"""
    try:
        settings = Settings.objects.first()
        if not settings:
            settings = Settings.objects.create()
        return {
            'settings': settings,
            'cmc_api_key_configured': bool(django_settings.COINMARKETCAP_API_KEY),
        }
    except Exception:
        logger.exception("Failed to load settings context")
        return {'settings': None, 'cmc_api_key_configured': False}
