import logging
from .models import Settings

logger = logging.getLogger(__name__)

def settings_context(request):
    """Add settings to template context"""
    try:
        settings = Settings.objects.first()
        if not settings:
            settings = Settings.objects.create()
        return {'settings': settings}
    except Exception:
        logger.exception("Failed to load settings context")
        return {'settings': None}
