from .models import Settings

def settings_context(request):
    """Add settings to template context"""
    try:
        settings = Settings.objects.first()
        if not settings:
            settings = Settings.objects.create()
        return {'settings': settings}
    except:
        return {'settings': None}
