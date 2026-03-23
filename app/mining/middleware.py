from django.conf import settings
from django.shortcuts import redirect
from django.urls import reverse


class PasswordProtectionMiddleware:
    """Require a password to access the app when APP_PASSWORD is set."""

    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        app_password = getattr(settings, 'APP_PASSWORD', '')

        if app_password and not request.session.get('app_authenticated'):
            login_url = reverse('app_login')
            if request.path != login_url and not request.path.startswith('/admin/'):
                return redirect(f'{login_url}?next={request.get_full_path()}')

        return self.get_response(request)
