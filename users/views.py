from django.shortcuts import render, redirect
from django.urls import reverse
from django.contrib.auth import authenticate, login, logout
from django.contrib.auth.mixins import LoginRequiredMixin, UserPassesTestMixin
from django.contrib import messages
from django.views.generic import View, RedirectView
from users.forms import LoginForm


class AdminCheckMixin:
    def is_admin(self, user):
        if user.is_authenticated and user.is_superuser and user.is_staff:
            return True
        return False


class AdminLoginRequiredMixin(LoginRequiredMixin, UserPassesTestMixin, AdminCheckMixin):
    def test_func(self):
        return self.is_admin(self.request.user)
    
    def handle_no_permission(self):
        next = self.request.get_full_path()
        return redirect(reverse('login') + '?next=' + next)


class LoginView(AdminCheckMixin, View):
    template_name = 'registration/login.html'

    def get(self, request, *args, **kwargs):
        if self.is_admin(request.user):
            return redirect('dashboard')
        return render(request, self.template_name)
    
    def post(self, request, *args, **kwargs):
        form = LoginForm(request.POST)
        if form.is_valid():
            username = request.POST.get('username')
            password = request.POST.get('password')
            user = authenticate(request, username=username, password=password)
            if user and self.is_admin(user):
                login(request, user)
                redirect_url = request.GET.get('next', 'dashboard')
                return redirect(redirect_url)
        messages.error(request, 'Username or password is incorrect.')
        return render(request, self.template_name)


class LogoutView(AdminLoginRequiredMixin, RedirectView):
    pattern_name = 'login'

    def get_redirect_url(self, *args, **kwargs):
        logout(self.request)
        return super().get_redirect_url(*args, **kwargs)


class DashboardView(AdminLoginRequiredMixin, RedirectView):
    pattern_name = 'masterdata:customer-list'


def page_not_found(request, exception):
    # If onlye - return render(request, 'error/404.html'), it works on local/server with debug off. but it won't work with debug on.
    # return render(request, 'error/404.html')
    # render(request, 'error/404.html') returns HttpResponse, but its status code is 200. that's why it should be done as below.
    response = render(request, 'error/404.html', {})
    response.status_code = 404
    return response
    

def internal_error(request):
    response = render(request, 'error/500.html', {})
    response.status_code = 500
    return response
