import django_filters
from django import forms
from django.utils.translation import gettext_lazy as _
from django.db.models import Q
from .models import Customer, Hall, Sender, Product


class CustomerFilter(django_filters.FilterSet):
    keyword = django_filters.CharFilter(method='filter_by_keyword', widget=forms.TextInput(attrs={'class': 'form-control'}))

    class Meta:
        model = Customer
        fields = '__all__'
        
    def filter_by_keyword(self, queryset, name, value):
        queryset = queryset.filter(
            Q(name__icontains=value) |
            Q(frigana__icontains=value) |
            Q(tel__icontains=value) |
            Q(fax__icontains=value) |
            Q(address__icontains=value) |
            Q(excel__icontains=value)
        )
        return queryset


class HallFilter(django_filters.FilterSet):
    keyword = django_filters.CharFilter(method='filter_by_keyword', widget=forms.TextInput(attrs={'class': 'form-control'}))

    class Meta:
        model = Hall
        fields = '__all__'
        
    def filter_by_keyword(self, queryset, name, value):
        queryset = queryset.filter(
            Q(customer_name__icontains=value) |
            Q(customer_frigana__icontains=value) |
            Q(name__icontains=value) |
            Q(frigana__icontains=value) |
            Q(tel__icontains=value) |
            Q(fax__icontains=value) |
            Q(address__icontains=value) |
            Q(payee__icontains=value)
        )
        return queryset


class SenderFilter(django_filters.FilterSet):
    keyword = django_filters.CharFilter(method='filter_by_keyword', widget=forms.TextInput(attrs={'class': 'form-control'}))

    class Meta:
        model = Sender
        fields = '__all__'
        
    def filter_by_keyword(self, queryset, name, value):
        queryset = queryset.filter(
            Q(name__icontains=value) |
            Q(frigana__icontains=value) |
            Q(tel__icontains=value) |
            Q(fax__icontains=value) |
            Q(address__icontains=value)
        )
        return queryset


class ProductFilter(django_filters.FilterSet):
    name = django_filters.CharFilter(lookup_expr='icontains', widget=forms.TextInput(attrs={'class': 'form-control'}))

    class Meta:
        model = Product
        fields = ('name',)
