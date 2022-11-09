import django_filters
from django import forms
from django.utils.translation import gettext_lazy as _
from django.db.models import Q
from masterdata.models import Product, INPUT_FORMATS


class ProductFilter(django_filters.FilterSet):
    name = django_filters.CharFilter(lookup_expr='icontains', widget=forms.TextInput(attrs={'class': 'form-control'}))
    purchase_date = django_filters.DateFilter(
        input_formats=INPUT_FORMATS,
        widget=forms.TextInput(attrs={'class': 'form-control datepicker-nullable'}))
    supplier = django_filters.CharFilter(lookup_expr='icontains', widget=forms.TextInput(attrs={'class': 'form-control'}))
    person_in_charge = django_filters.CharFilter(widget=forms.TextInput(attrs={'class': 'form-control'}))
    quantity = django_filters.NumberFilter(widget=forms.TextInput(attrs={'class': 'form-control'}))

    class Meta:
        model = Product
        fields = ('name', 'purchase_date', 'supplier', 'person_in_charge', 'quantity')
