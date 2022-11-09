from django import forms
from django.forms import ModelForm
from .models import (
    Customer, Hall, Sender, Product, Document,
    PersonInCharge, INPUT_FORMATS,
)


class CustomerForm(ModelForm):
    class Meta:
        model = Customer
        fields = '__all__'


class HallForm(ModelForm):
    class Meta:
        model = Hall
        fields = '__all__'


class SenderForm(ModelForm):
    class Meta:
        model = Sender
        fields = '__all__'


class ProductForm(ModelForm):
    purchase_date = forms.DateField(input_formats=INPUT_FORMATS, required=False)
    class Meta:
        model = Product
        fields = '__all__'


class DocumentForm(ModelForm):
    class Meta:
        model = Document
        fields = '__all__'


class PersonInChargeForm(ModelForm):
    class Meta:
        model = PersonInCharge
        fields = '__all__'
