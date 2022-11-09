from django.forms import ModelForm
from django import forms
from masterdata.models import Product, INPUT_FORMATS


class ProductForm(ModelForm):
    purchase_date = forms.DateField(input_formats=INPUT_FORMATS)
    class Meta:
        model = Product
        fields = '__all__'


class ListingSearchForm(forms.Form):
    contract_id = forms.CharField(required=False)
    created_at = forms.DateField(input_formats=INPUT_FORMATS, required=False)
    created_at_to = forms.DateField(input_formats=INPUT_FORMATS, required=False)
    shipping_date_from = forms.DateField(input_formats=INPUT_FORMATS, required=False)
    shipping_date_to = forms.DateField(input_formats=INPUT_FORMATS, required=False)
    customer = forms.CharField(required=False)
    name = forms.CharField(required=False)
    person_in_charge = forms.CharField(required=False)
    inventory_status = forms.CharField(required=False)
    

class ListingLinkSearchForm(forms.Form):
    p_contract_id = forms.CharField(required=False)
    p_created_at = forms.DateField(input_formats=INPUT_FORMATS, required=False)
    p_created_at_to = forms.DateField(input_formats=INPUT_FORMATS, required=False)
    p_shipping_date_from = forms.DateField(input_formats=INPUT_FORMATS, required=False)
    p_shipping_date_to = forms.DateField(input_formats=INPUT_FORMATS, required=False)
    p_customer = forms.CharField(required=False)
    p_name = forms.CharField(required=False)
    p_person_in_charge = forms.CharField(required=False)
    p_inventory_status = forms.CharField(required=False)

    s_contract_id = forms.CharField(required=False)
    s_created_at = forms.DateField(input_formats=INPUT_FORMATS, required=False)
    s_created_at_to = forms.DateField(input_formats=INPUT_FORMATS, required=False)
    s_shipping_date_from = forms.DateField(input_formats=INPUT_FORMATS, required=False)
    s_shipping_date_to = forms.DateField(input_formats=INPUT_FORMATS, required=False)
    s_customer = forms.CharField(required=False)
    s_name = forms.CharField(required=False)
    s_person_in_charge = forms.CharField(required=False)
    s_inventory_status = forms.CharField(required=False)

    w_name = forms.CharField(required=False)
    w_person_in_charge = forms.CharField(required=False)
    w_inventory_status = forms.CharField(required=False)

class ListingLinkShowSearchForm(forms.Form):
    contract_id = forms.CharField(required=False)
    created_at = forms.DateField(input_formats=INPUT_FORMATS, required=False)
    created_at_to = forms.DateField(input_formats=INPUT_FORMATS, required=False)
    shipping_date_from = forms.DateField(input_formats=INPUT_FORMATS, required=False)
    shipping_date_to = forms.DateField(input_formats=INPUT_FORMATS, required=False)
    person_in_charge = forms.CharField(required=False)
    inventory_status = forms.CharField(required=False)
    report_date = forms.DateField(input_formats=INPUT_FORMATS, required=False)
