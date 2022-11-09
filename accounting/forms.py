from django.forms import ModelForm
from django import forms
from masterdata.models import INPUT_FORMATS

class SearchForm(forms.Form):
    contract_id = forms.CharField(required=False)
    created_at = forms.DateField(input_formats=INPUT_FORMATS, required=False)
    created_at_to = forms.DateField(input_formats=INPUT_FORMATS, required=False)
    customer = forms.CharField(required=False)
    none_date = forms.CharField(required=False) 
