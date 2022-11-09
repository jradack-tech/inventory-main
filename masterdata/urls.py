from django.urls import path
from .views import (
    CustomerListView, CustomerDetailAjaxView, CustomerUpdateView, CustomerDeleteView,
    HallListView, HallDetailAjaxView, HallUpdateView, HallDeleteView, 
    SenderDetailAjaxView, SenderUpdateView, SenderDeleteView,
    ProductListView, ProductDetailAjaxView, ProductUpdateView, ProductDeleteView,
    DocumentListView, DocumentDetailAjaxView, DocumentUpdateView, DocumentDeleteView,
    PersonInChargeListView, PersonInChargeDetailAjaxView, PersonInChargeUpdateView, PersonInChargeDeleteView,
    CustomerSearchAjaxView, HallSearchAjaxView, SenderSearchAjaxView, ProductSearchAjaxView,
    DocumentFeePriceAjaxView, HallDeleteEvaluateAjaxView
)

app_name = 'masterdata'
urlpatterns = [
    path('customers/', CustomerListView.as_view(), name='customer-list'),
    path('customer/', CustomerDetailAjaxView.as_view(), name='customer-detail'),
    path('customer-update/', CustomerUpdateView.as_view(), name='customer-update'),
    path('customer-delete/', CustomerDeleteView.as_view(), name='customer-delete'),

    path('halls/', HallListView.as_view(), name='hall-list'),
    path('hall/', HallDetailAjaxView.as_view(), name='hall-detail'),
    path('hall-update/', HallUpdateView.as_view(), name='hall-update'),
    path('hall-delete/', HallDeleteView.as_view(), name='hall-delete'),
    path('hall-del-evaluate/', HallDeleteEvaluateAjaxView.as_view(), name='hall-del-evaluate'),

    # path('senders/', SenderListView.as_view(), name='sender-list'),
    path('sender/', SenderDetailAjaxView.as_view(), name='sender-detail'),
    path('sender-update/', SenderUpdateView.as_view(), name='sender-update'),
    path('sender-delete/', SenderDeleteView.as_view(), name='sender-delete'),

    path('products/', ProductListView.as_view(), name='product-list'),
    path('product/', ProductDetailAjaxView.as_view(), name='product-detail'),
    path('product-update/', ProductUpdateView.as_view(), name='product-update'),
    path('product-delete/', ProductDeleteView.as_view(), name='product-delete'),

    path('documents/', DocumentListView.as_view(), name='document-list'),
    path('document/', DocumentDetailAjaxView.as_view(), name='document-detail'),
    path('document-update/', DocumentUpdateView.as_view(), name='document-update'),
    path('document-delete/', DocumentDeleteView.as_view(), name='document-delete'),
    
    path('people-in-charge/', PersonInChargeListView.as_view(), name='people-list'),
    path('person-in-charge/', PersonInChargeDetailAjaxView.as_view(), name='person-detail'),
    path('person-in-charge/update/', PersonInChargeUpdateView.as_view(), name='person-update'),
    path('person-in-charge/delete/', PersonInChargeDeleteView.as_view(), name='person-delete'),


    path('search-customer/', CustomerSearchAjaxView.as_view(), name='customer-search'),
    path('search-hall-customer/', CustomerSearchAjaxView.as_view(), name='hall-customer-search'),
    path('search-hall/', HallSearchAjaxView.as_view(), name='hall-search'),
    path('search-sender/', SenderSearchAjaxView.as_view(), name='sender-search'),
    path('search-product/', ProductSearchAjaxView.as_view(), name='product-search'),

    path('document-fee/', DocumentFeePriceAjaxView.as_view(), name='document-fee-price'),

]