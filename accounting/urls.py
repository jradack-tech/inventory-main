from django.urls import path
from .views import SalesListView, PurchasesListView

app_name = 'accounting'
urlpatterns = [
    path('sales/', SalesListView.as_view(), name='sales-list'),
    path('purchases/', PurchasesListView.as_view(), name='purchases-list'),
]
