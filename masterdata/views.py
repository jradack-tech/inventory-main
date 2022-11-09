from django.shortcuts import render, redirect
from django.views.generic.base import View, TemplateView
from django.views.generic.list import ListView
from django.views.generic.edit import FormMixin, FormView, CreateView
from django.http import JsonResponse
from django.db.models import Q
from users.views import AdminLoginRequiredMixin
from .models import Customer, Hall, Sender, Product, Document, DocumentFee, PersonInCharge
from .forms import CustomerForm, HallForm, ProductForm, DocumentForm, PersonInChargeForm, SenderForm
from .filters import CustomerFilter, HallFilter, SenderFilter, ProductFilter
import json
from contracts.models import HallSalesContract, HallPurchasesContract, HallContract

class CustomerListView(AdminLoginRequiredMixin, ListView):
    template_name = 'masterdata/customers.html'
    queryset = Customer.objects.all()
    context_object_name = 'customers'
    paginate_by = 10

    def get_queryset(self):
        return CustomerFilter(self.request.GET, queryset=self.queryset).qs.order_by('pk')

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        context['customer_filter'] = CustomerFilter(self.request.GET)
        return context

    def post(self, request, *args, **kwargs):
        form = CustomerForm(request.POST)
        sender_form = SenderForm(request.POST)
        if form.is_valid():
            form.save()
        if sender_form.is_valid():
            sender_form.save()
        return redirect('masterdata:customer-list')


class CustomerUpdateView(AdminLoginRequiredMixin, View):
    def post(self, *args, **kwargs):
        data = self.request.POST
        id = data['id']
        customer = Customer.objects.get(id=id)
        sender = Sender.objects.filter(customer_name=customer.customer_name, address=customer.address).first()
        customer.name = data['name']
        customer.frigana = data['frigana']
        customer.postal_code = data['postal_code']
        customer.address = data['address']
        customer.tel = data['tel']
        customer.fax = data['fax']
        customer.excel = data['excel']
        customer.save()

        sender.name = data['name']
        sender.frigana = data['frigana']
        sender.postal_code = data['postal_code']
        sender.address = data['address']
        sender.tel = data['tel']
        sender.fax = data['fax']
        sender.excel = data['excel']
        sender.save()

        return redirect('masterdata:customer-list')


class CustomerDeleteView(AdminLoginRequiredMixin, View):
    def post(self, *args, **kwargs):
        data = self.request.POST
        id = data['id']
        
        customer = Customer.objects.get(id=id)
        # customer.delete()

        sender = Sender.objects.filter(customer_name=customer.customer_name, address=customer.address).first()
        sender.delete()
        customer.delete()

        return redirect('masterdata:customer-list')


class CustomerDetailAjaxView(AdminLoginRequiredMixin, View):
    def post(self, *args, **kwargs):
        if self.request.method == 'POST' and self.request.is_ajax():
            id = self.request.POST.get('id')
            customer = Customer.objects.get(id=id)
            return JsonResponse({
                'name': customer.name,
                'frigana': customer.frigana,
                'postal_code': customer.postal_code,
                'address': customer.address,
                'tel': customer.tel,
                'fax': customer.fax,
                'excel': customer.excel
            }, status=200)
        return JsonResponse({'success': False}, status=400)


class HallListView(AdminLoginRequiredMixin, ListView):
    template_name = 'masterdata/halls.html'
    queryset = Hall.objects.all()
    context_object_name = 'halls'
    paginate_by = 10

    def get_queryset(self):
        return HallFilter(self.request.GET, queryset=self.queryset).qs.order_by('pk')

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        context['hall_filter'] = HallFilter(self.request.GET)
        return context

    def post(self, request, *args, **kwargs):
        form = HallForm(request.POST)
        if form.is_valid():
            form.save()
        return redirect('masterdata:hall-list')


class HallUpdateView(AdminLoginRequiredMixin, View):
    def post(self, *args, **kwargs):
        data = self.request.POST
        id = data['id']
        hall = Hall.objects.get(id=id)
        hall.name = data['name']
        hall.customer_name = data['customer_name']
        hall.customer_frigana = data['customer_frigana']
        hall.frigana = data['frigana']
        hall.postal_code = data['postal_code']
        hall.address = data['address']
        hall.tel = data['tel']
        hall.fax = data['fax']
        hall.payee = data['payee']
        hall.save()
        return redirect('masterdata:hall-list')


class HallDeleteView(AdminLoginRequiredMixin, View):
    def post(self, *args, **kwargs):
        data = self.request.POST
        id = data['id']
        hall = Hall.objects.get(id=id)

        hall.delete()
        return redirect('masterdata:hall-list')

class HallDeleteEvaluateAjaxView(AdminLoginRequiredMixin, View):
    def post(self, *args, **kwargs):
        if self.request.method == 'POST' and self.request.is_ajax():
            id = self.request.POST.get('id')
            hall = Hall.objects.get(id=id)
            contract_hall = HallSalesContract.objects.filter(hall=hall).first() or HallPurchasesContract.objects.filter(hall=hall).first()

            if not contract_hall:
                return JsonResponse({'is_exist': False}, status=200)
            else:
                return JsonResponse({'is_exist': True}, status=200)
            
        return JsonResponse({'is_exist': None}, status=400)

class HallDetailAjaxView(AdminLoginRequiredMixin, View):
    def post(self, *args, **kwargs):
        if self.request.method == 'POST' and self.request.is_ajax():
            id = self.request.POST.get('id')
            hall = Hall.objects.get(id=id)
            return JsonResponse({
                'name': hall.name,
                'customer_name': hall.customer_name,
                'customer_frigana': hall.customer_frigana,
                'frigana': hall.frigana,
                'postal_code': hall.postal_code,
                'address': hall.address,
                'tel': hall.tel,
                'fax': hall.fax,
                'payee': hall.payee
            }, status=200)
        return JsonResponse({'success': False}, status=400)


# class SenderListView(AdminLoginRequiredMixin, ListView):
#     template_name = 'masterdata/senders.html'
#     queryset = Sender.objects.all()
#     context_object_name = 'senders'
#     paginate_by = 10

#     def get_queryset(self):
#         return SenderFilter(self.request.GET, queryset=self.queryset).qs.order_by('pk')

#     def get_context_data(self, **kwargs):
#         context = super().get_context_data(**kwargs)
#         context['sender_filter'] = SenderFilter(self.request.GET)
#         return context

#     def post(self, request, *args, **kwargs):
#         form = SenderForm(request.POST)
#         if form.is_valid():
#             form.save()
#         return redirect('masterdata:sender-list')


class SenderUpdateView(AdminLoginRequiredMixin, View):
    def post(self, *args, **kwargs):
        data = self.request.POST
        id = data['id']
        sender = Sender.objects.get(id=id)
        sender.name = data['name']
        sender.frigana = data['frigana']
        sender.postal_code = data['postal_code']
        sender.address = data['address']
        sender.tel = data['tel']
        sender.fax = data['fax']
        sender.save()
        return redirect('masterdata:sender-list')


class SenderDeleteView(AdminLoginRequiredMixin, View):
    def post(self, *args, **kwargs):
        data = self.request.POST
        id = data['id']
        sender = Sender.objects.get(id=id)
        # sender = Customer.objects.get(id=id)
        sender.delete()
        return redirect('masterdata:sender-list')


class SenderDetailAjaxView(AdminLoginRequiredMixin, View):
    def post(self, *args, **kwargs):
        if self.request.method == 'POST' and self.request.is_ajax():
            id = self.request.POST.get('id')
            sender = Sender.objects.get(id=id)
            # sender = Customer.objects.get(id=id)
            return JsonResponse({
                'name': sender.name,
                'frigana': sender.frigana,
                'postal_code': sender.postal_code,
                'address': sender.address,
                'tel': sender.tel,
                'fax': sender.fax,
            }, status=200)
        return JsonResponse({'success': False}, status=400)


class ProductListView(AdminLoginRequiredMixin, ListView):
    template_name = 'masterdata/products.html'
    queryset = Product.objects.all()
    context_object_name = 'products'
    paginate_by = 10

    def get_queryset(self):
        return ProductFilter(self.request.GET, queryset=self.queryset).qs.order_by('pk')

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        context['product_filter'] = ProductFilter(self.request.GET)
        context['people'] = PersonInCharge.objects.all()
        return context

    def post(self, request, *args, **kwargs):
        form = ProductForm(request.POST)
        if form.is_valid():
            form.save()
        return redirect('masterdata:product-list')


class ProductUpdateView(AdminLoginRequiredMixin, View):
    def post(self, *args, **kwargs):
        id = self.request.POST.get('id')
        product = Product.objects.get(id=id)
        product_form = ProductForm(self.request.POST)
        if product_form.is_valid():
            data = product_form.cleaned_data
            product.name = data.get('name')
            product.type = data.get('type')
            product.maker = data.get('maker')
            product.purchase_date = data.get('purchase_date')
            product.supplier = data.get('supplier')
            product.person_in_charge = data.get('person_in_charge')
            product.quantity = data.get('quantity')
            product.price = data.get('price')
            product.save()
        return redirect('masterdata:product-list')


class ProductDeleteView(AdminLoginRequiredMixin, View):
    def post(self, *args, **kwargs):
        data = self.request.POST
        id = data['id']
        product = Product.objects.get(id=id)
        product.delete()
        return redirect('masterdata:product-list')


class ProductDetailAjaxView(AdminLoginRequiredMixin, View):
    def post(self, *args, **kwargs):
        if self.request.method == 'POST' and self.request.is_ajax():
            id = self.request.POST.get('id')
            product = Product.objects.get(id=id)
            return JsonResponse({
                'name': product.name,
                'type': product.type,
                'maker': product.maker,
                'purchase_date': product.purchase_date,
                'supplier': product.supplier,
                'person_in_charge': product.person_in_charge,
                'quantity': product.quantity,
                'price': product.price,
                # 'stock': product.stock,
            }, status=200)
        return JsonResponse({'success': False}, status=400)


class DocumentListView(AdminLoginRequiredMixin, ListView):
    template_name = 'masterdata/documents.html'
    queryset = Document.objects.all()
    context_object_name = 'documents'

    def post(self, request, *args, **kwargs):
        form = DocumentForm(request.POST)
        if form.is_valid():
            form.save()
        return redirect('masterdata:document-list')


class DocumentUpdateView(AdminLoginRequiredMixin, View):
    def post(self, *args, **kwargs):
        data = self.request.POST
        id = data['id']
        document = Document.objects.get(id=id)
        document.name = data['name']
        document.term = data['term']
        document.taxation = data['taxation']
        document.save()
        return redirect('masterdata:document-list')


class DocumentDeleteView(AdminLoginRequiredMixin, View):
    def post(self, *args, **kwargs):
        data = self.request.POST
        id = data['id']
        document = Document.objects.get(id=id)
        document.delete()
        return redirect('masterdata:document-list')


class DocumentDetailAjaxView(AdminLoginRequiredMixin, View):
    def post(self, *args, **kwargs):
        if self.request.method == 'POST' and self.request.is_ajax():
            id = self.request.POST.get('id')
            document = Document.objects.get(id=id)
            return JsonResponse({
                'name': document.name,
                'term': document.term,
                'taxation': document.taxation,
            }, status=200)
        return JsonResponse({'success': False}, status=400)


class PersonInChargeListView(AdminLoginRequiredMixin, ListView):
    template_name = 'masterdata/people_in_charge.html'
    queryset = PersonInCharge.objects.all()
    context_object_name = 'people'

    def post(self, request, *args, **kwargs):
        form = PersonInChargeForm(request.POST)
        if form.is_valid():
            form.save()
        return redirect('masterdata:people-list')


class PersonInChargeUpdateView(AdminLoginRequiredMixin, View):
    def post(self, *args, **kwargs):
        data = self.request.POST
        id = data['id']
        person = PersonInCharge.objects.get(id=id)
        person.name = data['name']
        person.save()
        return redirect('masterdata:people-list')


class PersonInChargeDeleteView(AdminLoginRequiredMixin, View):
    def post(self, *args, **kwargs):
        data = self.request.POST
        id = data['id']
        person = PersonInCharge.objects.get(id=id)
        person.delete()
        return redirect('masterdata:people-list')


class PersonInChargeDetailAjaxView(AdminLoginRequiredMixin, View):
    def post(self, *args, **kwargs):
        if self.request.method == 'POST' and self.request.is_ajax():
            id = self.request.POST.get('id')
            person = PersonInCharge.objects.get(id=id)
            return JsonResponse({
                'name': person.name,
            }, status=200)
        return JsonResponse({'success': False}, status=400)


class CustomerSearchAjaxView(AdminLoginRequiredMixin, View):
    def get(self, *args, **kwargs):
        if self.request.method == 'GET' and self.request.is_ajax():
            search = self.request.GET.get('q')
            page = int(self.request.GET.get('page', 1))
            start = 30 * (page - 1)
            end = 30 * page
            customer_qs = Customer.objects.filter(
                Q(name__icontains=search) |
                Q(frigana__icontains=search) |
                Q(tel__icontains=search) |
                Q(fax__icontains=search)
            )
            total_count = customer_qs.count()
            customer_qs = customer_qs.order_by('id')[start:end].values(
                'id', 'name', 'frigana', 'tel', 'fax', 'postal_code', 'address')
            customers = list(customer_qs)
            return JsonResponse({"customers": customers, "total_count": total_count}, safe=False, status=200)
        return JsonResponse({'success': False}, status=400)


class HallSearchAjaxView(AdminLoginRequiredMixin, View):
    def get(self, *args, **kwargs):
        if self.request.method == 'GET' and self.request.is_ajax():
            search = self.request.GET.get('q')
            page = int(self.request.GET.get('page', 1))
            start = 30 * (page - 1)
            end = 30 * page
            hall_qs = Hall.objects.filter(
                Q(customer_name__icontains=search) |
                Q(name__icontains=search) |
                Q(frigana__icontains=search) |
                Q(tel__icontains=search) |
                Q(fax__icontains=search)
            )
            total_count = hall_qs.count()
            hall_qs = hall_qs.order_by('id')[start:end].values(
                'id', 'name', 'frigana', 'tel', 'fax', 'address', 'customer_name', 'customer_frigana')
            halls = list(hall_qs)
            return JsonResponse({"halls": halls, "total_count": total_count}, safe=False, status=200)
        return JsonResponse({'success': False}, status=400)


class SenderSearchAjaxView(AdminLoginRequiredMixin, View):
    def get(self, *args, **kwargs):
        if self.request.method == 'GET' and self.request.is_ajax():
            search = self.request.GET.get('q')
            page = int(self.request.GET.get('page', 1))
            start = 30 * (page - 1)
            end = 30 * page
            # sender_qs = Sender.objects.filter(name__icontains=search)
            sender_qs = Sender.objects.filter(
                Q(name__icontains=search) |
                Q(frigana__icontains=search) |
                Q(tel__icontains=search) |
                Q(fax__icontains=search)
            )
            total_count = sender_qs.count()
            sender_qs = sender_qs.order_by('id')[start:end].values(
                'id', 'name', 'address', 'tel', 'fax', 'postal_code')
            senders = list(sender_qs)
            return JsonResponse({"senders": senders, "total_count": total_count}, safe=False, status=200)
        return JsonResponse({'success': False}, status=400)


class ProductSearchAjaxView(AdminLoginRequiredMixin, View):
    def get(self, *args, **kwargs):
        if self.request.method == 'GET' and self.request.is_ajax():
            search = self.request.GET.get('q')
            page = int(self.request.GET.get('page', 1))
            start = 30 * (page - 1)
            end = 30 * page
            product_qs = Product.objects.filter(name__icontains=search)
            total_count = product_qs.count()
            product_qs = product_qs.order_by(
                'id')[start:end].values('id', 'name')
            products = list(product_qs)
            return JsonResponse({"products": products, "total_count": total_count}, safe=False, status=200)
        return JsonResponse({'success': False}, status=400)


class DocumentFeePriceAjaxView(AdminLoginRequiredMixin, View):
    def post(self, *args, **kwargs):
        if self.request.method == 'POST' and self.request.is_ajax():
            id = self.request.POST.get('id')
            document_fee = DocumentFee.objects.get(id=id)
            model_price = document_fee.model_price
            unit_price = document_fee.unit_price
            application_fee = document_fee.application_fee
            return JsonResponse({"model_price": model_price, "unit_price": unit_price, "application_fee": application_fee}, safe=False, status=200)
        return JsonResponse({'success': False}, status=400)
