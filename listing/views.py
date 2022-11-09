from django.db.models.query import QuerySet
import xlwt
import json
from datetime import datetime
from django.shortcuts import redirect
from django.views.generic.base import View
from django.views.generic.list import ListView
from django.http import HttpResponse, JsonResponse, QueryDict
from django.utils.translation import gettext as _
from django.contrib.contenttypes.models import ContentType
from django.db.models import Q

from django.core import serializers
from django.forms.models import model_to_dict

from django.core.paginator import Paginator
from users.views import AdminLoginRequiredMixin
from masterdata.models import Product, STOCK_CHOICES
from contracts.models import ContractProduct
from listing.models import ExportHistory
from contracts.utilities import generate_random_number, date_dump, log_export_operation
from .filters import ProductFilter
from .forms import ListingSearchForm, ProductForm, ListingLinkSearchForm, ListingLinkShowSearchForm
from contracts.models import TraderPurchasesContract, TraderSalesContract, TraderLink, TraderContract, Milestone

cell_width = 256 * 20
wide_cell_width = 256 * 45
cell_height = int(256 * 1.5)
header_height = int(cell_height * 1.5)
font_size = 20 * 10 # pt

# purchase_queryset, sales_queryset, _, phid, _, shid = getQuerySet()

common_style = xlwt.easyxf('font: height 200; align: vert center, horiz left, wrap on;\
                            borders: top_color black, bottom_color black, right_color black, left_color black,\
                            left thin, right thin, top thin, bottom thin;')
date_style = xlwt.easyxf('font: height 200; align: vert center, horiz left, wrap on;\
                            borders: top_color black, bottom_color black, right_color black, left_color black,\
                            left thin, right thin, top thin, bottom thin;')
bold_style = xlwt.easyxf('font: bold on, height 200; align: vert center, horiz left, wrap on;\
                            borders: top_color black, bottom_color black, right_color black, left_color black,\
                            left thin, right thin, top thin, bottom thin;')

class SalesListView(AdminLoginRequiredMixin, ListView):
    template_name = 'listing/sales.html'
    context_object_name = 'products'
    paginate_by = 5

    def get_queryset(self):
        trader_class_id = ContentType.objects.get(model='tradersalescontract').id
        hall_class_id = ContentType.objects.get(model='hallsalescontract').id
        queryset = ContractProduct.objects.filter(
            Q(content_type_id=trader_class_id) |
            Q(content_type_id=hall_class_id)
        ).filter(available='T').order_by('-pk')
        search_form = ListingSearchForm(self.request.GET)
        if search_form.is_valid():
            contract_id = search_form.cleaned_data.get('contract_id')
            created_at =  search_form.cleaned_data.get('created_at')
            created_at_to =  search_form.cleaned_data.get('created_at_to')
            shipping_date_from =  search_form.cleaned_data.get('shipping_date_from')
            shipping_date_to =  search_form.cleaned_data.get('shipping_date_to')
            customer = search_form.cleaned_data.get('customer')
            name = search_form.cleaned_data.get('name')
            person_in_charge = search_form.cleaned_data.get('person_in_charge')
            inventory_status = search_form.cleaned_data.get('inventory_status')
            if contract_id:
                queryset = queryset.filter(
                    Q(trader_sales_contract__contract_id__icontains=contract_id) |
                    Q(hall_sales_contract__contract_id__icontains=contract_id)
                )
            if created_at:
                queryset = queryset.filter(
                    Q(trader_sales_contract__created_at__gte=created_at) |
                    Q(hall_sales_contract__created_at__gte=created_at)
                )

            if created_at_to:
                queryset = queryset.filter(
                    Q(trader_sales_contract__created_at__lte=created_at_to) |
                    Q(hall_sales_contract__created_at__lte=created_at_to)
                )
            
            if shipping_date_from:
                queryset = queryset.filter(
                    Q(trader_sales_contract__shipping_date__gte=shipping_date_from) |
                    Q(hall_sales_contract__shipping_date__gte=shipping_date_from)
                )

            if shipping_date_to:
                queryset = queryset.filter(
                    Q(trader_sales_contract__shipping_date__lte=shipping_date_to) |
                    Q(hall_sales_contract__shipping_date__lte=shipping_date_to)
                )

            if customer:
                queryset = queryset.filter(
                    Q(trader_sales_contract__customer__name__icontains=customer) |
                    # Q(hall_sales_contract__customer__name__icontains=customer)
                    Q(hall_sales_contract__hall__customer_name__icontains=customer)
                )

            if person_in_charge:
                queryset = queryset.filter(
                    Q(trader_sales_contract__person_in_charge__icontains=person_in_charge) |
                    Q(hall_sales_contract__person_in_charge__icontains=person_in_charge)
                )

            if name:
                queryset = queryset.filter(Q(product__name__icontains=name))
            if inventory_status:
                queryset = queryset.filter(Q(status=inventory_status))
        
        for i in range(len(queryset)):
            if queryset[i].content_type_id == hall_class_id:
                milestone = Milestone.objects.filter(object_id=queryset[i].object_id, content_type_id=hall_class_id).first()
                queryset[i].content_object.payment_due_date = milestone.date if milestone else None

        return queryset
    
    def post(self, request, *args, **kwargs):
        user_id = self.request.user.id
        log_export_operation(user_id, "{} - {}".format(_("List"), _("Sales")))
        
        response = HttpResponse(content_type='application/ms-excel')
        response['Content-Disposition'] = 'attachment; filename="listing_sales_{}.xls"'.format(generate_random_number())
        wb = xlwt.Workbook(encoding='utf-8')
        ws = wb.add_sheet("{} - {}".format(_('List'), _('Sales')))

        ws.set_header_str(str.encode(''))
        ws.set_footer_str(str.encode(''))
        ws.set_left_margin(0.314961)
        ws.set_top_margin(0.393701)
        ws.set_right_margin(0.314961)
        ws.set_bottom_margin(0.19685)

        ws.row(0).height_mismatch = True
        ws.row(0).height = header_height

        ws.write(0, 0, _('Contract ID'), bold_style)
        ws.write(0, 1, _('Contract date'), bold_style)
        ws.write(0, 2, _('Customer'), bold_style)
        ws.write(0, 3, _('Delivered place'), bold_style)
        ws.write(0, 4, _('Person in charge'), bold_style)
        ws.write(0, 5, _('Payment date'), bold_style)
        ws.write(0, 6, _('Product name'), bold_style)
        ws.write(0, 7, _('Unit count'), bold_style)
        ws.write(0, 8, _('Amount'), bold_style)
        ws.write(0, 9, _('Inventory status'), bold_style)

        for i in range(10):
            ws.col(i).width = cell_width
        ws.col(2).width = ws.col(3).width = ws.col(6).width = wide_cell_width

        row_no = 1
        date_style.num_format_str = 'yyyy/mm/dd' if self.request.LANGUAGE_CODE == 'ja' else 'mm/dd/yyyy'
        queryset = self.get_queryset()
        for product in queryset:
            contract = product.content_object
            contract_id = contract.contract_id
            contract_date = contract.created_at
            customer = contract.customer.name if contract.customer else None
            person_in_charge = contract.person_in_charge
            if product.content_type_id == ContentType.objects.get(model='hallsalescontract').id:
                destination = contract.hall.name if contract.hall else None
                payment_date = contract.shipping_date
            else:
                destination = None
                payment_date = contract.payment_due_date
            product_name = product.product.name
            quantity = product.quantity
            amount = product.amount
            status = product.status
            
            ws.write(row_no, 0, contract_id, common_style)
            ws.write(row_no, 1, contract_date, date_style)
            ws.write(row_no, 2, customer, common_style)
            ws.write(row_no, 3, destination, common_style)
            ws.write(row_no, 4, person_in_charge, common_style)
            ws.write(row_no, 5, payment_date, date_style)
            ws.write(row_no, 6, product_name, common_style)
            ws.write(row_no, 7, quantity, common_style)
            ws.write(row_no, 8, amount, common_style)
            ws.write(row_no, 9, str(dict(STOCK_CHOICES)[status]), common_style)
            row_no += 1
        
        wb.save(response)
        return response
    
    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        context['hall_contract_id'] = ContentType.objects.get(model='hallsalescontract').id
        params = self.request.GET.copy()
        for k, v in params.items():
            if v:
                context[k] = v
        return context


class PurchasesListView(AdminLoginRequiredMixin, ListView):
    template_name = 'listing/purchases.html'
    context_object_name = 'products'
    paginate_by = 5

    def get_queryset(self):

        trader_class_id = ContentType.objects.get(model='traderpurchasescontract').id
        hall_class_id = ContentType.objects.get(model='hallpurchasescontract').id
        queryset = ContractProduct.objects.filter(
            Q(content_type_id=trader_class_id) |
            Q(content_type_id=hall_class_id)
        ).filter(available='T').order_by('-pk')

        # purchase_trader_class_id = ContentType.objects.get(model='traderpurchasescontract').id
        # purchase_hall_class_id = ContentType.objects.get(model='hallpurchasescontract').id
        # purchase_queryset = ContractProduct.objects.filter(
        #     Q(content_type_id=purchase_trader_class_id) |
        #     Q(content_type_id=purchase_hall_class_id)
        # ).order_by('-pk')

        search_form = ListingSearchForm(self.request.GET)
        if search_form.is_valid():
            contract_id = search_form.cleaned_data.get('contract_id')
            created_at =  search_form.cleaned_data.get('created_at')
            created_at_to =  search_form.cleaned_data.get('created_at_to')
            shipping_date_from =  search_form.cleaned_data.get('shipping_date_from')
            shipping_date_to =  search_form.cleaned_data.get('shipping_date_to')
            customer = search_form.cleaned_data.get('customer')
            name = search_form.cleaned_data.get('name')
            person_in_charge = search_form.cleaned_data.get('person_in_charge')
            inventory_status = search_form.cleaned_data.get('inventory_status')
            if contract_id:
                queryset = queryset.filter(
                    Q(trader_purchases_contract__contract_id__icontains=contract_id) |
                    Q(hall_purchases_contract__contract_id__icontains=contract_id)
                )
            if created_at:
                queryset = queryset.filter(
                    Q(trader_purchases_contract__created_at__gte=created_at) |
                    Q(hall_purchases_contract__created_at__gte=created_at)
                )

            if created_at_to:
                queryset = queryset.filter(
                    Q(trader_purchases_contract__created_at__lte=created_at_to) |
                    Q(hall_purchases_contract__created_at__lte=created_at_to)
                )

            if shipping_date_from:
                queryset = queryset.filter(
                    Q(trader_purchases_contract__shipping_date__gte=shipping_date_from) |
                    Q(hall_purchases_contract__shipping_date__gte=shipping_date_from)
                )

            if shipping_date_to:
                queryset = queryset.filter(
                    Q(trader_purchases_contract__shipping_date__lte=shipping_date_to) |
                    Q(hall_purchases_contract__shipping_date__lte=shipping_date_to)
                )

            if customer:
                queryset = queryset.filter(
                    Q(trader_purchases_contract__customer__name__icontains=customer) |
                    Q(hall_purchases_contract__hall__customer_name__icontains=customer)
                )

            if person_in_charge:
                queryset = queryset.filter(
                    Q(trader_purchases_contract__person_in_charge__icontains=person_in_charge) |
                    Q(hall_purchases_contract__person_in_charge__icontains=person_in_charge)
                )
            if name:
                queryset = queryset.filter(Q(product__name__icontains=name))
            if inventory_status:
                queryset = queryset.filter(Q(status=inventory_status))

        for i in range(len(queryset)):
            if queryset[i].content_type_id == hall_class_id:
                milestone = Milestone.objects.filter(object_id=queryset[i].object_id, content_type_id=hall_class_id).first()
                queryset[i].content_object.payment_due_date = milestone.date if milestone else None

        return queryset
    
    def post(self, request, *args, **kwargs):
        user_id = self.request.user.id
        log_export_operation(user_id, "{} - {}".format(_("List"), _("Purchases")))

        
        response = HttpResponse(content_type='application/ms-excel')
        response['Content-Disposition'] = 'attachment; filename="listing_purchases_{}.xls"'.format(generate_random_number())
        wb = xlwt.Workbook(encoding='utf-8')
        ws = wb.add_sheet("{} - {}".format(_('List'), _('Purchases')))

        ws.set_header_str(str.encode(''))
        ws.set_footer_str(str.encode(''))
        ws.set_left_margin(0.314961)
        ws.set_top_margin(0.393701)
        ws.set_right_margin(0.314961)
        ws.set_bottom_margin(0.19685)

        ws.row(0).height_mismatch = True
        ws.row(0).height = header_height

        ws.write(0, 0, _('Contract ID'), bold_style)
        ws.write(0, 1, _('Contract date'), bold_style)
        ws.write(0, 2, _('Customer'), bold_style)
        ws.write(0, 3, _('Delivered place'), bold_style)
        ws.write(0, 4, _('Person in charge'), bold_style)
        ws.write(0, 5, _('Payment date'), bold_style)
        ws.write(0, 6, _('Product name'), bold_style)
        ws.write(0, 7, _('Unit count'), bold_style)
        ws.write(0, 8, _('Amount'), bold_style)
        ws.write(0, 9, _('Inventory status'), bold_style)

        for i in range(10):
            ws.col(i).width = cell_width
        ws.col(2).width = ws.col(3).width = ws.col(6).width = wide_cell_width

        row_no = 1
        date_style.num_format_str = 'yyyy/mm/dd' if self.request.LANGUAGE_CODE == 'ja' else 'mm/dd/yyyy'
        queryset = self.get_queryset()
        for product in queryset:
            contract = product.content_object
            contract_id = contract.contract_id
            contract_date = contract.created_at
            customer = contract.customer.name if contract.customer else None
            person_in_charge = contract.person_in_charge
            if product.content_type_id == ContentType.objects.get(model='hallpurchasescontract').id:
                destination = contract.hall.name if contract.hall else None
                payment_date = contract.shipping_date
            else:
                destination = None
                payment_date = contract.transfer_deadline
            product_name = product.product.name
            quantity = product.quantity
            amount = product.amount
            status = product.status

            ws.write(row_no, 0, contract_id, common_style)
            ws.write(row_no, 1, contract_date, date_style)
            ws.write(row_no, 2, customer, common_style)
            ws.write(row_no, 3, destination, common_style)
            ws.write(row_no, 4, person_in_charge, common_style)
            ws.write(row_no, 5, payment_date, date_style)
            ws.write(row_no, 6, product_name, common_style)
            ws.write(row_no, 7, quantity, common_style)
            ws.write(row_no, 8, amount, common_style)
            ws.write(row_no, 9, str(dict(STOCK_CHOICES)[status]), common_style)
            row_no += 1
        
        wb.save(response)
        return response
    
    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        context['hall_contract_id'] = ContentType.objects.get(model='hallpurchasescontract').id
        params = self.request.GET.copy()
        for k, v in params.items():
            if v:
                context[k] = v
        return context

class DeletedListView(AdminLoginRequiredMixin, ListView):
    template_name = 'listing/deleted.html'
    context_object_name = 'products'
    paginate_by = 5

    def get_queryset(self):
        trader_class_id1 = ContentType.objects.get(model='tradersalescontract').id
        trader_class_id2 = ContentType.objects.get(model='traderpurchasescontract').id
        hall_class_id1 = ContentType.objects.get(model='hallsalescontract').id
        hall_class_id2 = ContentType.objects.get(model='hallpurchasescontract').id
        queryset = ContractProduct.objects.filter(
            Q(content_type_id=trader_class_id1) |
            Q(content_type_id=trader_class_id2) |
            Q(content_type_id=hall_class_id1) |
            Q(content_type_id=hall_class_id2)
        ).filter(available='F').order_by('-pk')

        whole = queryset.all()

        search_form = ListingSearchForm(self.request.GET)
        if search_form.is_valid():
            contract_id = search_form.cleaned_data.get('contract_id')
            created_at =  search_form.cleaned_data.get('created_at')
            created_at_to =  search_form.cleaned_data.get('created_at_to')
            shipping_date_from =  search_form.cleaned_data.get('shipping_date_from')
            shipping_date_to =  search_form.cleaned_data.get('shipping_date_to')
            customer = search_form.cleaned_data.get('customer')
            name = search_form.cleaned_data.get('name')
            person_in_charge = search_form.cleaned_data.get('person_in_charge')
            inventory_status = search_form.cleaned_data.get('inventory_status')
            if contract_id:
                queryset = queryset.filter(
                    Q(trader_sales_contract__contract_id__icontains=contract_id) |
                    Q(hall_sales_contract__contract_id__icontains=contract_id) |
                    Q(trader_purchases_contract__contract_id__icontains=contract_id) |
                    Q(hall_purchases_contract__contract_id__icontains=contract_id)
                )
            if created_at:
                queryset = queryset.filter(
                    Q(trader_sales_contract__created_at__gte=created_at) |
                    Q(hall_sales_contract__created_at__gte=created_at) |
                    Q(trader_purchases_contract__created_at__gte=created_at) |
                    Q(hall_purchases_contract__created_at__gte=created_at)
                )

            if created_at_to:
                queryset = queryset.filter(
                    Q(trader_sales_contract__created_at__lte=created_at_to) |
                    Q(hall_sales_contract__created_at__lte=created_at_to) |
                    Q(trader_purchases_contract__created_at__lte=created_at_to) |
                    Q(hall_purchases_contract__created_at__lte=created_at_to)
                )

            if shipping_date_from:
                queryset = queryset.filter(
                    Q(trader_sales_contract__created_at__gte=shipping_date_from) |
                    Q(hall_sales_contract__created_at__gte=shipping_date_from) |
                    Q(trader_purchases_contract__created_at__gte=shipping_date_from) |
                    Q(hall_purchases_contract__created_at__gte=shipping_date_from)
                )

            if shipping_date_to:
                queryset = queryset.filter(
                    Q(trader_sales_contract__created_at__lte=shipping_date_to) |
                    Q(hall_sales_contract__created_at__lte=shipping_date_to) |
                    Q(trader_purchases_contract__created_at__lte=shipping_date_to) |
                    Q(hall_purchases_contract__created_at__lte=shipping_date_to)
                )

            if customer:
                queryset = queryset.filter(
                    Q(trader_sales_contract__customer__name__icontains=customer) |
                    Q(hall_sales_contract__customer__name__icontains=customer) |
                    Q(trader_purchases_contract__customer__name__icontains=customer) |
                    Q(hall_purchases_contract__customer__name__icontains=customer)
                )

            if person_in_charge:
                queryset = queryset.filter(
                    Q(trader_sales_contract__person_in_charge__icontains=person_in_charge) |
                    Q(hall_sales_contract__person_in_charge__icontains=person_in_charge) |
                    Q(trader_purchases_contract__person_in_charge__icontains=person_in_charge) |
                    Q(hall_purchases_contract__person_in_charge__icontains=person_in_charge)
                )

            if name:
                queryset = queryset.filter(Q(product__name__icontains=name))
            if inventory_status:
                queryset = queryset.filter(Q(status=inventory_status))

        return queryset.order_by('-pk')
  
    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        context['hall_contract_id'] = ContentType.objects.get(model='hallsalescontract').id
        params = self.request.GET.copy()
        for k, v in params.items():
            if v:
                context[k] = v
        return context


class LinkProductsAjaxView(AdminLoginRequiredMixin, View):
    def get(self):
        queryset = TraderLink.objects.order_by('-pk')
        return queryset

    def delete(self, *args, **kwargs):
        p_id = QueryDict(self.request.body).get('p_id')
        s_id = QueryDict(self.request.body).get('s_id')
        p_same_order = QueryDict(self.request.body).get('p_same_order')
        s_same_order = QueryDict(self.request.body).get('s_same_order')

        TraderLink.objects.filter(purchase_contract_id=p_id, sale_contract_id=s_id, purchase_same_order=p_same_order, sale_same_order=s_same_order).first().delete()

        return JsonResponse({'success': True}, status=200)

    def post(self, *args, **kwargs):
        if self.request.method == 'POST' and self.request.is_ajax():

            p_contract_id = self.request.POST.get('p_contract_id')
            p_same_order = self.request.POST.get('p_same_order')
            s_contract_id = self.request.POST.get('s_contract_id')
            s_same_order = self.request.POST.get('s_same_order')

            purchase_queryset, sales_queryset, _, _, _, _ = getQuerySet()

            purchase_contract = purchase_queryset.filter(
                Q(trader_purchases_contract__contract_id__icontains=p_contract_id) |
                Q(hall_purchases_contract__contract_id__icontains=p_contract_id)
            ).all()

            sum = 0
            for p in purchase_contract:
                sum += p.quantity
                if sum > int(p_same_order):
                    purchase_contract = p
                    break

            sale_contract = sales_queryset.filter(
                Q(trader_sales_contract__contract_id__icontains=s_contract_id) |
                Q(hall_sales_contract__contract_id__icontains=s_contract_id)
            ).all()

            sum = 0
            for s in sale_contract:
                sum += s.quantity
                if sum > int(s_same_order):
                    sale_contract = s
                    break
            
            link_item = TraderLink.objects.filter(purchase_contract=purchase_contract, sale_contract=sale_contract, purchase_same_order=p_same_order, sale_same_order=s_same_order)
            if link_item:
                return JsonResponse({'success': "There exist connection"}, status=200)
            
            # validate the contracts
            if purchase_contract.quantity > 0 and sale_contract.quantity > 0:

                data = {
                    'purchase_contract': purchase_contract,
                    'sale_contract': sale_contract,
                    'purchase_same_order': p_same_order,
                    'sale_same_order': s_same_order,
                }
                TraderLink.objects.create(**data)
                return JsonResponse({'success': True}, status=200)

        return JsonResponse({'success': False}, status=400)

class LinkListShowView(AdminLoginRequiredMixin, ListView):
    template_name = 'listing/linkshow.html'
    context_object_name = 'links'
    paginate_by = 10

    def get_queryset(self):
        purchase_queryset, sales_queryset, _, phid, _, shid = getQuerySet()

        # Get contract_id of both in link
        link_queryset = TraderLink.objects.order_by('-pk').all()

        lists = [(link.purchase_contract.id, link.purchase_same_order, link.sale_contract.id, link.sale_same_order) for link in link_queryset]
        sale_list = [link.sale_contract.id for link in link_queryset]
        # report_date_list = [link.report_date for link in link_queryset]


        payload = []
        for p_id, p_same_order, s_id, s_same_order in lists:
            
            p_item = purchase_queryset.filter(Q(id=p_id)).first()
            s_item = sales_queryset.filter(Q(id=s_id)).first()
            p_report_date = json.loads(p_item.report_date) if p_item.report_date else {}
            s_report_date = json.loads(s_item.report_date) if s_item.report_date else {}

            item1 = {
                "id": p_item.id,
                "same_order": p_same_order,
                "object_id": p_item.content_object.id,
                "class_id": p_item.content_type_id,
                "contract_id": p_item.content_object.contract_id,
                "created_at": p_item.content_object.created_at,
                "customer_name": p_item.content_object.customer.name if p_item.content_object.customer else '',
                "hall_name": p_item.content_object.hall.name if p_item.content_type_id == phid else '',
                "person_in_charge": p_item.content_object.person_in_charge,
                "shipping_date": p_item.content_object.shipping_date if p_item.content_type_id == phid else p_item.content_object.transfer_deadline if hasattr(p_item.content_object, 'transfer_deadline') else '',
                "product_name": p_item.product.name,
                "quantity": p_item.quantity,
                "amount": p_item.amount,
                "status_display": p_item.get_status_display(),
                "report_date": datetime.strptime(p_report_date[str(p_same_order)], '%Y/%m/%d').date() if str(p_same_order) in p_report_date else '',
            }

            item2 = {
                "id": s_item.id,
                "same_order": s_same_order,
                "object_id": s_item.content_object.id,
                "class_id": s_item.content_type_id,
                "contract_id": s_item.content_object.contract_id,
                "created_at": s_item.content_object.created_at,
                "customer_name": s_item.content_object.customer.name if s_item.content_object.customer else '',
                "hall_name": s_item.content_object.hall.name if s_item.content_type_id == shid else '',
                "person_in_charge": s_item.content_object.person_in_charge,
                "shipping_date": s_item.content_object.shipping_date if s_item.content_type_id == shid else s_item.content_object.payment_due_date if hasattr(s_item.content_object, 'payment_due_date') else '',
                "product_name": s_item.product.name,
                "quantity": s_item.quantity,
                "amount": s_item.amount,
                "status_display": s_item.get_status_display(),
                "report_date": datetime.strptime(s_report_date[str(s_same_order)], '%Y/%m/%d').date() if str(s_same_order) in s_report_date else '',
            }

            payload.append({"p": item1, "s": item2})

        # Search Items 
        search_form = ListingLinkShowSearchForm(self.request.GET)
        if search_form.is_valid():
            contract_id = search_form.cleaned_data.get('contract_id')
            created_at = search_form.cleaned_data.get('created_at')
            created_at_to = search_form.cleaned_data.get('created_at_to')
            shipping_date_from = search_form.cleaned_data.get('shipping_date_from')
            shipping_date_to = search_form.cleaned_data.get('shipping_date_to')
            person_in_charge = search_form.cleaned_data.get('person_in_charge')
            report_date = search_form.cleaned_data.get('report_date')
            inventory_status = search_form.cleaned_data.get('inventory_status')

            if contract_id:
                payload = [item for item in payload if item['p']['contract_id'] == contract_id or item['s']['contract_id'] == contract_id]
            if created_at:
                payload = [item for item in payload if item['p']['created_at'] >= created_at or item['s']['created_at'] >= created_at]
            if created_at_to:
                payload = [item for item in payload if item['p']['created_at'] <= created_at_to or item['s']['created_at'] <= created_at_to]
            if shipping_date_from:
                payload = [item for item in payload if item['p']['shipping_date'] >= shipping_date_from or item['s']['shipping_date'] >= shipping_date_from]
            if shipping_date_to:
                payload = [item for item in payload if item['p']['shipping_date'] <= shipping_date_to or item['s']['shipping_date'] <= shipping_date_to]
            if person_in_charge:
                payload = [item for item in payload if item['p']['person_in_charge'] == person_in_charge or item['s']['person_in_charge'] == person_in_charge]
            if report_date:
                payload = [item for item in payload if item['p']['report_date'] == report_date or item['s']['report_date'] == report_date]
            if inventory_status and inventory_status != 'All':
                payload = [item for item in payload if item['p']['status_display'] == inventory_status or item['s']['status_display'] == inventory_status]
        
        return payload

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)

        params = self.request.GET.copy()
        for k, v in params.items():
            if v:
                context[k] = v

        return context

    def post(self, request, *args, **kwargs):
        user_id = self.request.user.id
        log_export_operation(user_id, "{} - {}".format(_("List"), _("Links")))

        response = HttpResponse(content_type='application/ms-excel')
        response['Content-Disposition'] = 'attachment; filename="listing_Links_show_{}.xls"'.format(generate_random_number())
        
        wb = xlwt.Workbook(encoding='utf-8')
        ws = wb.add_sheet("{} - {}".format(_('List'), _('Links')))

        ws.set_header_str(str.encode(''))
        ws.set_footer_str(str.encode(''))
        ws.set_left_margin(0.314961)
        ws.set_top_margin(0.393701)
        ws.set_right_margin(0.314961)
        ws.set_bottom_margin(0.19685)

        ws.row(0).height_mismatch = True
        ws.row(0).height = header_height

        ws.write(0, 0, _('Contract ID'), bold_style)
        ws.write(0, 1, _('Contract date'), bold_style)
        ws.write(0, 2, _('Customer'), bold_style)
        ws.write(0, 3, _('Delivered place'), bold_style)
        ws.write(0, 4, _('Person in charge'), bold_style)
        ws.write(0, 5, _('Payment date'), bold_style)
        ws.write(0, 6, _('Product name'), bold_style)
        ws.write(0, 7, _('Unit count'), bold_style)
        ws.write(0, 8, _('Amount'), bold_style)
        ws.write(0, 9, _('Inventory status'), bold_style)
        ws.write(0, 10, _('Report Date'), bold_style)

        for i in range(11):
            ws.col(i).width = cell_width
        ws.col(2).width = ws.col(3).width = ws.col(6).width = wide_cell_width

        row_no = 1
        date_style.num_format_str = 'yyyy/mm/dd' if self.request.LANGUAGE_CODE == 'ja' else 'mm/dd/yyyy'
        sets = self.get_queryset()

        for item in sets:
            if not item: continue

            ws.write(row_no, 0, item['p']['contract_id'], common_style)
            ws.write(row_no, 1, item['p']['created_at'], date_style)
            ws.write(row_no, 2, item['p']['customer_name'], common_style)
            ws.write(row_no, 3, item['p']['hall_name'], common_style)
            ws.write(row_no, 4, item['p']['person_in_charge'], date_style)
            ws.write(row_no, 5, item['p']['shipping_date'], common_style)
            ws.write(row_no, 6, item['p']['product_name'], common_style)
            ws.write(row_no, 7, item['p']['quantity'], common_style)
            ws.write(row_no, 8, item['p']['amount'], common_style)
            ws.write(row_no, 9, item['p']['status_display'], common_style)
            ws.write(row_no, 10, item['p']['report_date'], date_style)
            row_no += 1

            ws.write(row_no, 0, item['s']['contract_id'], common_style)
            ws.write(row_no, 1, item['s']['created_at'], date_style)
            ws.write(row_no, 2, item['s']['customer_name'], common_style)
            ws.write(row_no, 3, item['s']['hall_name'], common_style)
            ws.write(row_no, 4, item['s']['person_in_charge'], date_style)
            ws.write(row_no, 5, item['s']['shipping_date'], common_style)
            ws.write(row_no, 6, item['s']['product_name'], common_style)
            ws.write(row_no, 7, item['s']['quantity'], common_style)
            ws.write(row_no, 8, item['s']['amount'], common_style)
            ws.write(row_no, 9, item['s']['status_display'], common_style)
            ws.write(row_no, 10, item['s']['report_date'], date_style)
            row_no += 1
        
        wb.save(response)
        return response

def getQuerySet():
    purchase_trader_class_id = ContentType.objects.get(model='traderpurchasescontract').id
    purchase_hall_class_id = ContentType.objects.get(model='hallpurchasescontract').id
    purchase_queryset = ContractProduct.objects.filter(
        Q(content_type_id=purchase_trader_class_id) |
        Q(content_type_id=purchase_hall_class_id)
    ).order_by('-pk')

    sales_trader_class_id = ContentType.objects.get(model='tradersalescontract').id
    sales_hall_class_id = ContentType.objects.get(model='hallsalescontract').id
    sales_queryset = ContractProduct.objects.filter(
        Q(content_type_id=sales_trader_class_id) |
        Q(content_type_id=sales_hall_class_id)
    ).order_by('-pk')

    return purchase_queryset, sales_queryset, purchase_trader_class_id, purchase_hall_class_id, sales_trader_class_id, sales_hall_class_id
    
class LinkProductReportDateUpdate(AdminLoginRequiredMixin, View):

    def post(self, *args, **kwargs):
        if self.request.method == 'POST' and self.request.is_ajax():
            order_id = int(self.request.POST.get('order_id'))

            contract_id = self.request.POST.get('contract_id')
            report_date = datetime.strptime(self.request.POST.get('report_date'), '%m/%d/%Y')
            is_purchase = self.request.POST.get('is_purchase')

            purchase_queryset, sales_queryset, _, _, _, _ = getQuerySet()

            contract = None

            if is_purchase == 'true':
                contract = purchase_queryset.filter(
                    Q(trader_purchases_contract__contract_id__icontains=contract_id) |
                    Q(hall_purchases_contract__contract_id__icontains=contract_id)
                ).first()

            else:
                contract = sales_queryset.filter(
                    Q(trader_sales_contract__contract_id__icontains=contract_id) |
                    Q(hall_sales_contract__contract_id__icontains=contract_id)
                ).first()

            current_report_date = json.loads(contract.report_date) if contract.report_date else {}
            current_report_date[str(order_id)] = datetime.strftime(report_date, '%Y/%m/%d')
            contract.report_date = json.dumps(current_report_date)

            contract.save()
            return JsonResponse({'success': True}, status=200)
        return JsonResponse({'success': False}, status=400)


class LinkListView(AdminLoginRequiredMixin, ListView):
    template_name = 'listing/link.html'
    context_object_name = 'products'

    def get_queryset(self):
        
        ITEMS_PER_PAGE = 10

        purchase_queryset, sales_queryset, _, phid, _, shid = getQuerySet()

        search_form = ListingLinkSearchForm(self.request.GET)
        
        if search_form.is_valid():

            
            p_contract_id = search_form.cleaned_data.get('p_contract_id')
            p_created_at =  search_form.cleaned_data.get('p_created_at')
            p_created_at_to =  search_form.cleaned_data.get('p_created_at_to')
            p_shipping_date_from =  search_form.cleaned_data.get('p_shipping_date_from')
            p_shipping_date_to =  search_form.cleaned_data.get('p_shipping_date_to')
            p_customer = search_form.cleaned_data.get('p_customer')
            p_name = search_form.cleaned_data.get('p_name')
            p_person_in_charge = search_form.cleaned_data.get('p_person_in_charge')
            p_inventory_status = search_form.cleaned_data.get('p_inventory_status') or 'P'

            s_contract_id = search_form.cleaned_data.get('s_contract_id')
            s_created_at =  search_form.cleaned_data.get('s_created_at')
            s_created_at_to =  search_form.cleaned_data.get('s_created_at_to')
            s_shipping_date_from =  search_form.cleaned_data.get('s_shipping_date_from')
            s_shipping_date_to =  search_form.cleaned_data.get('s_shipping_date_to')
            s_customer = search_form.cleaned_data.get('s_customer')
            s_name = search_form.cleaned_data.get('s_name')
            s_person_in_charge = search_form.cleaned_data.get('s_person_in_charge')
            s_inventory_status = search_form.cleaned_data.get('s_inventory_status') or 'P'

            w_name = search_form.cleaned_data.get('w_name')
            w_inventory_status = search_form.cleaned_data.get('w_inventory_status') or 'P'
            w_person_in_charge = search_form.cleaned_data.get('w_person_in_charge')

            # Purchase contract data
            if p_contract_id:
                purchase_queryset = purchase_queryset.filter(
                    Q(trader_purchases_contract__contract_id__icontains=p_contract_id) |
                    Q(hall_purchases_contract__contract_id__icontains=p_contract_id)
                )

            if p_created_at:
                purchase_queryset = purchase_queryset.filter(
                    Q(trader_purchases_contract__created_at__gte=p_created_at) |
                    Q(hall_purchases_contract__created_at__gte=p_created_at)
                )

            if p_created_at_to:
                purchase_queryset = purchase_queryset.filter(
                    Q(trader_purchases_contract__created_at__lte=p_created_at_to) |
                    Q(hall_purchases_contract__created_at__lte=p_created_at_to)
                )

            if p_shipping_date_from:
                queryset = queryset.filter(
                    Q(trader_purchases_contract__shipping_date__gte=p_shipping_date_from) |
                    Q(hall_purchases_contract__shipping_date__gte=p_shipping_date_from)
                )

            if p_shipping_date_to:
                queryset = queryset.filter(
                    Q(trader_purchases_contract__shipping_date__lte=p_shipping_date_to) |
                    Q(hall_purchases_contract__shipping_date__lte=p_shipping_date_to)
                )

            if p_customer:
                purchase_queryset = purchase_queryset.filter(
                    Q(trader_purchases_contract__customer__name__icontains=p_customer) |
                    # Q(hall_purchases_contract__customer__name__icontains=p_customer)
                    Q(hall_purchases_contract__hall__customer_name__icontains=p_customer)
                )

            if w_person_in_charge:
                purchase_queryset = purchase_queryset.filter(
                    Q(trader_purchases_contract__person_in_charge__icontains=w_person_in_charge) |
                    Q(hall_purchases_contract__person_in_charge__icontains=w_person_in_charge)
                )

            if p_person_in_charge:
                purchase_queryset = purchase_queryset.filter(
                    Q(trader_purchases_contract__person_in_charge__icontains=p_person_in_charge) |
                    Q(hall_purchases_contract__person_in_charge__icontains=p_person_in_charge)
                )

            if w_name:
                purchase_queryset = purchase_queryset.filter(Q(product__name__icontains=w_name))

            if w_inventory_status != 'W':
                purchase_queryset = purchase_queryset.filter(Q(status=w_inventory_status))

            if p_name:
                purchase_queryset = purchase_queryset.filter(Q(product__name__icontains=p_name))

            if p_inventory_status != 'W':
                purchase_queryset = purchase_queryset.filter(Q(status=p_inventory_status))

            # Sales contract data
            if s_contract_id:
                sales_queryset = sales_queryset.filter(
                    Q(trader_sales_contract__contract_id__icontains=s_contract_id) |
                    Q(hall_sales_contract__contract_id__icontains=s_contract_id)
                )

            if s_created_at:
                sales_queryset = sales_queryset.filter(
                    Q(trader_sales_contract__created_at__gte=s_created_at) |
                    Q(hall_sales_contract__created_at__gte=s_created_at)
                )

            if s_created_at_to:
                sales_queryset = sales_queryset.filter(
                    Q(trader_sales_contract__created_at__lte=s_created_at_to) |
                    Q(hall_sales_contract__created_at__lte=s_created_at_to)
                )

            if s_shipping_date_from:
                queryset = queryset.filter(
                    Q(trader_sales_contract__shipping_date__gte=s_shipping_date_from) |
                    Q(hall_sales_contract__shipping_date__gte=s_shipping_date_from)
                )

            if s_shipping_date_to:
                queryset = queryset.filter(
                    Q(trader_sales_contract__shipping_date__lte=s_shipping_date_to) |
                    Q(hall_sales_contract__shipping_date__lte=s_shipping_date_to)
                )

            if s_customer:
                sales_queryset = sales_queryset.filter(
                    Q(trader_sales_contract__customer__name__icontains=s_customer) |
                    # Q(hall_sales_contract__customer__name__icontains=s_customer)
                    Q(hall_sales_contract__hall__customer_name__icontains=s_customer)
                )

            if w_person_in_charge:
                sales_queryset = sales_queryset.filter(
                    Q(trader_sales_contract__person_in_charge__icontains=w_person_in_charge) |
                    Q(hall_sales_contract__person_in_charge__icontains=w_person_in_charge)
                )

            if s_person_in_charge:
                sales_queryset = sales_queryset.filter(
                    Q(trader_sales_contract__person_in_charge__icontains=s_person_in_charge) |
                    Q(hall_sales_contract__person_in_charge__icontains=s_person_in_charge)
                )
            
            if w_name:
                sales_queryset = sales_queryset.filter(Q(product__name__icontains=w_name))

            if w_inventory_status != 'W':
                sales_queryset = sales_queryset.filter(Q(status=w_inventory_status))

            if s_name:
                sales_queryset = sales_queryset.filter(Q(product__name__icontains=s_name))

            if s_inventory_status != 'W':
                sales_queryset = sales_queryset.filter(Q(status=s_inventory_status))

        purchase_data = purchase_queryset.order_by('-pk')
        sales_data = sales_queryset.order_by('-pk')

        # Contruct display data
        p_result = []
        p_contract_id_list = []
        prev_contract_id = None
        for p in purchase_data.all():
            pq = p.quantity
            if not prev_contract_id or prev_contract_id != p.content_object.contract_id:
                prev_contract_id = p.content_object.contract_id
                same_order = 0
            report_date = json.loads(p.report_date) if p.report_date else {}

            while True:
                if p.content_object == None: break
                deliver_place = ''
                if p.content_type_id == phid:
                    if p.content_object.hall:
                        deliver_place = p.content_object.hall.name

                item = {
                    "id": p.id,
                    "same_order": str(same_order),
                    "contract_id": p.content_object.contract_id,
                    "content_type_id": p.content_type_id,
                    "object_id": p.content_object.id,
                    "contract_date": p.content_object.created_at,
                    "customer_name": p.content_object.customer.name if p.content_object.customer else '',
                    "delivered_place": deliver_place,
                    "person_in_charge": p.content_object.person_in_charge,
                    "payment_date": p.content_object.shipping_date, # if p.content_type_id == phid else '',
                    "product_name": p.product.name if p.product else '',
                    "quantity": 1 if p.quantity > 0 else 0,
                    "amount": p.amount,
                    "price": p.price,
                    "inventory_status": p.get_status_display(),
                    "report_date": datetime.strptime(report_date[str(same_order)], '%Y/%m/%d').date() if str(same_order) in report_date else '',
                    'linked': False
                }
                p_contract_id_list.append(str(p.content_object.contract_id) + '-' + str(p.id) + str(same_order))
                p_result.append(item)
                pq -= 1
                same_order += 1
                if pq <= 0: break

        s_result = []
        s_contract_id_list = []
        prev_contract_id = None
        for s in sales_data.all():
            sq = s.quantity
            if not prev_contract_id or prev_contract_id != s.content_object.contract_id:
                prev_contract_id = s.content_object.contract_id
                same_order = 0
                
            report_date = json.loads(s.report_date) if s.report_date else {}
            while True:
                if s.content_object == None: break
                deliver_place = ''
                if s.content_type_id == shid:
                    if s.content_object.hall:
                        deliver_place = s.content_object.hall.name

                item = {
                    "id": s.id,
                    "same_order": str(same_order),
                    "contract_id": s.content_object.contract_id,
                    "content_type_id": s.content_type_id,
                    "object_id": s.content_object.id,
                    "contract_date": s.content_object.created_at,
                    "customer_name": s.content_object.customer.name if s.content_object.customer else '',
                    "delivered_place": deliver_place,
                    "person_in_charge": s.content_object.person_in_charge,
                    "payment_date": s.content_object.shipping_date if s.content_type_id == shid else '',
                    "product_name": s.product.name if s.product else '',
                    "quantity": 1 if s.quantity > 0 else 0,
                    "amount": s.amount,
                    "price": s.price,
                    "inventory_status": s.get_status_display(),
                    "report_date": datetime.strptime(report_date[str(same_order)], '%Y/%m/%d').date() if str(same_order) in report_date else '',
                    'linked': False
                }
                s_contract_id_list.append(str(s.content_object.contract_id) + '-' + str(s.id) + str(same_order))
                s_result.append(item)

                sq -= 1
                same_order += 1
                if sq <= 0: break;

        # Pagination
        paginator1 = Paginator(p_result, ITEMS_PER_PAGE)
        paginator2 = Paginator(s_result, ITEMS_PER_PAGE)

        params = self.request.GET.copy()
        first_page_num = params['page1'] if 'page1' in params else 1
        second_page_num = params['page2'] if 'page2' in params else 1

        # Name lists to show

        # Already exists
        traderLinks = TraderLink.objects.order_by('-pk').all()
        pids = [(pid.purchase_contract_id, pid.purchase_same_order) for pid in traderLinks]
        for idx, p in enumerate(p_result):
            p_result[idx]['linked'] = (int(p['id']), int(p['same_order'])) in pids

        sids = [(sid.sale_contract_id, sid.sale_same_order) for sid in traderLinks]
        for idx, s in enumerate(s_result):
            s_result[idx]['linked'] = (int(s['id']), int(s['same_order'])) in sids
            
        return [paginator1.page(first_page_num), paginator2.page(second_page_num), paginator1, paginator2, p_contract_id_list, s_contract_id_list, p_result, s_result]

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)

        context['p_hall_contract_id'] = ContentType.objects.get(model='hallpurchasescontract').id
        context['s_hall_contract_id'] = ContentType.objects.get(model='hallsalescontract').id
        
        params = self.request.GET.copy()
        for k, v in params.items():
            if v:
                context[k] = v

        context['paginator'] = [context['object_list'][2], context['object_list'][3]]
        context['page_obj'] = [context['object_list'][0], context['object_list'][1]]
        context['is_paginated'] = True

        context['p_whole_check_items'] = context['object_list'][4]
        context['s_whole_check_items'] = context['object_list'][5]

        return context
    
    def get_link_queryset(self, p_id, s_id):
        purchase_trader_class_id = ContentType.objects.get(model='traderpurchasescontract').id
        purchase_hall_class_id = ContentType.objects.get(model='hallpurchasescontract').id
        purchase_queryset = ContractProduct.objects.filter(
            Q(content_type_id=purchase_trader_class_id) |
            Q(content_type_id=purchase_hall_class_id)
        ).order_by('-pk')

        purchase_queryset = purchase_queryset.filter(
            Q(trader_purchases_contract__contract_id__icontains=p_id) |
            Q(hall_purchases_contract__contract_id__icontains=p_id)
        )

        sales_trader_class_id = ContentType.objects.get(model='tradersalescontract').id
        sales_hall_class_id = ContentType.objects.get(model='hallsalescontract').id
        sale_queryset = ContractProduct.objects.filter(
            Q(content_type_id=sales_trader_class_id) |
            Q(content_type_id=sales_hall_class_id)
        ).order_by('-pk')

        sale_queryset = sale_queryset.filter(
            Q(trader_sales_contract__contract_id__icontains=s_id) |
            Q(hall_sales_contract__contract_id__icontains=s_id)
        )
        try:
            purchase_queryset = purchase_queryset.get()
        except Exception:
            purchase_queryset = None
        try:
            sale_queryset = sale_queryset.get()
        except Exception:
            sale_queryset = None
        return [purchase_queryset, sale_queryset]

    def post(self, request, *args, **kwargs):
        income_items = json.loads(self.request.POST.get('purchase-printable-items'))
        outcome_items = json.loads(self.request.POST.get('sale-printable-items'))

        user_id = self.request.user.id
        log_export_operation(user_id, "{} - {}".format(_("List"), _("Links")))

        response = HttpResponse(content_type='application/ms-excel')
        response['Content-Disposition'] = 'attachment; filename="listing_Links_{}.xls"'.format(generate_random_number())
        
        p_id = self.request.POST.get('excel_purchase_contract_id')
        s_id = self.request.POST.get('excel_sale_contract_id')
        # if(not p_id or not s_id): return response
        if(not p_id): p_id = '000000000000000'
        if(not s_id): s_id = '000000000000000'
        
        wb = xlwt.Workbook(encoding='utf-8')
        ws = wb.add_sheet("{} - {}".format(_('List'), _('Links')))

        ws.set_header_str(str.encode(''))
        ws.set_footer_str(str.encode(''))
        ws.set_left_margin(0.314961)
        ws.set_top_margin(0.393701)
        ws.set_right_margin(0.314961)
        ws.set_bottom_margin(0.19685)

        ws.row(0).height_mismatch = True
        ws.row(0).height = header_height

        ws.write(0, 0, _('Contract ID'), bold_style)
        ws.write(0, 1, _('Contract date'), bold_style)
        ws.write(0, 2, _('Customer'), bold_style)
        ws.write(0, 3, _('Delivered place'), bold_style)
        ws.write(0, 4, _('Person in charge'), bold_style)
        ws.write(0, 5, _('Payment date'), bold_style)
        ws.write(0, 6, _('Product name'), bold_style)
        ws.write(0, 7, _('Unit count'), bold_style)
        ws.write(0, 8, _('Amount'), bold_style)
        ws.write(0, 9, _('Inventory status'), bold_style)
        ws.write(0, 10, _('Report Date'), bold_style)

        for i in range(11):
            ws.col(i).width = cell_width
        ws.col(2).width = ws.col(3).width = ws.col(6).width = wide_cell_width

        row_no = 1
        date_style.num_format_str = 'yyyy/mm/dd' if self.request.LANGUAGE_CODE == 'ja' else 'mm/dd/yyyy'
        sets = self.get_queryset()

        for p in sets[6]:
            if income_items[str(p['contract_id']) + '-' + str(p['same_order'])]:
                ws.write(row_no, 0, p['contract_id'], common_style)
                ws.write(row_no, 1, p['contract_date'], date_style)
                ws.write(row_no, 2, p['customer_name'], common_style)
                ws.write(row_no, 3, p['delivered_place'], common_style)
                ws.write(row_no, 4, p['person_in_charge'], common_style)
                ws.write(row_no, 5, p['payment_date'], date_style)
                ws.write(row_no, 6, p['product_name'], common_style)
                ws.write(row_no, 7, p['quantity'], common_style)
                ws.write(row_no, 8, p['amount'], common_style)
                ws.write(row_no, 9, p['inventory_status'], common_style)
                ws.write(row_no, 10, p['report_date'], date_style)
                row_no += 1

        for s in sets[7]:
            if outcome_items[str(s['contract_id']) + '-' + str(s['same_order'])]:
                ws.write(row_no, 0, s['contract_id'], common_style)
                ws.write(row_no, 1, s['contract_date'], date_style)
                ws.write(row_no, 2, s['customer_name'], common_style)
                ws.write(row_no, 3, s['delivered_place'], common_style)
                ws.write(row_no, 4, s['person_in_charge'], common_style)
                ws.write(row_no, 5, s['payment_date'], date_style)
                ws.write(row_no, 6, s['product_name'], common_style)
                ws.write(row_no, 7, s['quantity'], common_style)
                ws.write(row_no, 8, s['amount'], common_style)
                ws.write(row_no, 9, s['inventory_status'], common_style)
                ws.write(row_no, 10, s['report_date'], date_style)
                row_no += 1
        
        wb.save(response)
        return response
    

class ExportHistoryListView(AdminLoginRequiredMixin, ListView):
    template_name = 'listing/history.html'
    context_object_name = 'objects'
    paginate_by = 10

    def get_queryset(self):
        return ExportHistory.objects.order_by('-exported_at')
        # return ProductFilter(self.request.GET, queryset=self.queryset).qs.order_by('pk')
    
    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        # context['history_filter'] = HistoryFilter(self.request.GET)
        return context


class InventoryProductCreateView(AdminLoginRequiredMixin, View):
    def post(self, request, *args, **kwargs):
        product_form = ProductForm(request.POST)
        if product_form.is_valid():
            product_form.save()
        return redirect('listing:inventory-list')


class InventoryProductUpdateView(AdminLoginRequiredMixin, View):
    def post(self, request, *args, **kwargs):
        id = request.POST.get('id')
        product = Product.objects.get(id=id)
        # product_form = ProductForm(request.POST)
        # if product_form.is_valid():
            # data = product_form.cleaned_data
            # product.name = data.get('name')
            # product.identifier = data.get('identifier')
            # product.purchase_date = data.get('purchase_date')
            # product.supplier = data.get('supplier')
            # product.person_in_charge = data.get('person_in_charge')
            # product.quantity = data.get('quantity')
            # product.price = data.get('price')
            # product.stock = data.get('stock')
            # product.amount = data.get('amount')
            # product.save()
        
        product.quantity = request.POST.get('quantity')
        product.save()
        return redirect('listing:inventory-list')


class InventoryProductDeleteView(AdminLoginRequiredMixin, View):
    def post(self, request, *args, **kwargs):
        id = request.POST.get('id')
        product = Product.objects.get(id=id)
        product.delete()
        return redirect('listing:inventory-list')


class InventoryProductDetailAjaxView(AdminLoginRequiredMixin, View):
    def post(self, *args, **kwargs):
        if self.request.method == 'POST' and self.request.is_ajax():
            id = self.request.POST.get('id')
            product = Product.objects.get(id=id)
            return JsonResponse({
                # 'name': product.name,
                # 'identifier': product.identifier,
                # 'purchase_date': date_dump(product.purchase_date, self.request.LANGUAGE_CODE),
                # 'supplier': product.supplier,
                # 'person_in_charge': product.person_in_charge,
                'quantity': product.quantity,
                # 'price': product.price,
                # 'stock': product.stock,
                # 'amount': product.amount,
            }, status=200)
        return JsonResponse({'success': False}, status=400)


class SalesProductUpdateView(AdminLoginRequiredMixin, View):
    def post(self, request, *args, **kwargs):
        id = request.POST.get('id')
        status = request.POST.get('status')
        product = ContractProduct.objects.get(id=id)
        product.status = status
        product.save()

        origin = request.POST.get('origin')
        if origin == 'link':
            return redirect('listing:link-list')
        elif origin == 'linkshow':
            return redirect('listing:link-list-show')
            
        return redirect('listing:sales-list')


class PurchasesProductUpdateView(AdminLoginRequiredMixin, View):
    def post(self, request, *args, **kwargs):
        id = request.POST.get('id')
        status = request.POST.get('status')
        product = ContractProduct.objects.get(id=id)
        product.status = status
        product.save()

        origin = request.POST.get('origin')

        if origin == 'link':
            return redirect('listing:link-list')
        elif origin == 'linkshow':
            return redirect('listing:link-list-show')
        return redirect('listing:purchases-list')


class SalesProductDetailAjaxView(AdminLoginRequiredMixin, View):
    def post(self, *args, **kwargs):
        if self.request.method == 'POST' and self.request.is_ajax():
            id = self.request.POST.get('id')
            product = ContractProduct.objects.get(id=id)
            contract = product.content_object
            contract_id = contract.contract_id
            product_name = product.product.name
            status = product.status
            return JsonResponse({
                'contract_id': contract_id,
                'product_name': product_name,
                'status': status
            }, status=200)
        return JsonResponse({'success': False}, status=400)


class PurchasesProductDetailAjaxView(AdminLoginRequiredMixin, View):
    def post(self, *args, **kwargs):
        if self.request.method == 'POST' and self.request.is_ajax():
            id = self.request.POST.get('id')
            product = ContractProduct.objects.get(id=id)
            contract = product.content_object
            contract_id = contract.contract_id
            product_name = product.product.name
            status = product.status
            return JsonResponse({
                'contract_id': contract_id,
                'product_name': product_name,
                'status': status
            }, status=200)
        return JsonResponse({'success': False}, status=400)
