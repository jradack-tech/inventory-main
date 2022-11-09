from django.shortcuts import render, redirect
from django.contrib.contenttypes.models import ContentType
from django.views.generic.base import TemplateView, View
from django.http import HttpResponse, JsonResponse
from django.utils.translation import gettext as _
from users.views import AdminLoginRequiredMixin
from masterdata.models import Sender, Document, DocumentFee, PersonInCharge
from .models import (
    ContractProduct, ContractDocument, ContractDocumentFee, Milestone,
    TraderSalesContract, TraderPurchasesContract, HallSalesContract, HallPurchasesContract,
)
from .forms import (
    TraderSalesContractForm, TraderPurchasesContractForm, HallSalesContractForm, HallPurchasesContractForm,
    ProductFormSet, DocumentFormSet, DocumentFeeFormSet, MilestoneFormSet, MilestoneValidationFormSet,
    TraderSalesProductSenderForm, TraderSalesDocumentSenderForm, TraderPurchasesProductSenderForm, TraderPurchasesDocumentSenderForm,
    ProductForm, DocumentForm, DocumentFeeForm, MilestoneForm,
)
from .utilities import generate_contract_id
from datetime import datetime
import math

class TraderSalesContractUpdateView(AdminLoginRequiredMixin, TemplateView):
    template_name = 'contracts/trader_sales.html'

    def get(self, request, *args, **kwargs):
        return render(request, self.template_name, self.get_context_data(**kwargs))

    def post(self, request, *args, **kwargs):
        id = kwargs.get('pk')
        contract_form = TraderSalesContractForm(self.request.POST, id=id)
        if contract_form.is_valid():
            contract = contract_form.save()
        else:
            return render(request, self.template_name, self.get_context_data(**kwargs))
        
        # shipping_method = contract_form.cleaned_data.get('shipping_method')
        # if shipping_method == 'R':
        #     product_sender_form = TraderSalesProductSenderForm(self.request.POST, contract_id=contract.id)
        #     if product_sender_form.is_valid():
        #         product_sender_form.save()
        #     document_sender_form = TraderSalesDocumentSenderForm(self.request.POST, contract_id=contract.id)
        #     if document_sender_form.is_valid():
        #         document_sender_form.save()
        
        product_sender_form = TraderSalesProductSenderForm(self.request.POST, contract_id=contract.id)
        if product_sender_form.is_valid():
            product_sender_form.save()
        document_sender_form = TraderSalesDocumentSenderForm(self.request.POST, contract_id=contract.id)
        if document_sender_form.is_valid():
            document_sender_form.save()
        
        product_formset = ProductFormSet(
            self.request.POST,
            form_kwargs={'contract_id': contract.id, 'contract_class': 'TraderSalesContract'},
            prefix='product'
        )
        if product_formset.is_valid():
            product_qs = contract.products.values('id')
            old_product_ids = set()
            for product in product_qs:
                old_product_ids.add(product['id'])
            updated_ids = set()
            for form in product_formset.forms:
                if form.is_valid():
                    if form.cleaned_data.get('id'):
                        updated_ids.add(form.cleaned_data.get('id'))
                    form.save()
            for product_id in old_product_ids:
                if product_id not in updated_ids:
                    ContractProduct.objects.get(id=product_id).delete()
        
        document_formset = DocumentFormSet(
            self.request.POST,
            form_kwargs={'contract_id': contract.id, 'contract_class': 'TraderSalesContract'},
            prefix='document'
        )
        if document_formset.is_valid():
            document_qs = contract.documents.values('id')
            old_document_ids = set()
            for document in document_qs:
                old_document_ids.add(document['id'])
            updated_ids = set()
            for form in document_formset.forms:
                if form.is_valid():
                    if form.cleaned_data.get('id'):
                        updated_ids.add(form.cleaned_data.get('id'))
                    form.save()
            for document_id in old_document_ids:
                if document_id not in updated_ids:
                    ContractDocument.objects.get(id=document_id).delete()
         
        return render(request, self.template_name, self.get_context_data(**kwargs))
    
    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        id = kwargs.get('pk')
        contract = TraderSalesContract.objects.get(id=id)
        contract_data = {
            'contract_id': contract.contract_id,
            'created_at': contract.created_at,
            'updated_at': contract.updated_at,
            'customer_id': contract.customer.id if contract.customer else None,
            'customer_name': contract.customer.name if contract.customer else None,
            'manager': contract.manager,
            'frigana': contract.customer.frigana if contract.customer else None,
            'postal_code': contract.customer.postal_code if contract.customer else None,
            'address': contract.customer.address if contract.customer else None,
            'tel': contract.customer.tel if contract.customer else None,
            'fax': contract.customer.fax if contract.customer else None,
            'person_in_charge': contract.person_in_charge,
            'remarks': contract.remarks,
            'sub_total': f"{contract.sub_total:,}",
            'fee': f"{contract.fee:,}",
            'tax': f"{contract.tax:,}",
            'total': f"{contract.total:,}",
            'billing_amount': contract.billing_amount,
            'shipping_method': contract.shipping_method,
            'shipping_date': contract.shipping_date,
            'payment_method': contract.payment_method,
            'payment_due_date': contract.payment_due_date,
            'memo': contract.memo
        }
        context['contract_form'] = TraderSalesContractForm(contract_data)

        if contract.senders.count():
            product_sender = contract.senders.filter(type='P').first()
            document_sender = contract.senders.filter(type='D').first()
            product_sender_form = TraderSalesProductSenderForm({
                'p_id': product_sender.id,
                'product_sender_id': product_sender.sender.id if product_sender.sender else None,
                'product_sender_name': product_sender.sender.name if product_sender.sender else None,
                'product_sender_address': product_sender.sender.address if product_sender.sender else None,
                'product_sender_tel': product_sender.sender.tel if product_sender.sender else None,
                'product_sender_postal_code': product_sender.sender.postal_code if product_sender.sender and product_sender.sender.postal_code else None,
                'product_sender_fax': product_sender.sender.fax if product_sender.sender else None,
                'product_expected_arrival_date': product_sender.expected_arrival_date
            })
            if product_sender_form.is_valid():
                context['product_sender_form'] = product_sender_form
            document_sender_form = TraderSalesDocumentSenderForm({
                'd_id': document_sender.id,
                'document_sender_id': document_sender.sender.id if document_sender.sender else None,
                'document_sender_name': document_sender.sender.name if document_sender.sender else None,
                'document_sender_address': document_sender.sender.address if document_sender.sender else None,
                'document_sender_tel': document_sender.sender.tel if document_sender.sender else None,
                'document_sender_postal_code': document_sender.sender.postal_code if document_sender.sender and document_sender.sender.postal_code else None,
                'document_sender_fax': document_sender.sender.fax if document_sender.sender else None,
                'document_expected_arrival_date': document_sender.expected_arrival_date
            })
            if document_sender_form.is_valid():
                context['document_sender_form'] = document_sender_form
        else:
            context['product_sender_form'] = TraderSalesProductSenderForm()
            context['document_sender_form'] = TraderSalesDocumentSenderForm()
        
        context['documents'] = Document.objects.all().values('id', 'name')
        context['people'] = PersonInCharge.objects.all().values('name')

        product_set = []
        products = contract.products.all()
        for product in products:
            price = math.ceil(product.price / 10000) * 10000
            fee = round(200 * product.quantity * (price / 100000)) if price > 100000 else product.fee
            data = {
                'id': product.id,
                'product_id': product.product.id,
                'name': product.product.name,
                'type': product.type,
                'quantity': product.quantity,
                'price': product.price,
                'tax': product.tax,
                'fee': fee, # product.fee,
                'amount': product.amount
            }
            product_form = ProductForm(data)
            if product_form.is_valid():
                product_set.append(data)
        context['productformset'] = ProductFormSet(initial=product_set, prefix='product')

        document_set = []
        documents = contract.documents.all()
        for document in documents:
            data = {
                'id': document.id,
                'document_id': document.document.id,
                'taxable': int(document.taxable),
                'tax': document.tax,
                'name': document.document.name,
                'quantity': document.quantity,
                'price': document.price,
                'amount': document.amount
            }
            document_form = DocumentForm(data)
            if document_form.is_valid():
                document_set.append(data)
        context['documentformset'] = DocumentFormSet(initial=document_set, prefix='document')
        
        return context


class TraderPurchasesContractUpdateView(AdminLoginRequiredMixin, TemplateView):
    template_name = 'contracts/trader_purchases.html'

    def get(self, request, *args, **kwargs):
        return render(request, self.template_name, self.get_context_data(**kwargs))

    def post(self, request, *args, **kwargs):
        id = kwargs.get('pk')
        contract_form = TraderPurchasesContractForm(self.request.POST, id=id)
        if contract_form.is_valid():
            contract = contract_form.save()
        else:
            return render(request, self.template_name, self.get_context_data(**kwargs))
        
        product_sender_form = TraderPurchasesProductSenderForm(self.request.POST, contract_id=contract.id)
        if product_sender_form.is_valid():
            product_sender_form.save()
        
        document_sender_form = TraderPurchasesDocumentSenderForm(self.request.POST, contract_id=contract.id)
        if document_sender_form.is_valid():
            document_sender_form.save()
            
        product_formset = ProductFormSet(
            self.request.POST,
            form_kwargs={'contract_id': contract.id, 'contract_class': 'TraderPurchasesContract'},
            prefix='product'
        )
        if product_formset.is_valid():
            product_qs = contract.products.values('id')
            old_product_ids = set()
            for product in product_qs:
                old_product_ids.add(product['id'])
            updated_ids = set()
            for form in product_formset.forms:
                if form.is_valid():
                    if form.cleaned_data.get('id'):
                        updated_ids.add(form.cleaned_data.get('id'))
                    form.save()
            for product_id in old_product_ids:
                if product_id not in updated_ids:
                    ContractProduct.objects.get(id=product_id).delete()

        document_formset = DocumentFormSet(
            self.request.POST,
            form_kwargs={'contract_id': contract.id, 'contract_class': 'TraderPurchasesContract'},
            prefix='document'
        )
        if document_formset.is_valid():
            document_qs = contract.documents.values('id')
            old_document_ids = set()
            for document in document_qs:
                old_document_ids.add(document['id'])
            updated_ids = set()
            for form in document_formset.forms:
                if form.is_valid():
                    if form.cleaned_data.get('id'):
                        updated_ids.add(form.cleaned_data.get('id'))
                    form.save()
            for document_id in old_document_ids:
                if document_id not in updated_ids:
                    ContractDocument.objects.get(id=document_id).delete()
        
        return render(request, self.template_name, self.get_context_data(**kwargs))
    
    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        id = kwargs.get('pk')
        contract = TraderPurchasesContract.objects.get(id=id)
        contract_data = {
            'contract_id': contract.contract_id,
            'created_at': contract.created_at,
            'updated_at': contract.updated_at,
            'customer_id': contract.customer.id if contract.customer else None,
            'customer_name': contract.customer.name if contract.customer else None,
            'manager': contract.manager,
            'frigana': contract.customer.frigana if contract.customer else None,
            'postal_code': contract.customer.postal_code if contract.customer else None,
            'address': contract.customer.address if contract.customer else None,
            'tel': contract.customer.tel if contract.customer else None,
            'fax': contract.customer.fax if contract.customer else None,
            'person_in_charge': contract.person_in_charge,
            'removal_date': contract.removal_date,
            'shipping_date': contract.shipping_date,
            'frame_color': contract.frame_color,
            'receipt': contract.receipt,
            'remarks': contract.remarks,
            'sub_total': f"{contract.sub_total:,}",
            'fee': f"{contract.fee:,}",
            'tax': f"{contract.tax:,}",
            'total': f"{contract.total:,}",
            'transfer_deadline': contract.transfer_deadline,
            'bank_name': contract.bank_name,
            'account_number': contract.account_number,
            'branch_name': contract.branch_name,
            'payment_due_date': contract.payment_due_date,
            'account_holder': contract.account_holder
        }
        context['contract_form'] = TraderPurchasesContractForm(contract_data)

        product_sender = contract.senders.filter(type='P').first()
        document_sender = contract.senders.filter(type='D').first()
        product_sender_form = TraderPurchasesProductSenderForm({
            'p_id': product_sender.id,
            'product_sender_id': product_sender.sender.id if product_sender.sender else None,
            'product_sender_name': product_sender.sender.name if product_sender.sender else None,
            'product_sender_address': product_sender.sender.address if product_sender.sender else None,
            'product_sender_tel': product_sender.sender.tel if product_sender.sender else None,
            'product_desired_arrival_date': product_sender.desired_arrival_date,
            'product_shipping_company': product_sender.shipping_company,
            'product_sender_remarks': product_sender.remarks
        })
        context['product_sender_form'] = product_sender_form

        document_sender_form = TraderPurchasesDocumentSenderForm({
            'd_id': document_sender.id,
            'document_sender_id': document_sender.sender.id if document_sender.sender else None,
            'document_sender_name': document_sender.sender.name if document_sender.sender else None,
            'document_sender_address': document_sender.sender.address if document_sender.sender else None,
            'document_sender_tel': document_sender.sender.tel if document_sender.sender else None,
            'document_desired_arrival_date': document_sender.desired_arrival_date if document_sender.sender else None,
            'document_shipping_company': document_sender.shipping_company,
            'document_sender_remarks': document_sender.remarks
        })
        context['document_sender_form'] = document_sender_form

        context['documents'] = Document.objects.all().values('id', 'name')
        context['people'] = PersonInCharge.objects.all().values('name')

        product_set = []
        products = contract.products.all()
        for product in products:
            data = {
                'id': product.id,
                'product_id': product.product.id,
                'name': product.product.name,
                'type': product.type,
                'quantity': product.quantity,
                'price': product.price,
                'tax': product.tax,
                'fee': product.fee,
                'amount': product.amount
            }
            product_form = ProductForm(data)
            if product_form.is_valid():
                product_set.append(data)
        context['productformset'] = ProductFormSet(initial=product_set, prefix='product')

        document_set = []
        documents = contract.documents.all()
        for document in documents:
            data = {
                'id': document.id,
                'document_id': document.document.id,
                'taxable': int(document.taxable),
                'tax': document.tax,
                'name': document.document.name,
                'quantity': document.quantity,
                'price': document.price,
                'amount': document.amount
            }
            document_form = DocumentForm(data)
            if document_form.is_valid():
                document_set.append(data)
        context['documentformset'] = DocumentFormSet(initial=document_set, prefix='document')

        return context


class HallSalesContractUpdateView(AdminLoginRequiredMixin, TemplateView):
    template_name = 'contracts/hall_sales.html'

    def get(self, request, *args, **kwargs):
        return render(request, self.template_name, self.get_context_data(**kwargs))
    
    def post(self, request, *args, **kwargs):
        id = kwargs.get('pk')
        contract_form = HallSalesContractForm(self.request.POST, id=id)
        # mdate = self.request.POST.dict()['milestone-0-date']
        # try: 
        #     try:
        #         mdate = datetime.strptime(mdate, '%m/%d/%Y')
        #     except:
        #         mdate = datetime.strptime(mdate, '%Y/%m/%d')
        # except:
        #     mdate = None
        if contract_form.is_valid():
            contract = contract_form.save()
        else:
            return render(request, self.template_name, self.get_context_data(**kwargs))
            
        product_formset = ProductFormSet(
            self.request.POST,
            form_kwargs={'contract_id': contract.id, 'contract_class': 'HallSalesContract'},
            prefix='product'
        )
        if product_formset.is_valid():
            product_qs = contract.products.values('id')
            old_product_ids = set()
            for product in product_qs:
                old_product_ids.add(product['id'])
            updated_ids = set()
            for form in product_formset.forms:
                if form.is_valid():
                    if form.cleaned_data.get('id'):
                        updated_ids.add(form.cleaned_data.get('id'))
                    form.save()
            for product_id in old_product_ids:
                if product_id not in updated_ids:
                    ContractProduct.objects.get(id=product_id).delete()

        document_formset = DocumentFormSet(
            self.request.POST,
            form_kwargs={'contract_id': contract.id, 'contract_class': 'HallSalesContract'},
            prefix='document'
        )
        if document_formset.is_valid():
            document_qs = contract.documents.values('id')
            old_document_ids = set()
            for document in document_qs:
                old_document_ids.add(document['id'])
            updated_ids = set()
            for form in document_formset.forms:
                if form.is_valid():
                    if form.cleaned_data.get('id'):
                        updated_ids.add(form.cleaned_data.get('id'))
                    form.save()
            for document_id in old_document_ids:
                if document_id not in updated_ids:
                    ContractDocument.objects.get(id=document_id).delete()
        
        document_fee_formset = DocumentFeeFormSet(
            self.request.POST,
            form_kwargs={'contract_id': contract.id, 'contract_class': 'HallSalesContract'},
            prefix='document_fee'
        )
        if document_fee_formset.is_valid():
            document_fee_qs = contract.document_fees.values('id')
            old_document_fee_ids = set()
            for document_fee in document_fee_qs:
                old_document_fee_ids.add(document_fee['id'])
            updated_ids = set()
            for form in document_fee_formset.forms:
                if form.is_valid():
                    if form.cleaned_data.get('id'):
                        updated_ids.add(form.cleaned_data.get('id'))
                    form.save()
            for document_fee_id in old_document_fee_ids:
                if document_fee_id not in updated_ids:
                    ContractDocumentFee.objects.get(id=document_fee_id).delete()
        
        milestone_formset = MilestoneFormSet(
            self.request.POST,
            form_kwargs={'contract_id': contract.id, 'contract_class': 'HallSalesContract'},
            prefix='milestone'
        )
        for form in milestone_formset.forms:
            if form.is_valid():
                form.save()
        
        return render(request, self.template_name, self.get_context_data(**kwargs))

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        id = kwargs.get('pk')
        contract = HallSalesContract.objects.get(id=id)

        contract_data = {
            'contract_id': contract.contract_id,
            'created_at': contract.created_at,
            'customer_id': contract.customer.id if contract.customer else None,
            # 'customer_name': contract.customer.name if contract.customer else None,
            "customer_name": contract.hall.customer_name if contract.hall else None,
            'hall_id': contract.hall.id if contract.hall else None,
            'hall_name': contract.hall.name if contract.hall else None,
            'hall_address': contract.hall.address if contract.hall else None,
            'hall_tel': contract.hall.tel if contract.hall else None,
            'remarks': contract.remarks,
            'sub_total': f"{contract.sub_total:,}",
            'tax': f"{contract.tax:,}",
            'fee': f"{contract.fee:,}",
            'total': f"{contract.total:,}",
            'shipping_date': contract.shipping_date,
            'opening_date': contract.opening_date,
            'payment_method': contract.payment_method,
            'transfer_account': contract.transfer_account,
            'person_in_charge': contract.person_in_charge,
            'confirmor': contract.confirmor,
        }
        context['contract_form'] = HallSalesContractForm(contract_data)
        context['update_hall_id'] = contract.hall.id if contract.hall else None
        context['documents'] = Document.objects.all().values('id', 'name')
        document_fee_lambda = lambda df: {'id': df.id, 'name': df.get_type_display()}
        document_fees = [document_fee_lambda(document_fee) for document_fee in DocumentFee.objects.all()]
        context['document_fees'] = document_fees
        context['senders'] = Sender.objects.all().values('id', 'name')

        product_set = []
        products = contract.products.all()
        for product in products:
            data = {
                'id': product.id,
                'product_id': product.product.id,
                'name': product.product.name,
                'type': product.type,
                'quantity': product.quantity,
                'price': product.price,
                'tax': product.tax,
                'fee': product.fee,
                'amount': product.amount
            }
            product_form = ProductForm(data)
            if product_form.is_valid():
                product_set.append(data)
        context['productformset'] = ProductFormSet(initial=product_set, prefix='product')
        
        document_set = []
        documents = contract.documents.all()
        for document in documents:
            data = {
                'id': document.id,
                'document_id': document.document.id,
                'taxable': int(document.taxable),
                'tax': document.tax,
                'name': document.document.name,
                'quantity': document.quantity,
                'price': document.price,
                'amount': document.amount
            }
            document_form = DocumentForm(data)
            if document_form.is_valid():
                document_set.append(data)
        context['documentformset'] = DocumentFormSet(initial=document_set, prefix='document')

        document_fee_set = []
        document_fees = contract.document_fees.all()
        for document_fee in document_fees:
            data = {
                'id': document_fee.id,
                'document_fee_id': document_fee.document_fee.id,
                'name': document_fee.document_fee.get_type_display,
                'tax': document_fee.tax,
                'model_price': document_fee.document_fee.model_price,
                'unit_price': document_fee.document_fee.unit_price,
                'application_fee': document_fee.document_fee.application_fee,
                'model_count': document_fee.model_count,
                'unit_count': document_fee.unit_count,
                'amount': document_fee.amount
            }
            document_fee_form = DocumentFeeForm(data)
            if document_fee_form.is_valid():
                document_fee_set.append(data)
        context['documentfeeformset'] = DocumentFeeFormSet(initial=document_fee_set, prefix='document_fee')
        milestone_set = []
        milestones = contract.milestones.all()
        for milestone in milestones:
            data = {
                'id': milestone.id,
                'date': milestone.date,
                'amount': milestone.amount
            }
            milestone_form = MilestoneForm(data)
            if milestone_form.is_valid():
                milestone_set.append(data)
        context['milestoneformset'] = MilestoneFormSet(initial=milestone_set, prefix='milestone')
        
        context['people'] = PersonInCharge.objects.all().values('name')

        return context


class HallPurchasesContractUpdateView(AdminLoginRequiredMixin, TemplateView):
    template_name = 'contracts/hall_purchases.html'

    def get(self, request, *args, **kwargs):
        return render(request, self.template_name, self.get_context_data(**kwargs))
    
    def post(self, request, *args, **kwargs):
        id = kwargs.get('pk')
        contract_form = HallPurchasesContractForm(self.request.POST, id=id)
        # mdate = self.request.POST.dict()['milestone-0-date']
        # try:
        #     try:
        #         mdate = datetime.strptime(mdate, '%m/%d/%Y')
        #     except:
        #         mdate = datetime.strptime(mdate, '%Y/%m/%d')
        # except:
        #     mdate = None
        if contract_form.is_valid():
            contract = contract_form.save()
            
        product_formset = ProductFormSet(
            self.request.POST,
            form_kwargs={'contract_id': contract.id, 'contract_class': 'HallPurchasesContract'},
            prefix='product'
        )
        if product_formset.is_valid():
            product_qs = contract.products.values('id')
            old_product_ids = set()
            for product in product_qs:
                old_product_ids.add(product['id'])
            updated_ids = set()
            for form in product_formset.forms:
                if form.is_valid():
                    if form.cleaned_data.get('id'):
                        updated_ids.add(form.cleaned_data.get('id'))
                    form.save()
            for product_id in old_product_ids:
                if product_id not in updated_ids:
                    ContractProduct.objects.get(id=product_id).delete()

        document_formset = DocumentFormSet(
            self.request.POST,
            form_kwargs={'contract_id': contract.id, 'contract_class': 'HallPurchasesContract'},
            prefix='document'
        )
        if document_formset.is_valid():
            document_qs = contract.documents.values('id')
            old_document_ids = set()
            for document in document_qs:
                old_document_ids.add(document['id'])
            updated_ids = set()
            for form in document_formset.forms:
                if form.is_valid():
                    if form.cleaned_data.get('id'):
                        updated_ids.add(form.cleaned_data.get('id'))
                    form.save()
            for document_id in old_document_ids:
                if document_id not in updated_ids:
                    ContractDocument.objects.get(id=document_id).delete()
        
        document_fee_formset = DocumentFeeFormSet(
            self.request.POST,
            form_kwargs={'contract_id': contract.id, 'contract_class': 'HallPurchasesContract'},
            prefix='document_fee'
        )
        if document_fee_formset.is_valid():
            document_fee_qs = contract.document_fees.values('id')
            old_document_fee_ids = set()
            for document_fee in document_fee_qs:
                old_document_fee_ids.add(document_fee['id'])
            updated_ids = set()
            for form in document_fee_formset.forms:
                if form.is_valid():
                    if form.cleaned_data.get('id'):
                        updated_ids.add(form.cleaned_data.get('id'))
                    form.save()
            for document_fee_id in old_document_fee_ids:
                if document_fee_id not in updated_ids:
                    ContractDocumentFee.objects.get(id=document_fee_id).delete()
        
        milestone_formset = MilestoneFormSet(
            self.request.POST,
            form_kwargs={'contract_id': contract.id, 'contract_class': 'HallPurchasesContract'},
            prefix='milestone'
        )
        for form in milestone_formset.forms:
            if form.is_valid():
                form.save()
        
        return render(request, self.template_name, self.get_context_data(**kwargs))

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        id = kwargs.get('pk')
        contract = HallPurchasesContract.objects.get(id=id)
        contract_data = {
            'contract_id': contract.contract_id,
            'created_at': contract.created_at,
            'customer_id': contract.customer.id if contract.customer else None,
            # 'customer_name': contract.customer.name if contract.customer else None,
            "customer_name": contract.hall.customer_name if contract.hall else None,
            'hall_id': contract.hall.id if contract.hall else None,
            'hall_name': contract.hall.name if contract.hall else None,
            'hall_address': contract.hall.address if contract.hall else None,
            'hall_tel': contract.hall.tel if contract.hall else None,
            'remarks': contract.remarks,
            'sub_total': f"{contract.sub_total:,}",
            'tax': f"{contract.tax:,}",
            'fee': f"{contract.fee:,}",
            'total': f"{contract.total:,}",
            'shipping_date': contract.shipping_date,
            'opening_date': contract.opening_date,
            'payment_method': contract.payment_method,
            'transfer_account': contract.transfer_account,
            'person_in_charge': contract.person_in_charge,
            'confirmor': contract.confirmor,
            'memo': contract.memo,
        }
        context['contract_form'] = HallPurchasesContractForm(contract_data)
        context['update_hall_id'] = contract.hall.id if contract.hall else None
        context['documents'] = Document.objects.all().values('id', 'name')
        document_fee_lambda = lambda df: {'id': df.id, 'name': df.get_type_display()}
        document_fees = [document_fee_lambda(document_fee) for document_fee in DocumentFee.objects.all()]
        context['document_fees'] = document_fees
        context['senders'] = Sender.objects.all().values('id', 'name')

        product_set = []
        products = contract.products.all()
        for product in products:
            data = {
                'id': product.id,
                'product_id': product.product.id,
                'name': product.product.name,
                'type': product.type,
                'quantity': product.quantity,
                'price': product.price,
                'tax': product.tax,
                'fee': product.fee,
                'amount': product.amount
            }
            product_form = ProductForm(data)
            if product_form.is_valid():
                product_set.append(data)
        context['productformset'] = ProductFormSet(initial=product_set, prefix='product')
        
        document_set = []
        documents = contract.documents.all()
        for document in documents:
            data = {
                'id': document.id,
                'document_id': document.document.id,
                'taxable': int(document.taxable),
                'tax': document.tax,
                'name': document.document.name,
                'quantity': document.quantity,
                'price': document.price,
                'amount': document.amount
            }
            document_form = DocumentForm(data)
            if document_form.is_valid():
                document_set.append(data)
        context['documentformset'] = DocumentFormSet(initial=document_set, prefix='document')

        document_fee_set = []
        document_fees = contract.document_fees.all()
        for document_fee in document_fees:
            data = {
                'id': document_fee.id,
                'document_fee_id': document_fee.document_fee.id,
                'name': document_fee.document_fee.get_type_display,
                'tax': document_fee.tax,
                'model_price': document_fee.document_fee.model_price,
                'unit_price': document_fee.document_fee.unit_price,
                'application_fee': document_fee.document_fee.application_fee,
                'model_count': document_fee.model_count,
                'unit_count': document_fee.unit_count,
                'amount': document_fee.amount
            }
            document_fee_form = DocumentFeeForm(data)
            if document_fee_form.is_valid():
                document_fee_set.append(data)
        context['documentfeeformset'] = DocumentFeeFormSet(initial=document_fee_set, prefix='document_fee')
        milestone_set = []
        milestones = contract.milestones.all()
        for milestone in milestones:
            data = {
                'id': milestone.id,
                'date': milestone.date,
                'amount': milestone.amount
            }
            milestone_form = MilestoneForm(data)
            if milestone_form.is_valid():
                milestone_set.append(data)
        context['milestoneformset'] = MilestoneFormSet(initial=milestone_set, prefix='milestone')

        context['people'] = PersonInCharge.objects.all().values('name')

        return context
