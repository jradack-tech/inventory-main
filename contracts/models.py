from django.db import models
from django.contrib.contenttypes.models import ContentType
from django.contrib.contenttypes.fields import GenericRelation, GenericForeignKey
from masterdata.models import (
    Customer, Hall, Sender, Product, Document, DocumentFee,
    PRODUCT_TYPE_CHOICES, STOCK_CHOICES, SHIPPING_METHOD_CHOICES,
    PAYMENT_METHOD_CHOICES, ITEM_CHOICES, THRESHOLD_PRICE,
)
class ContractProduct(models.Model):
    product = models.ForeignKey(Product, on_delete=models.SET_NULL, null=True)
    type = models.CharField(max_length=1, choices=PRODUCT_TYPE_CHOICES)
    quantity = models.PositiveIntegerField()
    price = models.PositiveIntegerField()
    status = models.CharField(max_length=1, choices=STOCK_CHOICES, default='P')
    content_type = models.ForeignKey(ContentType, on_delete=models.CASCADE)
    object_id = models.PositiveIntegerField()
    content_object = GenericForeignKey('content_type', 'object_id')

    report_date = models.CharField(max_length=1024, null=True)
    available = models.CharField(max_length=1, null=True, blank=True, default='T')

    @property
    def amount(self):
        return self.quantity * self.price
    
    @property
    def tax(self):
        return int(self.amount * 0.1)

    @property
    def fee(self):
        price = round(self.price / 1000) * 1000
        if price > THRESHOLD_PRICE:
            return int(200 * self.quantity * (price / THRESHOLD_PRICE))
        else:
            return 100 * self.quantity
            # return self.quantity

class ContractDocument(models.Model):
    document = models.ForeignKey(Document, on_delete=models.SET_NULL, null=True)
    quantity = models.PositiveIntegerField()
    price = models.PositiveIntegerField()
    content_type = models.ForeignKey(ContentType, on_delete=models.CASCADE)
    object_id = models.PositiveIntegerField()
    content_object = GenericForeignKey('content_type', 'object_id')

    @property
    def amount(self):
        return self.quantity * self.price
    
    @property
    def taxable(self):
        return self.document.taxable
        
    @property
    def tax(self):
        if self.taxable:
            return int(self.amount * 0.1)
        return 0

class ContractDocumentFee(models.Model):
    document_fee = models.ForeignKey(DocumentFee, on_delete=models.SET_NULL, null=True)
    model_count = models.PositiveIntegerField()
    unit_count = models.PositiveIntegerField()
    content_type = models.ForeignKey(ContentType, on_delete=models.CASCADE)
    object_id = models.PositiveIntegerField()
    content_object = GenericForeignKey('content_type', 'object_id')

    @property
    def amount(self):
        model_price = self.document_fee.model_price
        unit_price = self.document_fee.unit_price
        application_fee = self.document_fee.application_fee
        return self.model_count * model_price + self.unit_count * unit_price + application_fee
    
    @property
    def tax(self):
        return int(self.amount * 0.1)

class Milestone(models.Model):
    date = models.DateField(null=True)
    amount = models.PositiveIntegerField(null=True)
    content_type = models.ForeignKey(ContentType, on_delete=models.CASCADE)
    object_id = models.PositiveIntegerField()
    content_object = GenericForeignKey('content_type', 'object_id')

class TraderContract(models.Model):
    contract_id = models.CharField(max_length=200)
    created_at = models.DateField()
    updated_at = models.DateField()
    customer = models.ForeignKey(Customer, on_delete=models.SET_NULL, null=True)
    manager = models.CharField(max_length=200, null=True, blank=True)
    person_in_charge = models.CharField(max_length=200, null=True, blank=True)
    remarks = models.TextField(null=True, blank=True)
    shipping_date = models.DateField(null=True, blank=True)
    fee = models.IntegerField(default=0)

    available = models.CharField(max_length=1, null=True, blank=True, default='T')

    class Meta:
        abstract = True

    @property
    def sub_total(self):
        sum = 0
        for product in self.products.all():
            sum += product.amount
        for document in self.documents.all():
            sum += document.amount
        return sum
    
    @property
    def tax(self):
        sum = 0
        for product in self.products.all():
            sum += product.tax
        for document in self.documents.all():
            sum += document.tax
        return sum

    @property
    def total(self):
        return self.sub_total + self.tax + self.fee
    
    @property
    def billing_amount(self):
        return self.total

    @property
    def taxed_total(self):
        return self.total - self.fee

class TraderSalesContract(TraderContract):
    shipping_method = models.CharField(max_length=1, choices=SHIPPING_METHOD_CHOICES)
    payment_method = models.CharField(max_length=2, choices=PAYMENT_METHOD_CHOICES)
    payment_due_date = models.DateField(null=True, blank=True)
    memo = models.TextField(null=True, blank=True)
    products = GenericRelation(ContractProduct, related_query_name='trader_sales_contract')
    documents = GenericRelation(ContractDocument, related_query_name='trader_sales_contract')

    sale_print_date = models.DateField(null=True, blank=True)

class TraderPurchasesContract(TraderContract):
    shipping_method = models.CharField(max_length=1, choices=SHIPPING_METHOD_CHOICES, null=True, blank=True)
    payment_due_date = models.DateField(null=True, blank=True)
    removal_date = models.DateField(null=True, blank=True)
    frame_color = models.CharField(max_length=100, null=True, blank=True)
    receipt = models.CharField(max_length=100, null=True, blank=True)
    transfer_deadline = models.DateField(null=True, blank=True)
    bank_name = models.CharField(max_length=200, null=True, blank=True)
    account_number = models.CharField(max_length=200, null=True, blank=True)
    branch_name = models.CharField(max_length=200, null=True, blank=True)
    account_holder = models.CharField(max_length=200, null=True, blank=True)
    products = GenericRelation(ContractProduct, related_query_name='trader_purchases_contract')
    documents = GenericRelation(ContractDocument, related_query_name='trader_purchases_contract')

    purchase_print_date = models.DateField(null=True, blank=True)

class TraderLink(models.Model):
    purchase_contract = models.ForeignKey(ContractProduct, on_delete=models.CASCADE, related_name="purchase_contract_link")
    sale_contract = models.ForeignKey(ContractProduct, on_delete=models.CASCADE, related_name="sale_contract_link")
    purchase_same_order = models.IntegerField(default=0, null=True)
    sale_same_order = models.IntegerField(default=0, null=True)

class TraderSalesSender(models.Model):
    contract = models.ForeignKey(TraderSalesContract, on_delete=models.CASCADE, related_name='senders')
    type = models.CharField(max_length=1, choices=ITEM_CHOICES)
    sender = models.ForeignKey(Sender, on_delete=models.SET_NULL, null=True)
    expected_arrival_date = models.DateField(null=True, blank=True)
    
class TraderPurchasesSender(models.Model):
    contract = models.ForeignKey(TraderPurchasesContract, on_delete=models.CASCADE, related_name='senders')
    type = models.CharField(max_length=1, choices=ITEM_CHOICES)
    sender = models.ForeignKey(Sender, on_delete=models.SET_NULL, null=True)
    desired_arrival_date = models.DateField(null=True, blank=True)
    shipping_company = models.CharField(max_length=100, null=True, blank=True)
    remarks = models.TextField(null=True, blank=True)

class HallContract(models.Model):
    contract_id = models.CharField(max_length=200)
    created_at = models.DateField()
    customer = models.ForeignKey(Customer, on_delete=models.SET_NULL, null=True)
    hall = models.ForeignKey(Hall, on_delete=models.SET_NULL, null=True)
    remarks = models.TextField(null=True, blank=True)
    fee = models.IntegerField(default=0)
    shipping_date = models.DateField(null=True, blank=True)
    opening_date = models.DateField(null=True, blank=True)
    payment_method = models.CharField(max_length=2, choices=PAYMENT_METHOD_CHOICES)
    transfer_account = models.CharField(max_length=255, null=True, blank=True)
    person_in_charge = models.CharField(max_length=200, null=True, blank=True)
    confirmor = models.CharField(max_length=200, null=True, blank=True)

    available = models.CharField(max_length=1, null=True, blank=True, default='T')

    class Meta:
        abstract = True
    
    @property
    def sub_total(self):
        sum = 0
        for product in self.products.all():
            sum += product.amount
        for document in self.documents.all():
            sum += document.amount
        for document_fee in self.document_fees.all():
            sum += document_fee.amount
        return sum
    
    @property
    def tax(self):
        sum = 0
        for product in self.products.all():
            sum += product.tax
        for document in self.documents.all():
            sum += document.tax
        for document_fee in self.document_fees.all():
            sum += document_fee.tax
        return sum

    @property
    def total(self):
        return self.sub_total + self.tax + self.fee
    
    @property
    def taxed_total(self):
        return self.total - self.fee

class HallSalesContract(HallContract):
    products = GenericRelation(ContractProduct, related_query_name='hall_sales_contract')
    documents = GenericRelation(ContractDocument, related_query_name='hall_sales_contract')
    document_fees = GenericRelation(ContractDocumentFee, related_query_name='hall_sales_contract')
    milestones = GenericRelation(Milestone, related_query_name='hall_sales_contract')

    sale_print_date = models.DateField(null=True, blank=True)

class HallPurchasesContract(HallContract):
    memo = models.TextField(null=True, blank=True)
    products = GenericRelation(ContractProduct, related_query_name='hall_purchases_contract')
    documents = GenericRelation(ContractDocument, related_query_name='hall_purchases_contract')
    document_fees = GenericRelation(ContractDocumentFee, related_query_name='hall_purchases_contract')
    milestones = GenericRelation(Milestone, related_query_name='hall_purchases_contract')

    purchase_print_date = models.DateField(null=True, blank=True)
