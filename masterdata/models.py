from django.db import models
from django.utils.translation import gettext_lazy as _

POSTAL_CODE = '537―0021'
ADDRESS = '大阪府大阪市東成区東中本2丁目4―15'
COMPANY_NAME = 'バッジオ株式会社'
CEO = '金 昇志'
TEL = '06-6753-8078'
FAX = '06-6753-8079'
TRANSFER_ACCOUNT = 'りそな銀行　船場支店（101）　普通　0530713　バッジオカブシキガイシャ'
REFAX = '06-6753-8079'
P_SENSOR_NUMBER = '8240-2413-3628'
INPUT_FORMATS = ['%Y/%m/%d', '%m/%d/%Y']
THRESHOLD_PRICE = 100000

SECURE_PAYMENT = 'あんしん決済'
NO_FEE_SALES = '非課売上'
FEE_SALES = '課税売上10%'
NO_FEE_PURCHASES = '非課仕入'
FEE_PURCHASES = '課対仕入10%'

SHIPPING_METHOD_CHOICES = (
    ('D', _('Delivery')),
    ('R', _('Receipt')),
    ('C', _('ID Change')),
    ('B', _('* Blank')),
)

PAYMENT_METHOD_CHOICES = (
    ('TR', _('Transfer')),
    ('CH', _('Check')),
    ('BL', _('Bill')),
    ('CA', _('Cash')),
)

PRODUCT_TYPE_CHOICES = (
    ('M', _('Main body')),
    ('F', _('Frame')),
    ('C', _('Cell')),
    ('N', _('Nail sheet')),
)

ITEM_CHOICES = (
    ('P', _('Product')),
    ('D', _('Document'))
)

STOCK_CHOICES = (
    ('D', _('Done')),
    ('P', _('Pending'))
)

TYPE_CHOICES = (
    ('P', _('Pachinko')),
    ('S', _('Slot'))
)


class MasterData(models.Model):
    name = models.CharField(max_length=200)
    customer_name = models.CharField(max_length=200, null=True, blank=True)
    customer_frigana = models.CharField(max_length=200, null=True, blank=True)
    frigana = models.CharField(max_length=200)
    postal_code = models.CharField(max_length=20, null=True, blank=True)
    address = models.CharField(max_length=200)
    tel = models.CharField(max_length=100, null=True, blank=True)
    fax = models.CharField(max_length=100, null=True, blank=True)
    
    class Meta:
        abstract = True
    
    def __str__(self):
        return self.name


class Customer(MasterData):
    excel = models.CharField(max_length=200, null=True, blank=True)


class Hall(MasterData):
    payee = models.CharField(max_length=200)


class Sender(MasterData):
    excel = models.CharField(max_length=200, null=True, blank=True)


class Product(models.Model):
    name = models.CharField(max_length=200)
    maker = models.CharField(max_length=200)
    type = models.CharField(max_length=1, choices=TYPE_CHOICES)
    purchase_date = models.DateField(null=True, blank=True)
    supplier = models.CharField(max_length=200, null=True, blank=True)
    person_in_charge = models.CharField(max_length=200, null=True, blank=True)
    quantity = models.IntegerField(null=True, blank=True)
    price = models.IntegerField(null=True, blank=True)
    # stock = models.IntegerField(null=True, blank=True)
    # amount = models.IntegerField(null=True, blank=True)

    @property
    def amount(self):
        if self.price and self.quantity:
            return self.price * self.quantity
        return None

    def __str__(self):
        return self.name


class Document(models.Model):
    name = models.CharField(max_length=200)
    term = models.CharField(max_length=200)
    taxation = models.CharField(max_length=100)

    def __str__(self):
        return self.name
    
    @property
    def taxable(self):
        return self.name != SECURE_PAYMENT


class DocumentFee(models.Model):
    type = models.CharField(max_length=1, choices=TYPE_CHOICES, default='P')
    model_price = models.IntegerField()
    unit_price = models.IntegerField()
    application_fee = models.IntegerField(default=30000)


# class InventoryProduct(models.Model):
#     name = models.CharField(max_length=200)
#     identifier = models.CharField(max_length=20)
#     purchase_date = models.DateField()
#     supplier = models.CharField(max_length=200)
#     person_in_charge = models.CharField(max_length=200)
#     quantity = models.IntegerField()
#     price = models.IntegerField()
#     stock = models.IntegerField()

#     @property
#     def amount(self):
#         if self.price and self.quantity:
#             return self.price * self.quantity
#         return None


class PersonInCharge(models.Model):
    name = models.CharField(max_length=200)
