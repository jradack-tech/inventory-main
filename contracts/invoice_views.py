import xlwt
from django.views.generic.base import View
from django.http import HttpResponse
from django.utils.translation import gettext as _
from users.views import AdminLoginRequiredMixin
from masterdata.models import (
    Customer, Hall, Sender, Product, Document, DocumentFee,
    PRODUCT_TYPE_CHOICES, SHIPPING_METHOD_CHOICES, PAYMENT_METHOD_CHOICES, TYPE_CHOICES,
    P_SENSOR_NUMBER, COMPANY_NAME, ADDRESS, TEL, FAX, POSTAL_CODE, CEO
)
from .forms import (
    TraderSalesContractForm, TraderPurchasesContractForm, HallSalesContractForm, HallPurchasesContractForm,
    TraderSalesProductSenderForm, TraderSalesDocumentSenderForm,
    TraderPurchasesProductSenderForm, TraderPurchasesDocumentSenderForm,
    ProductFormSet, DocumentFormSet, DocumentFeeFormSet, MilestoneFormSet
)
from .utilities import get_shipping_date_label, ordinal, log_export_operation
from datetime import datetime

def getDate(date): 
    result = None
    try: 
        try:
            result = datetime.strptime(date, '%m/%d/%Y')
        except:
            try:
                result = datetime.strptime(date, '%Y/%m/%d')
            except:
                try:
                    result = datetime.strptime(date, '%m-%d-%Y')
                except:
                    result = datetime.strptime(date, '%Y-%m-%d')
    except:
        result = None

    return result

class TraderSalesInvoiceViewOnly(AdminLoginRequiredMixin, View):
    def post(self, *args, **kwargs):
        top_padding_height = int(256 * 0.4)
        cell_width_list = [ 256 *  10, 256 * 12, int(256 * 9.5), int(256 * 2.0), 
                            int(256 * 4.5), int(256 * 7.5), 256 * 10, int(256 * 6.5),
                            int(256 * 13.5), int(256 * 8.5), int(256 * 7.5)]
        cell_height = int(20 * 15)

        header_height = int(20 * 21)
        space_height = int(20 * 5.25)
        company_height = int(20 * 23.25)
        height_18 = int(20 * 18)
        height_16 = int(20 * 16)
        height_16_5 = int(20 * 16.5)
        height_15 = 20 * 15
        height_13_5 = int(20 * 13.5)
        height_12 = 20 * 12

        border_bottom = xlwt.easyxf('borders: bottom_color black, bottom thin;')
        border_left = xlwt.easyxf('borders: left_color black, left thin;')
        border_right = xlwt.easyxf('borders: right_color black, right thin;')
        border_top = xlwt.easyxf('borders: top_color black, top thin;')
        bold_medium_top = xlwt.easyxf('borders: top_color black, top medium;')

        font_11_left = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: vert center, horiz left')
        font_11_center_with_border = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: vert center, horiz center;\
                                                borders: top_color black, left_color black, right_color black, bottom_color black,\
                                                top thin, bottom thin, left thin, right thin;')
        font_11_left_with_border = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: vert center, horiz left;\
                                                borders: top_color black, left_color black, right_color black, bottom_color black,\
                                                top thin, bottom thin, left thin, right thin;')
        font_11_center_with_right_border = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: vert center, horiz center;\
                                                borders: top_color black, left_color black, right_color black, bottom_color black,\
                                                top thin, bottom thin, left thin, right medium;')
        font_11_center_with_left_side_border = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: vert center, horiz center;\
                                                borders: top_color black, left_color black, right_color black, bottom_color black,\
                                                top medium, bottom medium, left medium, right thin;')
        font_11_center_with_top_border = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: vert center, horiz center;\
                                                borders: top_color black, left_color black, right_color black, bottom_color black,\
                                                top medium, bottom thin, left thin, right thin;') 
        font_11_center_with_top_right_border = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: vert center, horiz center;\
                                                borders: top_color black, left_color black, right_color black, bottom_color black,\
                                                top medium, bottom thin, left thin, right medium;', num_format_str='"¥"#,###') 
        font_11_center_with_right_border = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: vert center, horiz center;\
                                                borders: top_color black, left_color black, right_color black, bottom_color black,\
                                                top thin, bottom thin, left thin, right medium;')   
        title_style = xlwt.easyxf('font: height 360, name ＭＳ Ｐゴシック, color black;\
                                    align: vert center, horiz center, wrap on;')
        company_title_style = xlwt.easyxf('font: bold on, height 320, name ＭＳ Ｐゴシック; align: vert center, horiz center, wrap off, shrink on;\
                                            borders: bottom_color black, bottom thin;')
        company_label_style = xlwt.easyxf('font: height 280, name ＭＳ Ｐゴシック; align: vert center, horiz center, wrap on;\
                                            borders: bottom_color black, bottom thin;')
        company_label_style1 = xlwt.easyxf('font: height 280, name ＭＳ Ｐゴシック; align: vert center, horiz left, wrap on;')
        product_table_first_th_style = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: vert center, horiz center, wrap on;\
                                            borders: top_color black, left_color black, right_color black, top medium, left medium, right thin;')
        product_table_th_style = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: vert center, horiz center, wrap on;\
                                            borders: top_color black, right_color black, top medium, right thin;')
        product_table_last_th_style = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: vert center, horiz center, wrap on;\
                                            borders: top_color black, right_color black, top medium, right medium;')
        product_first_content_style = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: vert center, horiz left, wrap no;\
                                            borders: top_color black, left_color black, right_color black, bottom_color black, bottom thin, top thin, left medium, right thin;')
        product_first_content_style_thin = xlwt.easyxf('font: height 180, name ＭＳ Ｐゴシック; align: vert center, horiz left, wrap no;\
                                            borders: top_color black, left_color black, right_color black, bottom_color black, bottom thin, top thin, left medium, right thin;')
        primary_text_style = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: vert center, horiz left, wrap on;')
        product_content_style = xlwt.easyxf('font: height 220, name ＭＳ ゴシック; align: vert center, horiz right, wrap on;\
                                            borders: top_color black, right_color black, bottom_color black, bottom thin, top thin, right thin;',
                                            num_format_str='#,##0')
        product_last_content_style = xlwt.easyxf('font: height 220, name ＭＳ ゴシック; align: vert center, horiz right, wrap on;\
                                            borders: top_color black, right_color black, bottom_color black, bottom thin, top thin, right medium;',
                                            num_format_str='#,##0')                               
        subtotal_label_style = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: vert center, horiz center, wrap on;\
                                            borders: top_color black, left_color black, top medium, left medium;')
        subtotal_value_style = xlwt.easyxf('font: height 220, name ＭＳ ゴシック; align: vert center, horiz right, wrap on;\
                                            borders: top_color black, left_color black, right_color black, top medium, left medium, right medium;',
                                            num_format_str='#,##0')
        fee_label_style = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: vert center, horiz center, wrap on;\
                                            borders: top_color black, left_color black, top thin, left medium;')
        fee_value_style = xlwt.easyxf('font: height 220, name ＭＳ ゴシック; align: vert center, horiz right, wrap on;\
                                            borders: top_color black, left_color black, right_color black, top thin, left medium, right medium;',
                                            num_format_str='#,##0')
        total_label_style = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: vert center, horiz center, wrap on;\
                                            borders: top_color black, left_color black, bottom_color black,\
                                            bottom medium, top thin, left medium;')
        total_value_style = xlwt.easyxf('font: height 220, bold on, name ＭＳ ゴシック; align: vert center, horiz right, wrap on;\
                                            borders: top_color black, left_color black, right_color black, bottom_color black,\
                                            bottom medium, top thin, left medium, right medium;',
                                            num_format_str='"¥"#,###')
        
        transfer_account_style = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: vert center, horiz center;\
                                                borders: left_color black, top_color black, top thin, left thin;')

        TTL_CENTER = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz center, vert center, wrap on; borders: top_color black, left_color black, bottom_color black, right_color black, top medium, left medium, right thin, bottom thin;')
        TT_CENTER = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz center, vert center, wrap on; borders: top_color black, bottom_color black, right_color black, top medium, right thin, bottom thin;')
        TTR_CENTER = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz center, vert center, wrap on; borders: top_color black, left_color black, bottom_color black, right_color black, top medium, left thin, right medium, bottom thin;')
        TC_RIGHT = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz right, wrap on; borders: bottom_color black, right_color black, right thin, bottom thin;', num_format_str='#,###')
        TR_RIGHT = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz right, wrap on; borders: bottom_color black, right_color black, right medium, bottom thin;', num_format_str='#,###')
        F11 = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: wrap on;')
        LINE_TOP = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: wrap on; borders: top_color black, top medium;')
        TTR_RIGHT = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz right, wrap on; borders: top_color black, left_color black, bottom_color black, right_color black, top medium, left thin, right medium, bottom thin;', num_format_str='#,###')
        TL_CENTER = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz center, vert center, wrap on; borders: bottom_color black, left_color black, right_color black, right thin, bottom thin, left medium;')
        TBL_CENTER = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz center, vert center, wrap on; borders: left_color black, bottom_color black, right_color black, left medium, right thin, bottom medium;')
        TBR_RIGHT_TOTAL = xlwt.easyxf('font: height 240, name ＭＳ Ｐゴシック, bold on; align: horiz right, vert center, wrap on; borders: bottom_color black, right_color black, right medium, bottom medium;', num_format_str='"¥"#,###')
        TL = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: wrap on; borders: bottom_color black, left_color black, right_color black, right thin, bottom thin, left medium;')

        user_id = self.request.user.id
        log_export_operation(user_id, "{} - {}".format(_("Sales contract"), _("Trader sales")))
        contract_form = TraderSalesContractForm(self.request.POST)
        contract_id = contract_form.data.get('contract_id', '')
        created_at = contract_form.data.get('created_at', '')
        updated_at = contract_form.data.get('updated_at', '')
        person_in_charge = contract_form.data.get('person_in_charge', '')
        customer_id = contract_form.data.get('customer_id')
        manager = contract_form.data.get('manager')
        company = frigana = postal_code = address = tel = fax = ""
        
        if customer_id:
            customer = Customer.objects.get(id=customer_id)
            company = customer.name
            frigana = customer.frigana
            postal_code = customer.postal_code
            address = customer.address
            tel = customer.tel
            fax = customer.fax

        response = HttpResponse(content_type='application/ms-excel')
        response['Content-Disposition'] = 'attachment; filename="trader_sales_contract_{}.xls"'.format(contract_id)
        # response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        # response['Content-Disposition'] = 'attachment; filename="trader_sales_contract_{}.xlsx"'.format(contract_id)
        wb = xlwt.Workbook(encoding='utf-8')
        ws = wb.add_sheet('業者売用請求書原本', cell_overwrite_ok=True)

        ws.set_header_str(str.encode(''))
        ws.set_footer_str(str.encode(''))
        ws.set_left_margin(0.314961)
        ws.set_top_margin(0.393701)
        ws.set_right_margin(0.314961)
        ws.set_bottom_margin(0.19685)

        

        for i in range(11):
            ws.col(i).width = cell_width_list[i]

        for i in range(1, 80):
            ws.row(i).height_mismatch = True
            ws.row(i).height = cell_height

        
        product_formset = ProductFormSet(
            self.request.POST,
            prefix='product'
        )
        num_of_products = product_formset.total_form_count()

        document_formset = DocumentFormSet(
            self.request.POST,
            prefix='document'
        )
        num_of_documents = document_formset.total_form_count()

        total_number = num_of_products + num_of_documents

        shipping_method = contract_form.data.get('shipping_method')
        shipping_date = contract_form.data.get('shipping_date')
        payment_method = contract_form.data.get('payment_method')
        payment_due_date = contract_form.data.get('payment_due_date')
        sub_total = contract_form.data.get('sub_total')
        tax = contract_form.data.get('tax')
        fee = contract_form.data.get('fee')
        total = contract_form.data.get('total')
        remarks = contract_form.data.get('remarks')
        shipping_date_label, shipping_sender = get_shipping_date_label(shipping_method)

        try:
            sub_total = int(sub_total.replace(',', '')) if len(sub_total.replace(',', '')) > 0 else 0
            tax = int(tax.replace(',', '')) if len(tax.replace(',', '')) > 0 else 0
            fee = int(fee.replace(',', '')) if len(fee.replace(',', '')) > 0 else 0
            total = int(total.replace(',', '')) if len(total.replace(',', '')) > 0 else 0
        except ValueError:
            sub_total = 0
            tax = 0
            fee = 0
            total = 0

        #### Start drawing ###############################################################

        ws.row(0).height_mismatch = True
        ws.row(0).height = top_padding_height

        ws.row(1).height = header_height

        ws.write_merge(1, 1, 1, 8, '請  求  書', title_style)
        # ws.write_merge(1, 1, 9, 10, created_at, contract_date_style)

        ws.row(2).height = height_16
        ws.row(3).height = height_16

        ws.write_merge(2, 2, 9, 10, '契約日', font_11_center_with_border)
        ws.write_merge(3, 3, 9, 10, created_at, font_11_center_with_border)

        ws.row(4).height = company_height
        ws.write_merge(4, 4, 0, 2, company, company_title_style)
        ws.write_merge(4, 4, 3, 4, _('CO.'), company_label_style)

        ws.row(5).height = height_18

        # ws.write_merge(6, 6, 1, 2, manager, company_supplier)
        # ws.write(6, 3, '', border_bottom)
        # ws.write(6, 4, _('Mr.'), company_supplier_label)
        ws.write(5, 0, '住所：', primary_text_style)
        ws.write_merge(5, 5, 1, 8, address, primary_text_style)

        ws.row(6).height = space_height
        ws.row(7).height = height_16_5
        # ws.write(6, 0, 'TEL:', primary_text_style)
        ws.write_merge(7, 7, 0, 1, 'TEL: {}'.format(tel), font_11_left)
        ws.write_merge(7, 7, 2, 5, 'FAX: {}'.format(fax), font_11_left)

        ws.row(8).height = height_16_5
        ws.row(9).height = height_16_5
        ws.write_merge(8, 9, 5, 10, "バッジオ株式会社", company_label_style1)

        ws.row(10).height = height_16_5
        ws.write_merge(10, 10, 5, 10, "〒537-0021　大阪府大阪市東成区東中本2丁目4-15", font_11_left)

        ws.row(11).height = height_16_5
        ws.write_merge(11, 11, 5, 10, "TEL 06-6753-8078  FAX 06-6753-8079", font_11_left)

        ws.row(12).height = height_16_5
        ws.write(12, 5, '担当：', font_11_left)
        ws.write_merge(12, 12, 6, 10, person_in_charge, font_11_left)

        ws.row(13).height = space_height

        ws.row(14).height = height_15
        ws.write_merge(14, 14, 0, 6, '商　品　名', product_table_first_th_style)
        ws.write(14, 7, '数量', product_table_th_style)
        ws.write(14, 8, '単　価', product_table_th_style)
        ws.write_merge(14, 14, 9, 10, '金　額', product_table_last_th_style)

        # Product Table
        row_no = 15
        if num_of_products:
            for form in product_formset.forms:
                form.is_valid()

                # if prev_counter == 6:
                #     break
                # prev_counter += 1

                id = form.cleaned_data.get('product_id')
                product_name = Product.objects.get(id=id).name
                quantity = form.cleaned_data.get('quantity', 0)
                price = form.cleaned_data.get('price', 0)
                amount = quantity * price

                if total_number > 6: 
                    continue

                ws.row(row_no).height = height_18
                if len(product_name) > 23:
                    ws.write_merge(row_no, row_no, 0, 6, product_name, product_first_content_style_thin)
                else:
                    ws.write_merge(row_no, row_no, 0, 6, product_name, product_first_content_style)
                ws.write(row_no, 7, quantity, product_content_style)
                ws.write(row_no, 8, price, product_content_style)
                ws.write_merge(row_no, row_no, 9, 10, amount, product_last_content_style)
                row_no += 1

        # Document Table
        if num_of_documents:
            for form in document_formset.forms:
                form.is_valid()

                id = form.cleaned_data.get('document_id')
                document_name = Document.objects.get(id=id).name
                document_name = document_name.replace('（売上）', '').replace('（仕入）', '')
                quantity = form.cleaned_data.get('quantity', 0)
                price = form.cleaned_data.get('price', 0)
                amount = quantity * price

                if total_number > 6: 
                    continue

                ws.row(row_no).height = height_18
                ws.write_merge(row_no, row_no, 0, 6, document_name, product_first_content_style)
                ws.write(row_no, 7, quantity, product_content_style)
                ws.write(row_no, 8, price, product_content_style)
                ws.write_merge(row_no, row_no, 9, 10, amount, product_last_content_style)
                row_no += 1
        
        # if total_number < 6:
        forRange = 6 - total_number if total_number <= 6 else 6
        for i in range(0, forRange):
            ws.row(row_no).height = height_18
            ws.write_merge(row_no, row_no, 0, 6, None, product_first_content_style)
            ws.write(row_no, 7, None, product_content_style)
            ws.write(row_no, 8, None, product_content_style)
            ws.write_merge(row_no, row_no, 9, 10, None, product_last_content_style)
            row_no += 1
        
        ws.row(row_no).height = height_15
        # ws.write_merge(row_no, row_no, 0, 1, '{}: '.format(shipping_date_label), shipping_date_label_style)
        # ws.write_merge(row_no, row_no, 2, 6, shipping_date, shipping_date_style)
        ws.write_merge(row_no, row_no, 0, 6, '', bold_medium_top)
        ws.write_merge(row_no, row_no, 7, 8, '小　計', subtotal_label_style)
        ws.write_merge(row_no, row_no, 9, 10, sub_total, subtotal_value_style)
        row_no += 1

        ws.row(row_no).height = height_15
        # ws.write(row_no, 0, '※備考', remark_lebel_style)
        ws.write_merge(row_no, row_no, 7, 8, '消費税（10%）', fee_label_style)
        ws.write_merge(row_no, row_no, 9, 10, tax, fee_value_style)
        row_no += 1

        ws.row(row_no).height = height_15
        # ws.write_merge(row_no, row_no + 1, 0, 6, remarks, remark_content_style)
        ws.write_merge(row_no, row_no, 7, 8, '保険代（非課税）', fee_label_style)
        ws.write_merge(row_no, row_no, 9, 10, fee, fee_value_style)
        row_no += 1

        ws.row(row_no).height = height_15
        
        ws.write_merge(row_no, row_no, 7, 8, '合　計', total_label_style)
        ws.write_merge(row_no, row_no, 9, 10, total, total_value_style)
        row_no += 1

        ws.row(row_no).height = space_height
        row_no += 1

        ws.write(row_no, 4, '下記の通り御請求申し上げます。', font_11_left)
        row_no += 1

        ws.write_merge(row_no, row_no, 0, 2, '発送日', font_11_center_with_border)
        ws.write_merge(row_no, row_no + 4, 4, 5, 'お支払内訳', font_11_center_with_left_side_border)

        for i in range(5):
            if i == 0:
                ws.write(row_no + i, 6, '初回', font_11_center_with_top_border)
                ws.write_merge(row_no, row_no, 7, 8, shipping_date, font_11_center_with_top_border)
                ws.write_merge(row_no, row_no, 9, 10, total, font_11_center_with_top_right_border)
            else:
                ws.write(row_no + i, 6, str(i+1) + '回', font_11_center_with_border)
                ws.write_merge(row_no + i, row_no + i, 7, 8, '', font_11_center_with_border)
                ws.write_merge(row_no + i, row_no + i, 9, 10, '', font_11_center_with_right_border)

        row_no += 1
        ws.write_merge(row_no, row_no, 0, 2, shipping_date, font_11_center_with_border)
        row_no += 2

        ws.write_merge(row_no, row_no, 0, 2, 'お支払方法', font_11_center_with_border)
        ws.write_merge(row_no + 1, row_no + 1, 0, 2, '振込', font_11_center_with_border)

        # ws.write_merge(row_no, row_no + 1, 5, 7, 'お支払期限', billed_deadline_label_style)
        # ws.write_merge(row_no, row_no + 1, 8, 10, payment_due_date, billed_deadline_value_style)
        row_no += 1
        ws.row(row_no).height = height_16

        row_no += 1
        ws.row(row_no).height = height_16
        ws.write_merge(row_no, row_no, 6, 10, '', bold_medium_top)

        row_no += 1
        ws.row(row_no).height = height_16
        
        ws.write(row_no, 0, '※ お振込み手数料は貴社ご負担でお願い致します。', font_11_left)
        row_no += 1

        ws.write_merge(row_no, row_no, 0, 1, '振込先口座', transfer_account_style)
        ws.write_merge(row_no, row_no, 2, 10, 'りそな銀行　船場支店（101）　普通　0530713　バッジオカブシキガイシャ', font_11_left_with_border)
        row_no += 1

        ws.write(row_no, 0, '', border_left)
        ws.write_merge(row_no, row_no, 2, 10, '', font_11_left_with_border)
        row_no += 1

        ws.write(row_no, 0, '', border_left)
        ws.write_merge(row_no, row_no, 2, 10, '', font_11_left_with_border)
        row_no += 1

        ws.row(row_no).height = space_height
        for i in range(0, 11):
            ws.write(row_no, i, '', border_top)
        row_no += 1

        ws.write_merge(row_no, row_no, 0, 4, 'P-SENSOR {} {}'.format(_('Member ID'), P_SENSOR_NUMBER), font_11_left)
        row_no += 1
        ws.row(row_no).height = space_height
        row_no += 1
        ws.row(row_no).height = height_16
        row_no += 1
        ws.write(row_no, 0, 'No.{}'.format(contract_id), font_11_left)
            
        # -------------------------------------------------------------------------------------------------------------------------
        if total_number > 6:
            ws = wb.add_sheet('別紙', cell_overwrite_ok=True)

            ws.set_header_str(str.encode(''))
            ws.set_footer_str(str.encode(''))
            ws.set_left_margin(0.314961)
            ws.set_top_margin(0.393701)
            ws.set_right_margin(0.314961)
            ws.set_bottom_margin(0.19685)

            w_arr = [9.15, 8.04, 10.04, 2.04, 5.04, 8.04, 6.71, 7.15, 12.59, 6.71, 6.71]
            for i in range(11): 
                ws.col(i).width = int(300 * w_arr[i])
            for i in range(56):
                ws.row(i).height_mismatch = True
                ws.row(i).height = int(20 * 18)
            ws.row(44).height = int(20 * 3.75)
            ws.row(45).height = int(20 * 15.75)

            ####################################### Table #####################################
            ws.write_merge(0, 0, 0, 6, '商　品　名', TTL_CENTER)
            ws.write(0, 7, '数量', TT_CENTER)
            ws.write(0, 8, '単　価', TT_CENTER)
            ws.write_merge(0, 0, 9, 10, '金　額', TTR_CENTER)

            row_no = 1

            total_quantity = total_price = 0
            ##### Products
            if num_of_products:
                for form in product_formset.forms:
                    form.is_valid()

                    # if next_counter < 6:
                    #     next_counter += 1
                    #     continue

                    id = form.cleaned_data.get('product_id')
                    product_name = Product.objects.get(id=id).name
                    quantity = form.cleaned_data.get('quantity', 0)
                    price = form.cleaned_data.get('price', 0)
                    amount = quantity * price

                    total_quantity += quantity
                    total_price += price

                    ws.write_merge(row_no, row_no, 0, 6, product_name, TL)
                    ws.write(row_no, 7, quantity, TC_RIGHT)
                    ws.write(row_no, 8, price, TC_RIGHT)
                    ws.write_merge(row_no, row_no, 9, 10, amount, TR_RIGHT)

                    row_no += 1

            ###### Documents
            if num_of_documents:
                for form in document_formset.forms:
                    form.is_valid()

                    # if next_counter < 6:
                    #     next_counter += 1
                    #     continue

                    id = form.cleaned_data.get('document_id')
                    document_name = Document.objects.get(id=id).name.replace('（売上）', '').replace('（仕入）', '')
                    quantity = form.cleaned_data.get('quantity', 0)
                    price = form.cleaned_data.get('price', 0)
                    amount = quantity * price

                    total_quantity += quantity
                    total_price += price

                    ws.write_merge(row_no, row_no, 0, 6, document_name, TL)
                    ws.write(row_no, 7, quantity, TC_RIGHT)
                    ws.write(row_no, 8, price, TC_RIGHT)
                    ws.write_merge(row_no, row_no, 9, 10, amount, TR_RIGHT)

                    row_no += 1

            if total_number - 1 < 39:
                for i in range(0, 39 - total_number + 1):

                    # if prev_counter <= 6:
                    #     prev_counter += 1
                    #     continue

                    ws.write_merge(row_no, row_no, 0, 6, None, TL)
                    ws.write(row_no, 7, None, TC_RIGHT)
                    ws.write(row_no, 8, None, TC_RIGHT)
                    ws.write_merge(row_no, row_no, 9, 10, None, TR_RIGHT)

                    row_no += 1

            ws.write_merge(40, 40, 0, 6, '※備考', LINE_TOP)
            ws.write_merge(40, 40, 7, 8, '小　計', TTL_CENTER)
            ws.write_merge(40, 40, 9, 10, sub_total, TTR_RIGHT)
            ws.write_merge(41, 41, 7, 8, '消費税（10%）', TL_CENTER)        
            ws.write_merge(41, 41, 9, 10, tax, TR_RIGHT)
            ws.write_merge(42, 42, 7, 8, '保険代（非課税）', TL_CENTER)
            ws.write_merge(42, 42, 9, 10, fee, TR_RIGHT)
            ws.write_merge(43, 43, 7, 8, '合　計', TBL_CENTER)
            ws.write_merge(43, 43, 9, 10, total, TBR_RIGHT_TOTAL)

            ws.write_merge(45, 45, 0, 3, 'No.{}'.format(contract_id), F11)
        wb.save(response)
        return response

class TraderSalesInvoiceView(AdminLoginRequiredMixin, View):
    def post(self, *args, **kwargs):
        top_padding_height = int(256 * 0.4)
        cell_width_list = [ 256 *  10, 256 * 9, int(256 * 10.5), int(256 * 2.0), 
                            int(256 * 4.5), int(256 * 7.5), 256 * 10, int(256 * 6.5),
                            int(256 * 13.5), int(256 * 8.5), int(256 * 7.5)]
        cell_height = int(20 * 15)

        header_height = int(20 * 20.1)
        space_height = int(20 * 4.35)
        company_height = int(20 * 19.25)
        height_18 = int(20 * 16)
        height_16 = int(20 * 14)
        height_16_1 = int(20 * 15)
        height_16_5 = int(20 * 14.5)
        height_15 = 20 * 13
        height_13_5 = int(20 * 12.5)
        height_12 = 20 * 10

        border_bottom = xlwt.easyxf('borders: bottom_color black, bottom thin;')
        border_left = xlwt.easyxf('borders: left_color black, left thin;')
        border_right = xlwt.easyxf('borders: right_color black, right thin;')
        border_top = xlwt.easyxf('borders: top_color black, top thin;')
        bold_medium_top = xlwt.easyxf('borders: top_color black, top medium;')

        font_11_left = xlwt.easyxf('font: height 190, name ＭＳ Ｐゴシック; align: vert center, horiz left')
        font_12_left = xlwt.easyxf('font: height 240, name ＭＳ Ｐゴシック; align: vert center, horiz left')
        font_11_left_r = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: vert center, horiz left')
        font_10_left = xlwt.easyxf('font: height 200, name ＭＳ Ｐゴシック; align: vert center, horiz left')
        font_14_left = xlwt.easyxf('font: height 260, name ＭＳ Ｐゴシック; align: vert center, horiz left')
        font_14_left1 = xlwt.easyxf('font: height 280, name ＭＳ Ｐゴシック; align: vert center, horiz left')
        font_11_center_with_border = xlwt.easyxf('font: height 190, name ＭＳ Ｐゴシック; align: vert center, horiz center;\
                                                borders: top_color black, left_color black, right_color black, bottom_color black,\
                                                top thin, bottom thin, left thin, right thin;')
        font_11_left_with_border = xlwt.easyxf('font: height 190, name ＭＳ Ｐゴシック; align: vert center, horiz left;\
                                                borders: top_color black, left_color black, right_color black, bottom_color black,\
                                                top thin, bottom thin, left thin, right thin;')
        title_style = xlwt.easyxf('font: height 360, name ＭＳ Ｐゴシック, color black;\
                                    align: vert center, horiz center, wrap on;')
        contract_date_style = xlwt.easyxf('font: height 190, name ＭＳ Ｐゴシック; align: vert bottom, horiz center, wrap on;')
        company_title_style = xlwt.easyxf('font: bold on, height 320, name ＭＳ Ｐゴシック; align: vert center, horiz center, wrap off, shrink on;\
                                            borders: bottom_color black, bottom thin;')
        company_label_style = xlwt.easyxf('font: height 260, name ＭＳ Ｐゴシック; align: vert center, horiz center, wrap on;\
                                            borders: bottom_color black, bottom thin;')
        in_charge_style = xlwt.easyxf('font: height 190, name ＭＳ Ｐゴシック; align: vert center, horiz left, wrap off;\
                                        borders: bottom_color black, bottom thin;')
        company_supplier = xlwt.easyxf('font: height 180, name ＭＳ Ｐゴシック; align: vert center, horiz center, wrap on;\
                                        borders: bottom_color black, bottom thin;')
        company_supplier_label = xlwt.easyxf('font: height 180, name ＭＳ Ｐゴシック; align: vert center, horiz right, wrap on;\
                                            borders: bottom_color black, bottom thin;')
        product_table_first_th_style = xlwt.easyxf('font: height 190, name ＭＳ Ｐゴシック; align: vert center, horiz center, wrap on;\
                                            borders: top_color black, left_color black, right_color black, top medium, left medium, right thin;')
        product_table_th_style = xlwt.easyxf('font: height 190, name ＭＳ Ｐゴシック; align: vert center, horiz center, wrap on;\
                                            borders: top_color black, right_color black, top medium, right thin;')
        product_table_last_th_style = xlwt.easyxf('font: height 190, name ＭＳ Ｐゴシック; align: vert center, horiz center, wrap on;\
                                            borders: top_color black, right_color black, top medium, right medium;')
        product_first_content_style = xlwt.easyxf('font: height 240, name ＭＳ Ｐゴシック; align: vert center, horiz left, wrap no;\
                                            borders: top_color black, left_color black, right_color black, bottom_color black, bottom thin, top thin, left medium, right thin;')
        product_first_content_style_thin = xlwt.easyxf('font: height 180, name ＭＳ Ｐゴシック; align: vert center, horiz left, wrap no;\
                                            borders: top_color black, left_color black, right_color black, bottom_color black, bottom thin, top thin, left medium, right thin;')
        product_content_style = xlwt.easyxf('font: height 240, name ＭＳ ゴシック; align: vert center, horiz right, wrap on;\
                                            borders: top_color black, right_color black, bottom_color black, bottom thin, top thin, right thin;',
                                            num_format_str='#,##0')
        product_last_content_style = xlwt.easyxf('font: height 240, name ＭＳ Ｐゴシック; align: vert center, horiz right, wrap on;\
                                            borders: top_color black, right_color black, bottom_color black, bottom thin, top thin, right medium;',
                                            num_format_str='#,##0') 
        shipping_date_style = xlwt.easyxf('font: height 240, name ＭＳ Ｐゴシック; align: vert center, horiz left, wrap on;\
                                            borders: top_color black, top medium;')
        shipping_date_label_style = xlwt.easyxf('font: height 190, name ＭＳ Ｐゴシック; align: vert center, horiz left, wrap on;\
                                            borders: top_color black, left_color black, left thin, top medium;')
        remark_lebel_style = xlwt.easyxf('font: height 190, name ＭＳ Ｐゴシック; align: vert center, horiz left;\
                                        borders: left_color black, left thin;')
        remark_content_style = xlwt.easyxf('font: height 190, name ＭＳ Ｐゴシック; align: vert center, horiz left;\
                                        borders: left_color black, bottom_color black, bottom thin, left thin;')
        subtotal_label_style = xlwt.easyxf('font: height 190, name ＭＳ Ｐゴシック; align: vert center, horiz center, wrap on;\
                                            borders: top_color black, left_color black, top medium, left medium;')
        subtotal_value_style = xlwt.easyxf('font: height 240, name ＭＳ ゴシック; align: vert center, horiz right, wrap on;\
                                            borders: top_color black, left_color black, right_color black, top medium, left medium, right medium;',
                                            num_format_str='#,##0')
        fee_label_style = xlwt.easyxf('font: height 190, name ＭＳ Ｐゴシック; align: vert center, horiz center, wrap on;\
                                            borders: top_color black, left_color black, top thin, left medium;')
        fee_value_style = xlwt.easyxf('font: height 240, name ＭＳ ゴシック; align: vert center, horiz right, wrap on;\
                                            borders: top_color black, left_color black, right_color black, top thin, left medium, right medium;',
                                            num_format_str='#,##0')
        total_label_style = xlwt.easyxf('font: height 190, name ＭＳ Ｐゴシック; align: vert center, horiz center, wrap on;\
                                            borders: top_color black, left_color black, bottom_color black,\
                                            bottom medium, top thin, left medium;')
        total_value_style = xlwt.easyxf('font: height 240, bold on, name ＭＳ ゴシック; align: vert center, horiz right, wrap on;\
                                            borders: top_color black, left_color black, right_color black, bottom_color black,\
                                            bottom medium, top thin, left medium, right medium;',
                                            num_format_str='"¥"#,###')
        remark_text_style = xlwt.easyxf('font: height 180, name ＭＳ Ｐゴシック; align: vert center, horiz left;')
        sender_table_title_style = xlwt.easyxf('font: bold on, height 220, name ＭＳ Ｐゴシック; align: vert center, horiz center;\
                                                borders: bottom_color black, bottom thin;')
        sender_table_field_style = xlwt.easyxf('font: height 180, name ＭＳ Ｐゴシック; align: vert center, horiz left;\
                                                borders: left_color black, left thin;')
        sender_address_style = xlwt.easyxf('font: height 190, name ＭＳ Ｐゴシック; align: vert top, horiz left; \
                                            borders: right_color black, right thin; alignment: wrap 1;')
        sender_table_text_style = xlwt.easyxf('font: height 190, name ＭＳ Ｐゴシック; align: vert center, horiz left;\
                                                borders: right_color black, right thin;')
        seal_address_tel_style = xlwt.easyxf('font: height 190, name ＭＳ Ｐゴシック; alignment: wrap True; align: vert center, horiz left;\
                                                borders: right_color black, left_color black, left thin, right thin;')
        seal_address_tel_style1 = xlwt.easyxf('font: height 190, name ＭＳ Ｐゴシック; alignment: wrap True; align: vert center, horiz left;\
                                                borders: right_color black, left_color black, left thin, right thin;')
        seal_company_style = xlwt.easyxf('font: height 320, name ＭＳ Ｐゴシック; align: vert center, horiz left;\
                                                borders: right_color black, left_color black, left thin, right thin;')
        seal_supplier_style = xlwt.easyxf('font: height 190, name ＭＳ Ｐゴシック; align: vert center, horiz left;\
                                                borders: left_color black, left thin;')
        seal_mark_style = xlwt.easyxf('font: height 190, name ＭＳ Ｐゴシック; align: vert center, horiz left;\
                                                borders: right_color black, right thin;')
        billed_deadline_label_style = xlwt.easyxf('font: height 320, bold on, name ＭＳ Ｐゴシック; align: vert center, horiz center;\
                                                borders: right_color black, top_color black, bottom_color black, left_color black,\
                                                right thin, top medium, left medium, bottom medium;')
        billing_amount_value_style = xlwt.easyxf('font: height 320, bold on, name ＭＳ Ｐゴシック; align: vert center, horiz center;\
                                                borders: right_color black, top_color black, bottom_color black, left_color black,\
                                                right medium, top medium, left thin, bottom medium;',
                                                num_format_str='"¥"#,###')
        billed_deadline_value_style = xlwt.easyxf('font: height 320, bold on, name ＭＳ Ｐゴシック; align: vert center, horiz center;\
                                                borders: right_color black, top_color black, bottom_color black, left_color black,\
                                                right medium, top medium, left thin, bottom medium;')
        transfer_account_style = xlwt.easyxf('font: height 190, name ＭＳ Ｐゴシック; align: vert center, horiz center;\
                                                borders: left_color black, top_color black, top thin, left thin;')
        TTL_CENTER = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz center, vert center, wrap on; borders: top_color black, left_color black, bottom_color black, right_color black, top medium, left medium, right thin, bottom thin;')
        TT_CENTER = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz center, vert center, wrap on; borders: top_color black, bottom_color black, right_color black, top medium, right thin, bottom thin;')
        TTR_CENTER = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz center, vert center, wrap on; borders: top_color black, left_color black, bottom_color black, right_color black, top medium, left thin, right medium, bottom thin;')
        TC_RIGHT = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz right, wrap on; borders: bottom_color black, right_color black, right thin, bottom thin;', num_format_str='#,###')
        TR_RIGHT = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz right, wrap on; borders: bottom_color black, right_color black, right medium, bottom thin;', num_format_str='#,###')
        F11 = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: wrap on;')
        LINE_TOP = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: wrap on; borders: top_color black, top medium;')
        TTR_RIGHT = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz right, wrap on; borders: top_color black, left_color black, bottom_color black, right_color black, top medium, left thin, right medium, bottom thin;', num_format_str='#,###')
        TL_CENTER = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz center, vert center, wrap on; borders: bottom_color black, left_color black, right_color black, right thin, bottom thin, left medium;')
        TBL_CENTER = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz center, vert center, wrap on; borders: left_color black, bottom_color black, right_color black, left medium, right thin, bottom medium;')
        TBR_RIGHT_TOTAL = xlwt.easyxf('font: height 240, name ＭＳ Ｐゴシック, bold on; align: horiz right, vert center, wrap on; borders: bottom_color black, right_color black, right medium, bottom medium;', num_format_str='"¥"#,###')
        TL = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: wrap on; borders: bottom_color black, left_color black, right_color black, right thin, bottom thin, left medium;')

        user_id = self.request.user.id
        log_export_operation(user_id, "{} - {}".format(_("Sales contract"), _("Trader sales")))
        contract_form = TraderSalesContractForm(self.request.POST)
        contract_id = contract_form.data.get('contract_id', '')
        created_at = contract_form.data.get('created_at', '')
        updated_at = contract_form.data.get('updated_at', '')
        person_in_charge = contract_form.data.get('person_in_charge', '')
        customer_id = contract_form.data.get('customer_id')
        manager = contract_form.data.get('manager')
        company = frigana = postal_code = address = tel = fax = ""
        if customer_id:
            customer = Customer.objects.get(id=customer_id)
            company = customer.name
            frigana = customer.frigana
            postal_code = customer.postal_code
            address = customer.address
            tel = customer.tel
            fax = customer.fax

        product_formset = ProductFormSet(
            self.request.POST,
            prefix='product'
        )
        num_of_products = product_formset.total_form_count()

        document_formset = DocumentFormSet(
            self.request.POST,
            prefix='document'
        )
        num_of_documents = document_formset.total_form_count()

        total_number = num_of_products + num_of_documents

        shipping_method = contract_form.data.get('shipping_method')
        shipping_date = contract_form.data.get('shipping_date')
        payment_method = contract_form.data.get('payment_method')
        payment_due_date = contract_form.data.get('payment_due_date')
        sub_total = contract_form.data.get('sub_total')
        tax = contract_form.data.get('tax')
        fee = contract_form.data.get('fee')
        total = contract_form.data.get('total')
        remarks = contract_form.data.get('remarks')
        shipping_date_label, shipping_sender = get_shipping_date_label(shipping_method)

        try:
            sub_total = int(sub_total.replace(',', '')) if len(sub_total.replace(',', '')) > 0 else 0
            tax = int(tax.replace(',', '')) if len(tax.replace(',', '')) > 0 else 0
            fee = int(fee.replace(',', '')) if len(fee.replace(',', '')) > 0 else 0
            total = int(total.replace(',', '')) if len(total.replace(',', '')) > 0 else 0
        except ValueError:
            sub_total = 0
            tax = 0
            fee = 0
            total = 0

        response = HttpResponse(content_type='application/ms-excel')
        response['Content-Disposition'] = 'attachment; filename="trader_sales_contract_{}.xls"'.format(contract_id)
        # response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        # response['Content-Disposition'] = 'attachment; filename="trader_sales_contract_{}.xlsx"'.format(contract_id)
        wb = xlwt.Workbook(encoding='utf-8')
        ws = wb.add_sheet('請求書', cell_overwrite_ok=True)

        ws.set_header_str(str.encode(''))
        ws.set_footer_str(str.encode(''))
        ws.set_left_margin(0.314961)
        ws.set_top_margin(0.393701)
        ws.set_right_margin(0.314961)
        ws.set_bottom_margin(0.19685)

        for i in range(11):
            ws.col(i).width = cell_width_list[i]

        for i in range(1, 80):
            ws.row(i).height_mismatch = True
            ws.row(i).height = cell_height

        ws.row(0).height_mismatch = True
        ws.row(0).height = top_padding_height

        ws.row(1).height = header_height

        ws.write_merge(1, 1, 1, 8, '売買契約書　兼　請求書', title_style)
        ws.write_merge(1, 1, 9, 10, created_at, contract_date_style)

        ws.row(2).height = space_height

        ws.row(3).height = company_height
        ws.write_merge(3, 3, 0, 2, company, company_title_style)
        ws.write_merge(3, 3, 3, 4, _('CO.'), company_label_style)

        ws.row(4).height = space_height

        ws.row(5).height = int(20 * 19.35)
        ws.write_merge(5, 5, 1, 2, manager, company_supplier)
        ws.write(5, 3, '', border_bottom)
        ws.write(5, 4, _('Mr.'), company_supplier_label)

        ws.row(6).height = int(20 * 15.75)
        ws.write_merge(6, 6, 0, 6, address, font_10_left)
        ws.write_merge(6, 6, 7, 10, 'P-SENSOR {} {}'.format(_('Member ID'), P_SENSOR_NUMBER), font_10_left)

        ws.row(7).height = int(20 * 12.75)
        ws.write_merge(7, 7, 0, 2, 'TEL {}'.format(tel), font_12_left)
        ws.write_merge(7, 7, 3, 6, 'FAX {}'.format(fax), font_12_left)
        
        ws.write(7, 7, '担当：', in_charge_style)
        ws.write_merge(7, 7, 8, 10, person_in_charge, in_charge_style)

        ws.row(8).height = space_height

        ws.row(9).height = int(20 * 12.95)
        ws.write_merge(9, 9, 0, 6, '商　品　名', product_table_first_th_style)
        ws.write(9, 7, '数量', product_table_th_style)
        ws.write(9, 8, '単　価', product_table_th_style)
        ws.write_merge(9, 9, 9, 10, '金　額', product_table_last_th_style)

        # Product Table
        row_no = 10
        
        if num_of_products:
            for form in product_formset.forms:
                form.is_valid()

                id = form.cleaned_data.get('product_id')
                product_name = Product.objects.get(id=id).name
                quantity = form.cleaned_data.get('quantity', 0)
                price = form.cleaned_data.get('price', 0)
                amount = quantity * price

                if total_number > 6: 
                    continue

                ws.row(row_no).height = int(20 * 15)
                if len(product_name) > 23:
                    ws.write_merge(row_no, row_no, 0, 6, product_name, product_first_content_style_thin)
                else: 
                    ws.write_merge(row_no, row_no, 0, 6, product_name, product_first_content_style)
                ws.write(row_no, 7, quantity, product_content_style)
                ws.write(row_no, 8, price, product_content_style)
                ws.write_merge(row_no, row_no, 9, 10, amount, product_last_content_style)
                row_no += 1

        # Document Table
        if num_of_documents:
            for form in document_formset.forms:
                form.is_valid()

                id = form.cleaned_data.get('document_id')
                document_name = Document.objects.get(id=id).name
                document_name = document_name.replace('（売上）', '').replace('（仕入）', '')
                quantity = form.cleaned_data.get('quantity', 0)
                price = form.cleaned_data.get('price', 0)
                amount = quantity * price

                if total_number > 6: 
                    continue

                ws.row(row_no).height = int(20 * 15)
                ws.write_merge(row_no, row_no, 0, 6, document_name, product_first_content_style)
                ws.write(row_no, 7, quantity, product_content_style)
                ws.write(row_no, 8, price, product_content_style)
                ws.write_merge(row_no, row_no, 9, 10, amount, product_last_content_style)
                row_no += 1

        forRange = 6 - total_number if total_number <= 6 else 6
        for i in range(0, forRange):
            ws.row(row_no).height = int(20 * 15)
            ws.write_merge(row_no, row_no, 0, 6, None, product_first_content_style)
            ws.write(row_no, 7, None, product_content_style)
            ws.write(row_no, 8, None, product_content_style)
            ws.write_merge(row_no, row_no, 9, 10, None, product_last_content_style)
            row_no += 1

        ws.row(row_no).height = int(20 * 15)
        ws.write_merge(row_no, row_no, 0, 1, '{}: '.format(shipping_date_label), shipping_date_label_style)
        ws.write_merge(row_no, row_no, 2, 6, shipping_date, shipping_date_style)
        ws.write_merge(row_no, row_no, 7, 8, '小　計', subtotal_label_style)
        ws.write_merge(row_no, row_no, 9, 10, sub_total, subtotal_value_style)
        row_no += 1

        ws.row(row_no).height = int(20 * 15)
        ws.write(row_no, 0, '※備考', remark_lebel_style)
        ws.write_merge(row_no, row_no, 7, 8, '消費税（10%）', fee_label_style)
        ws.write_merge(row_no, row_no, 9, 10, tax, fee_value_style)
        row_no += 1

        ws.row(row_no).height = int(20 * 15)
        ws.write_merge(row_no, row_no + 1, 0, 6, remarks, remark_content_style)
        ws.write_merge(row_no, row_no, 7, 8, '保険代（非課税）', fee_label_style)
        ws.write_merge(row_no, row_no, 9, 10, fee, fee_value_style)
        row_no += 1

        ws.row(row_no).height = int(20 * 15)
        
        ws.write_merge(row_no, row_no, 7, 8, '合　計', total_label_style)
        ws.write_merge(row_no, row_no, 9, 10, total, total_value_style)
        row_no += 1

        ws.row(row_no).height = space_height
        row_no += 1
        remark_text_list = [
            "※ 売主、買主、双方の署名・捺印が揃った時点で契約が成立したものとします。",
            "※ 契約が成立後、原則として売買契約を解除できないものとします。",
            "※ Ｐ－ＳＥＮＳＯＲあんしん決済使用時は、納品後３日以内にＯＫボタンを押すこととする。",
            "※ 当該商品に不備及び故障がある場合の保障は、納品後３日以内とします。",
            "※ 買主が代金金額の支払いを完了するまでは、売買物件の所有権は売主において留保する。",
            "※ 売買物件の所有権は、引渡しが終了し、支払いが完了した時点で買主に移行するものとする。",
            "※ 下記の場合の売主は、何らかの手続きを経ずして、その選択により期限の利益を喪失せしめて残代金の",
            "  即時支払いを求めるか、又は本契約を解除して売買物件を引き上げる事が出来る。",
            "1)代金の支払いを1回でも怠った場合。",
            "2)仮差押え、仮処分等の執行を受け、整理・和議・破損等の申立てを受けた場合。"
        ]

        for remark_text in remark_text_list:
            ws.row(row_no).height = int(20 * 9.95)
            ws.write(row_no, 0, remark_text, remark_text_style)
            row_no += 1

        ws.row(row_no).height = space_height
        row_no += 1

        product_sender_form = TraderSalesProductSenderForm(self.request.POST)
        document_sender_form = TraderSalesDocumentSenderForm(self.request.POST)
        product_sender_id = self.request.POST.get('product_sender_id')
        product_expected_arrival_date = self.request.POST.get('product_expected_arrival_date')
        product_sender_company = ""
        if product_sender_id:
            product_sender = Sender.objects.get(id=product_sender_id)
            product_sender_company = product_sender.name
        product_sender_address = product_sender_form.data.get('product_sender_address')
        product_sender_tel = product_sender_form.data.get('product_sender_tel')
        product_sender_fax = product_sender_form.data.get('product_sender_fax')
        product_sender_postal_code = product_sender_form.data.get('product_sender_postal_code')

        document_sender_id = self.request.POST.get('document_sender_id')
        document_expected_arrival_date = self.request.POST.get('document_expected_arrival_date')
        document_sender_company = ""
        if document_sender_id:
            document_sender = Sender.objects.get(id=document_sender_id)
            document_sender_company = document_sender.name
        document_sender_address = document_sender_form.data.get('document_sender_address')
        document_sender_tel = document_sender_form.data.get('document_sender_tel', "")
        document_sender_fax = document_sender_form.data.get('document_sender_fax', "")
        document_sender_postal_code = document_sender_form.data.get('document_sender_postal_code')

        ws.row(row_no).height = int(20 * 11.45)
        ws.write_merge(row_no, row_no, 0, 5, '【商品発送先】', sender_table_title_style)
        ws.write_merge(row_no, row_no, 6, 10, '【書類発送先】', sender_table_title_style)
        row_no += 1

        ws.row(row_no).height = space_height
        ws.write(row_no, 0, '', border_left)
        ws.write(row_no, 6, '', border_left)
        ws.write(row_no, 5, '', border_right)
        ws.write(row_no, 10, '', border_right)
        row_no += 1
        
        ws.row(row_no).height = int(20 * 15)
        ws.write(row_no, 0, '会社名：', sender_table_field_style)
        ws.write_merge(row_no, row_no, 1, 5, product_sender_company, sender_table_text_style)
        ws.write(row_no, 6, '会社名：', sender_table_field_style)
        ws.write_merge(row_no, row_no, 7, 10, document_sender_company, sender_table_text_style)
        row_no += 1

        for i in range(0, 6):
            ws.row(row_no + i).height = int(20 * 15)
        ws.write(row_no, 0, '住所：', sender_table_field_style)
        ws.write(row_no, 6, '住所：', sender_table_field_style)
        ws.write_merge(row_no, row_no + 3, 1, 5, '〒 {} \n {}'.format(product_sender_postal_code, product_sender_address), sender_address_style)
        ws.write_merge(row_no, row_no + 3, 7, 10, '〒 {} \n {}'.format(document_sender_postal_code, document_sender_address), sender_address_style)
        row_no += 1

        ws.write(row_no, 0, '', border_left)
        ws.write(row_no, 6, '', sender_table_field_style)
        row_no += 1

        ws.write(row_no, 0, '', border_left)
        ws.write(row_no, 6, '', sender_table_field_style)
        row_no += 1

        ws.write(row_no, 0, '', border_left)
        ws.write(row_no, 6, '', sender_table_field_style)
        row_no += 1

        ws.write(row_no, 0, 'TEL／FAX：', sender_table_field_style)
        ws.write_merge(row_no, row_no, 1, 5, '{} / {}'.format(product_sender_tel, product_sender_fax), sender_table_text_style)
        ws.write(row_no, 6, 'TEL／FAX：', sender_table_field_style)
        ws.write_merge(row_no, row_no, 7, 10, '{} / {}'.format(document_sender_tel, document_sender_fax), sender_table_text_style)
        row_no += 1

        ws.write(row_no, 0, '到着予定日：', sender_table_field_style)
        ws.write(row_no, 6, '到着予定日：', sender_table_field_style)
        ws.write_merge(row_no, row_no, 1, 5, product_expected_arrival_date, sender_table_text_style)
        ws.write_merge(row_no, row_no, 7, 10, document_expected_arrival_date, sender_table_text_style)
        row_no += 1

        ws.row(row_no).height = space_height
        ws.write(row_no, 0, '', border_left)
        ws.write(row_no, 6, '', border_left)
        ws.write(row_no, 10, '', border_right)
        row_no += 1

        ws.row(row_no).height = space_height
        for i in range(0, 11):
            ws.write(row_no, i, '', border_top)
        row_no += 1

        ws.row(row_no).height = int(20 * 11.45)
        ws.write_merge(row_no, row_no, 0, 5, '【買主 署名捺印欄】', sender_table_title_style)
        ws.write_merge(row_no, row_no, 6, 10, '【売主 署名捺印欄】', sender_table_title_style)
        row_no += 1

        ws.row(row_no).height = space_height
        ws.write(row_no, 6, '', border_left)
        ws.write(row_no, 0, '', border_left)
        ws.write(row_no, 5, '', border_right)
        ws.write(row_no, 10, '', border_right)
        row_no += 1

        for i in range(0, 5):
            ws.row(row_no + i).height = int(20 * 16.5)
            ws.write(row_no + i, 0, '', border_left)
        
        # ws.write_merge(row_no, row_no, 0, 5, '〒537-0021　大阪府大阪市東成区東中本2丁目4-15', seal_address_tel_style)
        ws.write_merge(row_no, row_no, 6, 10, '〒537-0021　大阪府大阪市東成区東中本2丁目4-15', seal_address_tel_style1)
        row_no += 1
        
        # ws.write_merge(row_no, row_no + 1, 0, 5, 'バッジオ株式会社', seal_company_style)
        ws.write_merge(row_no, row_no + 1, 6, 10, 'バッジオ株式会社', seal_company_style)
        row_no += 2

        # ws.write_merge(row_no, row_no, 0, 4, '代表取締役　金　昇志', seal_supplier_style)
        ws.write_merge(row_no, row_no, 6, 9, '代表取締役　金　昇志', seal_supplier_style)
        ws.write(row_no, 5, '㊞', seal_mark_style)
        ws.write(row_no, 10, '㊞', seal_mark_style)
        row_no += 1

        # ws.write_merge(row_no, row_no, 0, 5, 'TEL 06-6753-8078 FAX 06-6753-8079', seal_address_tel_style)
        ws.write_merge(row_no, row_no, 6, 10, 'TEL 06-6753-8078 FAX 06-6753-8079', seal_address_tel_style)
        row_no += 1

        ws.row(row_no).height = space_height
        ws.write(row_no, 0, '', border_left)
        ws.write(row_no, 6, '', border_left)
        ws.write(row_no, 10, '', border_right)
        row_no += 1

        ws.row(row_no).height = space_height
        for i in range(0, 11):
            ws.write(row_no, i, '', border_top)
        row_no += 1

        ws.write(row_no, 5, '下記の通り御請求申し上げます。', font_11_left)
        row_no += 1

        ws.row(row_no).height = int(20 * 13.5)
        ws.write_merge(row_no, row_no, 0, 2, '運送方法', font_11_center_with_border)
        ws.write_merge(row_no, row_no + 1, 5, 7, '御請求金額', billed_deadline_label_style)
        ws.write_merge(row_no, row_no + 1, 8, 10, total, billing_amount_value_style)
        row_no += 1
        ws.row(row_no).height = int(20 * 13.5)

        shipping_method_label = ''
        if shipping_method == 'D': shipping_method_label = '発送'
        if shipping_method == 'R': shipping_method_label = '引取'
        if shipping_method == 'C': shipping_method_label = 'ID変更'
        if shipping_method == 'B': shipping_method_label = '* 空白'

        ws.write_merge(row_no, row_no, 0, 2, shipping_method_label, font_11_center_with_border)
        row_no += 1

        ws.row(row_no).height = int(20 * 8.25)
        row_no += 1

        ws.row(row_no).height = int(20 * 13.5)
        ws.row(row_no+1).height = int(20 * 13.5)
        ws.write_merge(row_no, row_no, 0, 2, 'お支払方法', font_11_center_with_border)
        ws.write_merge(row_no + 1, row_no + 1, 0, 2, '振込', font_11_center_with_border)

        ws.write_merge(row_no, row_no + 1, 5, 7, 'お支払期限', billed_deadline_label_style)
        ws.write_merge(row_no, row_no + 1, 8, 10, payment_due_date, billed_deadline_value_style)
        row_no += 2

        

        ws.row(row_no).height = space_height
        row_no += 1

        ws.row(row_no).height = int(20 * 14.1)
        ws.write(row_no, 0, '※ お振込み手数料は貴社ご負担でお願い致します。', font_11_left)
        row_no += 1

        ws.row(row_no).height = int(20 * 15)
        ws.write_merge(row_no, row_no, 0, 1, '振込先口座', transfer_account_style)
        ws.write_merge(row_no, row_no, 2, 10, 'りそな銀行　船場支店（101）　普通　0530713　バッジオカブシキガイシャ', font_11_left_with_border)
        row_no += 1
        ws.row(row_no).height = int(20 * 15)

        # ws.write(row_no, 0, '', border_left)
        # ws.write_merge(row_no, row_no, 2, 10, '', font_11_left_with_border)
        # row_no += 1

        ws.write(row_no, 0, '', border_left)
        ws.write_merge(row_no, row_no, 2, 10, '', font_11_left_with_border)
        row_no += 1

        ws.row(row_no).height = space_height
        for i in range(0, 11):
            ws.write(row_no, i, '', border_top)
        row_no += 1

        ws.row(row_no).height = int(20 * 17.25)
        ws.write(row_no, 0, '内容をご確認の上、ご捺印後弊社まで返信お願いします。', font_11_left)
        ws.write(row_no, 7, 'FAX 06-6753-8079', font_14_left1)
        row_no += 1

        ws.row(row_no).height = int(20 * 14.1)
        ws.write(row_no, 0, '※ 請求書原本が必要な場合は必ずチェックしてください。　【　必要　・　不要　】', font_11_left)
        row_no += 1

        ws.row(row_no).height = int(20 * 15)
        ws.write(row_no, 0, 'No.{}'.format(contract_id), xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: vert center, horiz left'))
        
        # -------------------------------------------------------------------------------------------------------------------------
        if total_number > 6:
            ws = wb.add_sheet('別紙', cell_overwrite_ok=True)

            ws.set_header_str(str.encode(''))
            ws.set_footer_str(str.encode(''))
            ws.set_left_margin(0.314961)
            ws.set_top_margin(0.393701)
            ws.set_right_margin(0.314961)
            ws.set_bottom_margin(0.19685)

            w_arr = [9.15, 8.04, 10.04, 2.04, 5.04, 8.04, 6.71, 7.15, 12.59, 6.71, 6.71]
            for i in range(11): 
                ws.col(i).width = int(300 * w_arr[i])
            for i in range(56):
                ws.row(i).height_mismatch = True
                ws.row(i).height = int(20 * 18)
            ws.row(44).height = int(20 * 3.75)
            ws.row(45).height = int(20 * 15.75)

            ####################################### Table #####################################
            ws.write_merge(0, 0, 0, 6, '商　品　名', TTL_CENTER)
            ws.write(0, 7, '数量', TT_CENTER)
            ws.write(0, 8, '単　価', TT_CENTER)
            ws.write_merge(0, 0, 9, 10, '金　額', TTR_CENTER)

            row_no = 1


            total_quantity = total_price = 0
            ##### Products
            if num_of_products:
                for form in product_formset.forms:
                    form.is_valid()

                    id = form.cleaned_data.get('product_id')
                    product_name = Product.objects.get(id=id).name
                    quantity = form.cleaned_data.get('quantity', 0)
                    price = form.cleaned_data.get('price', 0)
                    amount = quantity * price

                    total_quantity += quantity
                    total_price += price

                    ws.write_merge(row_no, row_no, 0, 6, product_name, TL)
                    ws.write(row_no, 7, quantity, TC_RIGHT)
                    ws.write(row_no, 8, price, TC_RIGHT)
                    ws.write_merge(row_no, row_no, 9, 10, amount, TR_RIGHT)

                    row_no += 1

            ###### Documents
            if num_of_documents:
                for form in document_formset.forms:
                    form.is_valid()

                    id = form.cleaned_data.get('document_id')
                    document_name = Document.objects.get(id=id).name.replace('（売上）', '').replace('（仕入）', '')
                    quantity = form.cleaned_data.get('quantity', 0)
                    price = form.cleaned_data.get('price', 0)
                    amount = quantity * price

                    total_quantity += quantity
                    total_price += price

                    ws.write_merge(row_no, row_no, 0, 6, document_name, TL)
                    ws.write(row_no, 7, quantity, TC_RIGHT)
                    ws.write(row_no, 8, price, TC_RIGHT)
                    ws.write_merge(row_no, row_no, 9, 10, amount, TR_RIGHT)

                    row_no += 1

            if total_number - 1 < 39:
                for i in range(0, 39 - total_number + 1):

                    ws.write_merge(row_no, row_no, 0, 6, None, TL)
                    ws.write(row_no, 7, None, TC_RIGHT)
                    ws.write(row_no, 8, None, TC_RIGHT)
                    ws.write_merge(row_no, row_no, 9, 10, None, TR_RIGHT)

                    row_no += 1

            ws.write_merge(40, 40, 0, 6, '※備考', LINE_TOP)
            ws.write_merge(40, 40, 7, 8, '小　計', TTL_CENTER)
            ws.write_merge(40, 40, 9, 10, sub_total, TTR_RIGHT)
            ws.write_merge(41, 41, 7, 8, '消費税（10%）', TL_CENTER)        
            ws.write_merge(41, 41, 9, 10, tax, TR_RIGHT)
            ws.write_merge(42, 42, 7, 8, '保険代（非課税）', TL_CENTER)
            ws.write_merge(42, 42, 9, 10, fee, TR_RIGHT)
            ws.write_merge(43, 43, 7, 8, '合　計', TBL_CENTER)
            ws.write_merge(43, 43, 9, 10, total, TBR_RIGHT_TOTAL)

            ws.write_merge(45, 45, 0, 3, 'No.{}'.format(contract_id), F11)

        wb.save(response)
        return response

class TraderPurchasesInvoiceView(AdminLoginRequiredMixin, View):
    def post(self, *args, **kwargs):
        
        ############################### Styles ################################
        top_padding_height = int(256 * 0.4)
        cell_width_list = [ 256 *  10, 256 * 9, int(256 * 12.5), int(256 * 2.0), 
                            int(256 * 4.5), int(256 * 7.5), 256 * 10, int(256 * 6.5),
                            int(256 * 13.5), int(256 * 8.5), int(256 * 7.5)]
        cell_height = int(20 * 15)

        header_height = int(20 * 20.1)
        space_height = int(20 * 4.35)
        company_height = int(20 * 19.25)
        height_18 = int(20 * 16)
        height_16 = int(20 * 14)
        height_16_1 = int(20 * 15)
        height_16_5 = int(20 * 14.5)
        height_15 = 20 * 13
        height_13_5 = int(20 * 12.5)
        height_12 = 20 * 10

        border_bottom = xlwt.easyxf('borders: bottom_color black, bottom thin;')
        border_left = xlwt.easyxf('borders: left_color black, left thin;')
        border_right = xlwt.easyxf('borders: right_color black, right thin;')
        border_top = xlwt.easyxf('borders: top_color black, top thin;')
        bold_medium_top = xlwt.easyxf('borders: top_color black, top medium;')

        font_11_left = xlwt.easyxf('font: height 190, name ＭＳ Ｐゴシック; align: vert center, horiz left')
        font_12_left = xlwt.easyxf('font: height 240, name ＭＳ Ｐゴシック; align: vert center, horiz left')
        font_11_left_r = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: vert center, horiz left')
        font_14_left = xlwt.easyxf('font: height 260, name ＭＳ Ｐゴシック; align: vert center, horiz left')
        font_14_left1 = xlwt.easyxf('font: height 280, name ＭＳ Ｐゴシック; align: vert center, horiz left')
        font_11_center_with_border = xlwt.easyxf('font: height 190, name ＭＳ Ｐゴシック; align: vert center, horiz center;\
                                                borders: top_color black, left_color black, right_color black, bottom_color black,\
                                                top thin, bottom thin, left thin, right thin;')
        font_11_left_with_border = xlwt.easyxf('font: height 190, name ＭＳ Ｐゴシック; align: vert center, horiz left;\
                                                borders: top_color black, left_color black, right_color black, bottom_color black,\
                                                top thin, bottom thin, left thin, right thin;')

        title_style = xlwt.easyxf('font: height 360, name ＭＳ Ｐゴシック, color black;\
                                    align: vert center, horiz center, wrap on;')
        contract_date_style = xlwt.easyxf('font: height 190, name ＭＳ Ｐゴシック; align: vert bottom, horiz center, wrap on;')
        company_title_style = xlwt.easyxf('font: bold on, height 320, name ＭＳ Ｐゴシック; align: vert center, horiz center, wrap off, shrink on;\
                                            borders: bottom_color black, bottom thin;')
        company_label_style = xlwt.easyxf('font: height 260, name ＭＳ Ｐゴシック; align: vert center, horiz center, wrap on;\
                                            borders: bottom_color black, bottom thin;')
        in_charge_style = xlwt.easyxf('font: height 190, name ＭＳ Ｐゴシック; align: vert center, horiz left, wrap off;\
                                        borders: bottom_color black, bottom thin;')
        company_supplier = xlwt.easyxf('font: height 180, name ＭＳ Ｐゴシック; align: vert center, horiz center, wrap on;\
                                        borders: bottom_color black, bottom thin;')
        company_supplier_label = xlwt.easyxf('font: height 180, name ＭＳ Ｐゴシック; align: vert center, horiz right, wrap on;\
                                            borders: bottom_color black, bottom thin;')
        product_table_first_th_style = xlwt.easyxf('font: height 190, name ＭＳ Ｐゴシック; align: vert center, horiz center, wrap on;\
                                            borders: top_color black, left_color black, right_color black, top medium, left medium, right thin;')
        product_table_th_style = xlwt.easyxf('font: height 190, name ＭＳ Ｐゴシック; align: vert center, horiz center, wrap on;\
                                            borders: top_color black, right_color black, top medium, right thin;')
        product_table_last_th_style = xlwt.easyxf('font: height 190, name ＭＳ Ｐゴシック; align: vert center, horiz center, wrap on;\
                                            borders: top_color black, right_color black, top medium, right medium;')
        product_first_content_style = xlwt.easyxf('font: height 240, name ＭＳ Ｐゴシック; align: vert center, horiz left, wrap no;\
                                            borders: top_color black, left_color black, right_color black, bottom_color black, bottom thin, top thin, left medium, right thin;')
        product_first_content_style_thin = xlwt.easyxf('font: height 180, name ＭＳ Ｐゴシック; align: vert center, horiz left, wrap no;\
                                            borders: top_color black, left_color black, right_color black, bottom_color black, bottom thin, top thin, left medium, right thin;')
        product_content_style = xlwt.easyxf('font: height 240, name ＭＳ ゴシック; align: vert center, horiz right, wrap on;\
                                            borders: top_color black, right_color black, bottom_color black, bottom thin, top thin, right thin;',
                                            num_format_str='#,##0')
        product_last_content_style = xlwt.easyxf('font: height 240, name ＭＳ Ｐゴシック; align: vert center, horiz right, wrap on;\
                                            borders: top_color black, right_color black, bottom_color black, bottom thin, top thin, right medium;',
                                            num_format_str='#,##0')
        shipping_date_style = xlwt.easyxf('font: height 240, name ＭＳ Ｐゴシック; align: vert center, horiz left, wrap on;\
                                            borders: top_color black, top medium;')
        shipping_date_label_style = xlwt.easyxf('font: height 190, name ＭＳ Ｐゴシック; align: vert center, horiz left, wrap on;\
                                            borders: top_color black, left_color black, left thin, top medium;')
        remark_lebel_style = xlwt.easyxf('font: height 190, name ＭＳ Ｐゴシック; align: vert center, horiz left;\
                                        borders: left_color black, left thin;')
        remark_content_style = xlwt.easyxf('font: height 190, name ＭＳ Ｐゴシック; align: vert center, horiz left;\
                                        borders: left_color black, bottom_color black, bottom thin, left thin;')
        subtotal_label_style = xlwt.easyxf('font: height 190, name ＭＳ Ｐゴシック; align: vert center, horiz center, wrap on;\
                                            borders: top_color black, left_color black, top medium, left medium;')
        subtotal_value_style = xlwt.easyxf('font: height 240, name ＭＳ ゴシック; align: vert center, horiz right, wrap on;\
                                            borders: top_color black, left_color black, right_color black, top medium, left medium, right medium;',
                                            num_format_str='#,##0')
        fee_label_style = xlwt.easyxf('font: height 190, name ＭＳ Ｐゴシック; align: vert center, horiz center, wrap on;\
                                            borders: top_color black, left_color black, top thin, left medium;')
        fee_value_style = xlwt.easyxf('font: height 240, name ＭＳ ゴシック; align: vert center, horiz right, wrap on;\
                                            borders: top_color black, left_color black, right_color black, top thin, left medium, right medium;',
                                            num_format_str='#,##0')
        total_label_style = xlwt.easyxf('font: height 190, name ＭＳ Ｐゴシック; align: vert center, horiz center, wrap on;\
                                            borders: top_color black, left_color black, bottom_color black,\
                                            bottom medium, top thin, left medium;')
        total_value_style = xlwt.easyxf('font: height 240, bold on, name ＭＳ ゴシック; align: vert center, horiz right, wrap on;\
                                            borders: top_color black, left_color black, right_color black, bottom_color black,\
                                            bottom medium, top thin, left medium, right medium;',
                                            num_format_str='"¥"#,###')
        remark_text_style = xlwt.easyxf('font: height 180, name ＭＳ Ｐゴシック; align: vert center, horiz left;')
        sender_table_title_style = xlwt.easyxf('font: bold on, height 180, name ＭＳ Ｐゴシック; align: vert center, horiz center;\
                                                borders: bottom_color black, bottom thin;')
        sender_table_field_style = xlwt.easyxf('font: height 180, name ＭＳ Ｐゴシック; align: vert center, horiz left;\
                                                borders: left_color black, left thin;')
        sender_address_style = xlwt.easyxf('font: height 190, name ＭＳ Ｐゴシック; align: vert top, horiz left; \
                                            borders: right_color black, right thin; alignment: wrap 1;')
        sender_table_text_style = xlwt.easyxf('font: height 190, name ＭＳ Ｐゴシック; align: vert center, horiz left;\
                                                borders: right_color black, right thin;')
        seal_address_tel_style = xlwt.easyxf('font: height 240, name ＭＳ Ｐゴシック; alignment: wrap True; align: vert center, horiz left;\
                                                borders: right_color black, left_color black, left thin, right thin;')
        seal_address_tel_style1 = xlwt.easyxf('font: height 190, name ＭＳ Ｐゴシック; alignment: wrap True; align: vert center, horiz center;\
                                                borders: right_color black, left_color black, left thin, right thin;')
        seal_company_style = xlwt.easyxf('font: height 320, name ＭＳ Ｐゴシック; align: vert center, horiz center;\
                                                borders: right_color black, left_color black, left thin, right thin;')
        seal_supplier_style = xlwt.easyxf('font: height 190, name ＭＳ Ｐゴシック; align: vert center, horiz center;\
                                                borders: left_color black, left thin;')
        seal_mark_style = xlwt.easyxf('font: height 190, name ＭＳ Ｐゴシック; align: vert center, horiz left;\
                                                borders: right_color black, right thin;')
        billed_deadline_label_style = xlwt.easyxf('font: height 320, bold on, name ＭＳ Ｐゴシック; align: vert center, horiz center;\
                                                borders: right_color black, top_color black, bottom_color black, left_color black,\
                                                right thin, top medium, left medium, bottom medium;')
        billing_amount_value_style = xlwt.easyxf('font: height 320, bold on, name ＭＳ Ｐゴシック; align: vert center, horiz center;\
                                                borders: right_color black, top_color black, bottom_color black, left_color black,\
                                                right medium, top medium, left thin, bottom medium;',
                                                num_format_str='"¥"#,###')
        billed_deadline_value_style = xlwt.easyxf('font: height 320, bold on, name ＭＳ Ｐゴシック; align: vert center, horiz center;\
                                                borders: right_color black, top_color black, bottom_color black, left_color black,\
                                                right medium, top medium, left thin, bottom medium;')
        transfer_account_style = xlwt.easyxf('font: height 190, name ＭＳ Ｐゴシック; align: vert center, horiz center;\
                                                borders: left_color black, top_color black, top thin, left thin;')

        TTL_CENTER = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz center, vert center, wrap on; borders: top_color black, left_color black, bottom_color black, right_color black, top medium, left medium, right thin, bottom thin;')
        TT_CENTER = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz center, vert center, wrap on; borders: top_color black, bottom_color black, right_color black, top medium, right thin, bottom thin;')
        TTR_CENTER = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz center, vert center, wrap on; borders: top_color black, left_color black, bottom_color black, right_color black, top medium, left thin, right medium, bottom thin;')
        TC_RIGHT = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz right, wrap on; borders: bottom_color black, right_color black, right thin, bottom thin;', num_format_str='#,###')
        TR_RIGHT = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz right, wrap on; borders: bottom_color black, right_color black, right medium, bottom thin;', num_format_str='#,###')
        F11 = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: wrap on;')
        LINE_TOP = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: wrap on; borders: top_color black, top medium;')
        TTR_RIGHT = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz right, wrap on; borders: top_color black, left_color black, bottom_color black, right_color black, top medium, left thin, right medium, bottom thin;', num_format_str='#,###')
        TL_CENTER = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz center, vert center, wrap on; borders: bottom_color black, left_color black, right_color black, right thin, bottom thin, left medium;')
        TBL_CENTER = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz center, vert center, wrap on; borders: left_color black, bottom_color black, right_color black, left medium, right thin, bottom medium;')
        TBR_RIGHT_TOTAL = xlwt.easyxf('font: height 240, name ＭＳ Ｐゴシック, bold on; align: horiz right, vert center, wrap on; borders: bottom_color black, right_color black, right medium, bottom medium;', num_format_str='"¥"#,###')
        TL = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: wrap on; borders: bottom_color black, left_color black, right_color black, right thin, bottom thin, left medium;')

        user_id = self.request.user.id
        log_export_operation(user_id, "{} - {}".format(_("Sales contract"), _("Trader sales")))
        contract_form = TraderPurchasesContractForm(self.request.POST)
        contract_id = contract_form.data.get('contract_id', '')
        created_at = contract_form.data.get('created_at', '')
        updated_at = contract_form.data.get('updated_at', '')
        person_in_charge = contract_form.data.get('person_in_charge', '')
        customer_id = contract_form.data.get('customer_id')
        manager = contract_form.data.get('manager')
        company = frigana = postal_code = address = tel = fax = ""
        if customer_id:
            customer = Customer.objects.get(id=customer_id)
            company = customer.name
            frigana = customer.frigana
            postal_code = customer.postal_code
            address = customer.address
            tel = customer.tel
            fax = customer.fax

        product_formset = ProductFormSet(
            self.request.POST,
            prefix='product'
        )

        num_of_products = product_formset.total_form_count()

        document_formset = DocumentFormSet(
            self.request.POST,
            prefix='document'
        )
        num_of_documents = document_formset.total_form_count()

        total_number = num_of_products + num_of_documents

        shipping_method = contract_form.data.get('shipping_method')
        shipping_date = contract_form.data.get('shipping_date')
        payment_method = contract_form.data.get('payment_method')
        payment_due_date = contract_form.data.get('payment_due_date')
        sub_total = contract_form.data.get('sub_total')
        tax = contract_form.data.get('tax')
        fee = contract_form.data.get('fee')
        total = contract_form.data.get('total')
        remarks = contract_form.data.get('remarks')
        shipping_date_label, shipping_sender = get_shipping_date_label(shipping_method)

        try:
            sub_total = int(sub_total.replace(',', '')) if len(sub_total.replace(',', '')) > 0 else 0
            tax = int(tax.replace(',', '')) if len(tax.replace(',', '')) > 0 else 0
            fee = int(fee.replace(',', '')) if len(fee.replace(',', '')) > 0 else 0
            total = int(total.replace(',', '')) if len(total.replace(',', '')) > 0 else 0
        except ValueError:
            sub_total = 0
            tax = 0
            fee = 0
            total = 0

        response = HttpResponse(content_type='application/ms-excel')
        response['Content-Disposition'] = 'attachment; filename="trader_purchases_contract_{}.xls"'.format(contract_id)
        # response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        # response['Content-Disposition'] = 'attachment; filename="trader_purchases_contract_{}.xlsx"'.format(contract_id)
        wb = xlwt.Workbook(encoding='utf-8')
        ws = wb.add_sheet("{}-{}".format(_('Sales contract'), _('Trader purchases')), cell_overwrite_ok=True)

        ws.set_header_str(str.encode(''))
        ws.set_footer_str(str.encode(''))
        ws.set_left_margin(0.314961)
        ws.set_top_margin(0.393701)
        ws.set_right_margin(0.314961)
        ws.set_bottom_margin(0.19685)

        for i in range(8):
            ws.col(i).width = cell_width_list[i]

        for i in range(1, 80):
            ws.row(i).height_mismatch = True
            ws.row(i).height = cell_height


        ws.row(0).height_mismatch = True
        ws.row(0).height = top_padding_height

        ws.row(1).height = header_height

        ws.write_merge(1, 1, 1, 8, '売買契約書　兼　請求書', title_style)
        ws.write_merge(1, 1, 9, 10, created_at, contract_date_style)

        ws.row(2).height = space_height
        ws.row(3).height = company_height
        ws.write_merge(3, 3, 0, 2, company, company_title_style)
        ws.write_merge(3, 3, 3, 4, _('CO.'), company_label_style)

        ws.row(4).height = space_height

        ws.row(5).height = int(20 * 19.35)
        ws.write_merge(5, 5, 1, 2, manager, company_supplier)
        ws.write(5, 3, '', border_bottom)
        ws.write(5, 4, _('Mr.'), company_supplier_label)

        ws.row(6).height = int(20 * 15.75)
        ws.write_merge(6, 6, 0, 6, address, font_11_left)
        ws.write_merge(6, 6, 7, 10, 'P-SENSOR {} {}'.format(_('Member ID'), P_SENSOR_NUMBER), font_11_left)

        ws.row(7).height = int(20 * 12.75)
        ws.write_merge(7, 7, 0, 2, 'TEL {}'.format(tel), font_11_left_r)
        ws.write_merge(7, 7, 3, 6, 'FAX {}'.format(fax), font_11_left_r)

        ws.write(7, 7, '担当：', in_charge_style)
        ws.write_merge(7, 7, 8, 10, person_in_charge, in_charge_style)

        ws.row(8).height = space_height
        ws.row(9).height = int(20 * 12.95)
        ws.write_merge(9, 9, 0, 6, '商　品　名', product_table_first_th_style)
        ws.write(9, 7, '数量', product_table_th_style)
        ws.write(9, 8, '単　価', product_table_th_style)
        ws.write_merge(9, 9, 9, 10, '金　額', product_table_last_th_style)

        row_no = 10

        if num_of_products:
            for form in product_formset.forms:
                form.is_valid()

                id = form.cleaned_data.get('product_id')
                product_name = Product.objects.get(id=id).name
                quantity = form.cleaned_data.get('quantity', 0)
                price = form.cleaned_data.get('price', 0)
                amount = quantity * price

                if total_number > 6: 
                    continue

                ws.row(row_no).height = int(20 * 15)
                if len(product_name) > 23:
                    ws.write_merge(row_no, row_no, 0, 6, product_name, product_first_content_style_thin)
                else:
                    ws.write_merge(row_no, row_no, 0, 6, product_name, product_first_content_style)
                ws.write(row_no, 7, quantity, product_content_style)
                ws.write(row_no, 8, price, product_content_style)
                ws.write_merge(row_no, row_no, 9, 10, amount, product_last_content_style)
                row_no += 1
        
        if num_of_documents:
            for form in document_formset.forms:
                form.is_valid()

                
                id = form.cleaned_data.get('document_id')
                document_name = Document.objects.get(id=id).name
                document_name = document_name.replace('（売上）', '').replace('（仕入）', '')
                quantity = form.cleaned_data.get('quantity', 0)
                price = form.cleaned_data.get('price', 0)
                amount = quantity * price

                if total_number > 6: 
                    continue

                ws.row(row_no).height = int(20 * 15)
                ws.write_merge(row_no, row_no, 0, 6, document_name, product_first_content_style)
                ws.write(row_no, 7, quantity, product_content_style)
                ws.write(row_no, 8, price, product_content_style)
                ws.write_merge(row_no, row_no, 9, 10, amount, product_last_content_style)
                row_no += 1

        forRange = 6 - total_number if total_number <= 6 else 6
        for i in range(0, forRange):
            ws.row(row_no).height = int(20 * 15)
            ws.write_merge(row_no, row_no, 0, 6, None, product_first_content_style)
            ws.write(row_no, 7, None, product_content_style)
            ws.write(row_no, 8, None, product_content_style)
            ws.write_merge(row_no, row_no, 9, 10, None, product_last_content_style)
            row_no += 1

        # row_no += 1
        ws.row(row_no).height = int(20 * 15)
        ws.write_merge(row_no, row_no, 0, 1, '{}: '.format(shipping_date_label), shipping_date_label_style)
        ws.write_merge(row_no, row_no, 2, 6, shipping_date, shipping_date_style)
        ws.write_merge(row_no, row_no, 7, 8, '小　計', subtotal_label_style)
        ws.write_merge(row_no, row_no, 9, 10, sub_total, subtotal_value_style)
        row_no += 1

        ws.row(row_no).height = int(20 * 15)
        ws.write(row_no, 0, '※備考', remark_lebel_style)
        ws.write_merge(row_no, row_no, 7, 8, '消費税（10%）', fee_label_style)
        ws.write_merge(row_no, row_no, 9, 10, tax, fee_value_style)
        row_no += 1

        ws.row(row_no).height = int(20 * 15)
        ws.write_merge(18, 19, 0, 6, remarks, remark_content_style)
        ws.write_merge(18, 18, 7, 8, '保険代（非課税）', fee_label_style)
        ws.write_merge(18, 18, 9, 10, fee, fee_value_style)
        row_no += 1

        ws.row(row_no).height = int(20 * 15)
        
        ws.write_merge(row_no, row_no, 7, 8, '合　計', total_label_style)
        ws.write_merge(row_no, row_no, 9, 10, total, total_value_style)
        row_no += 1

        ws.row(row_no).height = space_height
        row_no += 1
        remark_text_list = [
            "※ 売主、買主、双方の署名・捺印が揃った時点で契約が成立したものとします。",
            "※ 契約が成立後、原則として売買契約を解除できないものとします。",
            "※ Ｐ－ＳＥＮＳＯＲあんしん決済使用時は、納品後３日以内にＯＫボタンを押すこととする。",
            "※ 当該商品に不備及び故障がある場合の保障は、納品後３日以内とします。",
            "※ 買主が代金金額の支払いを完了するまでは、売買物件の所有権は売主において留保する。",
            "※ 売買物件の所有権は、引渡しが終了し、支払いが完了した時点で買主に移行するものとする。",
            "※ 下記の場合の売主は、何らかの手続きを経ずして、その選択により期限の利益を喪失せしめて残代金の",
            "  即時支払いを求めるか、又は本契約を解除して売買物件を引き上げる事が出来る。",
            "1)代金の支払いを1回でも怠った場合。",
            "2)仮差押え、仮処分等の執行を受け、整理・和議・破損等の申立てを受けた場合。"
        ]

        for remark_text in remark_text_list:
            ws.row(row_no).height = int(20 * 9.95)
            ws.write(row_no, 0, remark_text, remark_text_style)
            row_no += 1

        ws.row(row_no).height = space_height
        row_no += 1

        product_sender_form = TraderPurchasesProductSenderForm(self.request.POST)
        document_sender_form = TraderPurchasesDocumentSenderForm(self.request.POST)
        product_sender_id = self.request.POST.get('product_sender_id')
        product_shipping_company = self.request.POST.get('product_shipping_company')
        product_sender_remarks = self.request.POST.get('product_sender_remarks')
        product_desired_arrival_date = self.request.POST.get('product_desired_arrival_date')
        product_sender_company = None
        if product_sender_id:
            product_sender = Sender.objects.get(id=product_sender_id)
            product_sender_company = product_sender.name
        product_sender_address = product_sender_form.data.get('product_sender_address')
        product_sender_tel = product_sender_form.data.get('product_sender_tel') or ''
        product_sender_fax = product_sender_form.data.get('product_sender_fax') or ''
        product_sender_postal_code = product_sender_form.data.get('product_sender_postal_code')

        document_sender_company = None
        document_sender_id = self.request.POST.get('document_sender_id')
        document_shipping_company = self.request.POST.get('document_shipping_company')
        document_sender_remarks = self.request.POST.get('document_sender_remarks')
        document_desired_arrival_date = self.request.POST.get('document_desired_arrival_date')
        if document_sender_id:
            document_sender = Sender.objects.get(id=document_sender_id)
            document_sender_company = document_sender.name
        document_sender_address = document_sender_form.data.get('document_sender_address')
        document_sender_tel = document_sender_form.data.get('document_sender_tel') or ''
        document_sender_fax = document_sender_form.data.get('document_sender_fax') or ''
        document_sender_postal_code = document_sender_form.data.get('document_sender_postal_code')

        ws.row(row_no).height = int(20 * 11.45)
        ws.write_merge(row_no, row_no, 0, 5, '【商品発送先】', sender_table_title_style)
        ws.write_merge(row_no, row_no, 6, 10, '【書類発送先】', sender_table_title_style)
        row_no += 1

        ws.row(row_no).height = space_height
        ws.write(row_no, 0, '', border_left)
        ws.write(row_no, 6, '', border_left)
        ws.write(row_no, 5, '', border_right)
        ws.write(row_no, 10, '', border_right)
        row_no += 1
        
        ws.row(row_no).height = int(20 * 15)
        ws.write(row_no, 0, '会社名：', sender_table_field_style)
        ws.write_merge(row_no, row_no, 1, 5, product_sender_company, sender_table_text_style)
        ws.write(row_no, 6, '会社名：', sender_table_field_style)
        ws.write_merge(row_no, row_no, 7, 10, document_sender_company, sender_table_text_style)
        row_no += 1

        for i in range(0, 6):
            ws.row(row_no + i).height = int(20 * 15)
        ws.write(row_no, 0, '住所：', sender_table_field_style)
        ws.write(row_no, 6, '住所：', sender_table_field_style)
        ws.write_merge(row_no, row_no + 3, 1, 5, '〒 {} \n {}'.format(product_sender_postal_code, product_sender_address), sender_address_style)
        ws.write_merge(row_no, row_no + 3, 7, 10, '〒 {} \n {}'.format(document_sender_postal_code, document_sender_address), sender_address_style)
        row_no += 1

        ws.write(row_no, 0, '', border_left)
        ws.write(row_no, 6, '', sender_table_field_style)
        row_no += 1

        ws.write(row_no, 0, '', border_left)
        ws.write(row_no, 6, '', sender_table_field_style)
        row_no += 1

        ws.write(row_no, 0, '', border_left)
        ws.write(row_no, 6, '', sender_table_field_style)
        row_no += 1

        ws.write(row_no, 0, 'TEL／FAX：', sender_table_field_style)
        ws.write_merge(row_no, row_no, 1, 5, product_sender_tel + '/' + product_sender_fax, sender_table_text_style)
        ws.write(row_no, 6, 'TEL／FAX：', sender_table_field_style)
        ws.write_merge(row_no, row_no, 7, 10, document_sender_tel + '/' + document_sender_fax, sender_table_text_style)
        row_no += 1

        # ws.write(row_no, 0, '到着予定日：', sender_table_field_style)
        # ws.write_merge(row_no, row_no, 1, 5, product_desired_arrival_date, sender_table_text_style)
        # ws.write(row_no, 6, '到着予定日：', sender_table_field_style)
        # ws.write_merge(row_no, row_no, 7, 10, document_desired_arrival_date, sender_table_text_style)
        # row_no += 1

        ws.row(row_no).height = space_height
        ws.write(row_no, 0, '', border_left)
        ws.write(row_no, 6, '', border_left)
        ws.write(row_no, 10, '', border_right)
        row_no += 1

        ws.row(row_no).height = space_height
        for i in range(0, 11):
            ws.write(row_no, i, '', border_top)
        row_no += 1

        ws.row(row_no).height = int(20 * 11.45)
        ws.write_merge(row_no, row_no, 0, 5, '【買主 署名捺印欄】', sender_table_title_style)
        ws.write_merge(row_no, row_no, 6, 10, '【売主 署名捺印欄】', sender_table_title_style)
        row_no += 1

        ws.row(row_no).height = space_height
        ws.write(row_no, 6, '', border_left)
        ws.write(row_no, 0, '', border_left)
        ws.write(row_no, 5, '', border_right)
        ws.write(row_no, 10, '', border_right)
        row_no += 1

        for i in range(0, 6):
            ws.row(row_no + i).height = int(20 * 16.5)
            ws.write(row_no + i, 0, '', border_left)
            ws.write(row_no + i, 10, '', border_right)
        
        ws.write_merge(row_no, row_no, 0, 5, '〒537-0021　大阪府大阪市東成区東中本2丁目4-15', seal_address_tel_style1)
        row_no += 1
        
        ws.write_merge(row_no, row_no + 1, 0, 5, 'バッジオ株式会社', seal_company_style)
        row_no += 2

        ws.write_merge(row_no, row_no, 0, 4, '代表取締役　金　昇志', seal_supplier_style)
        ws.write(row_no, 5, '㊞', seal_mark_style)
        ws.write(row_no, 10, '㊞', seal_mark_style)
        row_no += 1

        ws.write_merge(row_no, row_no, 0, 5, 'TEL 06-6753-8078 FAX 06-6753-8079', seal_address_tel_style)
        row_no += 1

        ws.row(row_no).height = space_height
        ws.write(row_no, 0, '', border_left)
        ws.write(row_no, 6, '', border_left)
        ws.write(row_no, 10, '', border_right)
        row_no += 1

        ws.row(row_no).height = space_height
        for i in range(0, 11):
            ws.write(row_no, i, '', border_top)
        row_no += 1

        ws.write(row_no, 5, '下記の通り御請求申し上げます。', font_11_left)
        row_no += 1

        ws.row(row_no).height = int(20 * 13.5)

        ws.write_merge(row_no, row_no, 0, 2, '運送方法', font_11_center_with_border)
        ws.write_merge(row_no, row_no + 1, 5, 7, '御請求金額', billed_deadline_label_style)
        ws.write_merge(row_no, row_no + 1, 8, 10, total, billing_amount_value_style)
        row_no += 1

        ws.row(row_no).height = int(20 * 13.5)

        shipping_method_label = '発送'
        if shipping_method == 'D': shipping_method_label = '発送'
        if shipping_method == 'R': shipping_method_label = '引取'
        if shipping_method == 'C': shipping_method_label = 'ID変更'
        if shipping_method == 'B': shipping_method_label = '* 空白'

        ws.write_merge(row_no, row_no, 0, 2, shipping_method_label, font_11_center_with_border)
        row_no += 1
        ws.row(row_no).height = int(20 * 8.25)
        row_no += 1

        ws.row(row_no).height = int(20 * 13.5)
        ws.row(row_no+1).height = int(20 * 13.5)

        ws.write_merge(row_no, row_no, 0, 2, 'お支払方法', font_11_center_with_border)
        ws.write_merge(row_no + 1, row_no + 1, 0, 2, '振込', font_11_center_with_border)

        ws.write_merge(row_no, row_no + 1, 5, 7, 'お支払期限', billed_deadline_label_style)
        ws.write_merge(row_no, row_no + 1, 8, 10, payment_due_date, billed_deadline_value_style)
        row_no += 2

        ws.row(row_no).height = space_height
        row_no += 1

        ws.row(row_no).height = int(20 * 14.1)
        ws.write(row_no, 0, '※ お振込み手数料は貴社ご負担でお願い致します。', font_11_left)
        row_no += 1

        ws.row(row_no).height = int(20 * 15)

        ws.write_merge(row_no, row_no, 0, 1, '振込先口座', transfer_account_style)
        ws.write_merge(row_no, row_no, 2, 10, '', font_11_left_with_border)
        row_no += 1

        ws.row(row_no).height = int(20 * 15)

        # ws.write(row_no, 0, '', border_left)
        # ws.write_merge(row_no, row_no, 2, 10, '', font_11_left_with_border)
        # row_no += 1

        ws.write(row_no, 0, '', border_left)
        ws.write_merge(row_no, row_no, 2, 10, '', font_11_left_with_border)
        row_no += 1

        ws.row(row_no).height = space_height
        for i in range(0, 11):
            ws.write(row_no, i, '', border_top)
        row_no += 1

        ws.row(row_no).height = int(20 * 17.25)
        ws.write(row_no, 0, '内容をご確認の上、ご捺印後弊社まで返信お願いします。', font_11_left)
        ws.write(row_no, 7, 'FAX 06-6753-8079', font_14_left1)
        row_no += 1

        ws.row(row_no).height = int(20 * 14.1)
        ws.write(row_no, 0, '※ 請求書原本が必要な場合は必ずチェックしてください。　【　必要　・　不要　】', font_11_left)
        row_no += 1

        ws.row(row_no).height = int(20 * 15)

        ws.write(row_no, 0, 'No.{}'.format(contract_id), xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: vert center, horiz left'))
        
        # -------------------------------------------------------------------------------------------------------------------------
        if total_number > 6:
            ws = wb.add_sheet('別紙', cell_overwrite_ok=True)

            ws.set_header_str(str.encode(''))
            ws.set_footer_str(str.encode(''))
            ws.set_left_margin(0.314961)
            ws.set_top_margin(0.393701)
            ws.set_right_margin(0.314961)
            ws.set_bottom_margin(0.19685)

            w_arr = [9.15, 8.04, 10.04, 2.04, 5.04, 8.04, 6.71, 7.15, 12.59, 6.71, 6.71]
            for i in range(11): 
                ws.col(i).width = int(300 * w_arr[i])
            for i in range(56):
                ws.row(i).height_mismatch = True
                ws.row(i).height = int(20 * 18)
            ws.row(44).height = int(20 * 3.75)
            ws.row(45).height = int(20 * 15.75)

            ####################################### Table #####################################
            ws.write_merge(0, 0, 0, 6, '商　品　名', TTL_CENTER)
            ws.write(0, 7, '数量', TT_CENTER)
            ws.write(0, 8, '単　価', TT_CENTER)
            ws.write_merge(0, 0, 9, 10, '金　額', TTR_CENTER)

            row_no = 1

            total_quantity = total_price = 0
            ##### Products
            if num_of_products:
                for form in product_formset.forms:
                    form.is_valid()

                    id = form.cleaned_data.get('product_id')
                    product_name = Product.objects.get(id=id).name
                    quantity = form.cleaned_data.get('quantity', 0)
                    price = form.cleaned_data.get('price', 0)
                    amount = quantity * price

                    total_quantity += quantity
                    total_price += price

                    ws.write_merge(row_no, row_no, 0, 6, product_name, TL)
                    ws.write(row_no, 7, quantity, TC_RIGHT)
                    ws.write(row_no, 8, price, TC_RIGHT)
                    ws.write_merge(row_no, row_no, 9, 10, amount, TR_RIGHT)

                    row_no += 1

            ###### Documents
            if num_of_documents:
                for form in document_formset.forms:
                    form.is_valid()

                    id = form.cleaned_data.get('document_id')
                    document_name = Document.objects.get(id=id).name.replace('（売上）', '').replace('（仕入）', '')
                    quantity = form.cleaned_data.get('quantity', 0)
                    price = form.cleaned_data.get('price', 0)
                    amount = quantity * price

                    total_quantity += quantity
                    total_price += price

                    ws.write_merge(row_no, row_no, 0, 6, document_name, TL)
                    ws.write(row_no, 7, quantity, TC_RIGHT)
                    ws.write(row_no, 8, price, TC_RIGHT)
                    ws.write_merge(row_no, row_no, 9, 10, amount, TR_RIGHT)

                    row_no += 1

            if total_number - 1 < 39:
                for i in range(0, 39 - total_number + 1):

                    ws.write_merge(row_no, row_no, 0, 6, None, TL)
                    ws.write(row_no, 7, None, TC_RIGHT)
                    ws.write(row_no, 8, None, TC_RIGHT)
                    ws.write_merge(row_no, row_no, 9, 10, None, TR_RIGHT)

                    row_no += 1

            ws.write_merge(40, 40, 0, 6, '※備考', LINE_TOP)
            ws.write_merge(40, 40, 7, 8, '小　計', TTL_CENTER)
            ws.write_merge(40, 40, 9, 10, sub_total, TTR_RIGHT)
            ws.write_merge(41, 41, 7, 8, '消費税（10%）', TL_CENTER)        
            ws.write_merge(41, 41, 9, 10, tax, TR_RIGHT)
            ws.write_merge(42, 42, 7, 8, '保険代（非課税）', TL_CENTER)
            ws.write_merge(42, 42, 9, 10, fee, TR_RIGHT)
            ws.write_merge(43, 43, 7, 8, '合　計', TBL_CENTER)
            ws.write_merge(43, 43, 9, 10, total, TBR_RIGHT_TOTAL)

            ws.write_merge(45, 45, 0, 3, 'No.{}'.format(contract_id), F11)
        
        wb.save(response)
        return response

class HallSalesInvoiceView(AdminLoginRequiredMixin, View):
    def post(self, *args, **kwargs):

        ######################################### Style Define ########################################
        w_arr = [9.15, 8.04, 10.04, 2.04, 7.04, 6.04, 6.71, 7.15, 12.59, 6.71, 6.71]

        h_arr = [5.25, 21.75, 17.25, 6, 15.75, 6, 23.25, 12.75, 12.75, 12.75, 
                12.75, 3.75, 18.75, 18.75, 18.75, 4.5, 16.5, 16.5, 4.5, 18,
                15, 15, 15, 15, 15, 15, 15, 15, 15, 15,
                15, 15, 15, 15, 15, 15, 18, 18, 18, 18,
                3.75, 18, 18, 18, 18, 18, 18, 4.5, 18, 18,
                18, 18, 3.75, 25, 3.75, 13.5, 13.5, 13.5, 13.5, 13.5]

        F11 = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: wrap on;')
        F11_BOLD = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック, bold on, underline on; align: wrap on;')
        LINE_TOP = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: wrap on; borders: top_color black, top medium;')

        TTL = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: wrap on; borders: top_color black, left_color black, bottom_color black, right_color black, top medium, left medium, right thin, bottom thin;', num_format_str='#,###')
        TTR = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: wrap on; borders: top_color black, left_color black, bottom_color black, right_color black, top medium, left thin, right medium, bottom thin;')
        TT = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: wrap on; borders: top_color black, bottom_color black, right_color black, top medium, right thin, bottom thin;')
        TC = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: wrap on; borders: bottom_color black, right_color black, right thin, bottom thin;')
        TBL = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: wrap on; borders: left_color black, bottom_color black, right_color black, left medium, right thin, bottom medium;')
        TB = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: wrap on; borders: bottom_color black, right_color black, right thin, bottom medium;')
        TBR = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: wrap on; borders: bottom_color black, right_color black, right medium, bottom medium;', num_format_str='#,###')
        TR = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: wrap on; borders: bottom_color black, right_color black, right medium, bottom thin;', num_format_str='#,###')
        TL = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: wrap on; borders: bottom_color black, left_color black, right_color black, right thin, bottom thin, left medium;')

        TTL_CENTER = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz center, vert center, wrap on; borders: top_color black, left_color black, bottom_color black, right_color black, top medium, left medium, right thin, bottom thin;')
        TTR_CENTER = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz center, vert center, wrap on; borders: top_color black, left_color black, bottom_color black, right_color black, top medium, left thin, right medium, bottom thin;')
        TT_CENTER = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz center, vert center, wrap on; borders: top_color black, bottom_color black, right_color black, top medium, right thin, bottom thin;')
        TC_CENTER = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz center, vert center, wrap on; borders: bottom_color black, right_color black, right thin, bottom thin;')
        TBL_CENTER = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz center, vert center, wrap on; borders: left_color black, bottom_color black, right_color black, left medium, right thin, bottom medium;')
        TB_CENTER = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz center, vert center, wrap on; borders: bottom_color black, right_color black, right thin, bottom medium;')
        TBR_CENTER = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz center, vert center, wrap on; borders: bottom_color black, right_color black, right medium, bottom medium;')
        TR_CENTER = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz center, vert center, wrap on; borders: bottom_color black, right_color black, right medium, bottom thin;')
        TL_CENTER = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz center, vert center, wrap on; borders: bottom_color black, left_color black, right_color black, right thin, bottom thin, left medium;')

        TTL_RIGHT = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz right, wrap on; borders: top_color black, left_color black, bottom_color black, right_color black, top medium, left medium, right thin, bottom thin;', num_format_str='#,###')
        TTR_RIGHT = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz right, wrap on; borders: top_color black, left_color black, bottom_color black, right_color black, top medium, left thin, right medium, bottom thin;', num_format_str='#,###')
        TT_RIGHT = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz right, wrap on; borders: top_color black, bottom_color black, right_color black, top medium, right thin, bottom thin;', num_format_str='#,###')
        TC_RIGHT = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz right, wrap on; borders: bottom_color black, right_color black, right thin, bottom thin;', num_format_str='#,###')
        TBL_RIGHT = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz right, wrap on; borders: left_color black, bottom_color black, right_color black, left medium, right thin, bottom medium;', num_format_str='#,###')
        TB_RIGHT = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz right, wrap on; borders: bottom_color black, right_color black, right thin, bottom medium;', num_format_str='#,###')
        TBR_RIGHT = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz right, wrap on; borders: bottom_color black, right_color black, right medium, bottom medium;', num_format_str='#,###')
        TR_RIGHT = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz right, wrap on; borders: bottom_color black, right_color black, right medium, bottom thin;', num_format_str='#,###')
        TL_RIGHT = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz right, wrap on; borders: bottom_color black, left_color black, right_color black, right thin, bottom thin, left medium;', num_format_str='#,###')

        TBR_RIGHT_TOTAL = xlwt.easyxf('font: height 240, name ＭＳ Ｐゴシック, bold on; align: horiz right, vert center, wrap on; borders: bottom_color black, right_color black, right medium, bottom medium;', num_format_str='"¥"#,###')

        TTL_THIN = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: wrap on; borders: top_color black, left_color black, bottom_color black, right_color black, top thin, left thin, right thin, bottom thin;')
        TTR_THIN = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: wrap on; borders: top_color black, left_color black, bottom_color black, right_color black, top thin, left thin, right thin, bottom thin;', num_format_str='#,###')
        TT_THIN = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: wrap on; borders: top_color black, bottom_color black, right_color black, top thin, right thin, bottom thin;')
        TC_THIN = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: wrap on; borders: bottom_color black, right_color black, right thin, bottom thin;')
        TBL_THIN = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: wrap on; borders: left_color black, bottom_color black, right_color black, left thin, right thin, bottom thin;')
        TB_THIN = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: wrap on; borders: bottom_color black, right_color black, right thin, bottom thin;')
        TBR_THIN = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: wrap on; borders: bottom_color black, right_color black, right thin, bottom thin;', num_format_str='#,###')
        TR_THIN = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: wrap on; borders: bottom_color black, right_color black, right thin, bottom thin;', num_format_str='#,###')
        TL_THIN = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: wrap on; borders: bottom_color black, left_color black, right_color black, right thin, bottom thin, left thin;')

        TTL_THIN_CENTER = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz center, vert center, wrap on; borders: top_color black, left_color black, bottom_color black, right_color black, top thin, left thin, right thin, bottom thin;')
        TTR_THIN_CENTER = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz center, vert center, wrap on; borders: top_color black, left_color black, bottom_color black, right_color black, top thin, left thin, right thin, bottom thin;', num_format_str='#,###')
        TT_THIN_CENTER = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz center, vert center, wrap on; borders: top_color black, bottom_color black, right_color black, top thin, right thin, bottom thin;')
        TC_THIN_CENTER = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz center, vert center, wrap on; borders: bottom_color black, right_color black, right thin, bottom thin;')
        TBL_THIN_CENTER = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz center, vert center, wrap on; borders: left_color black, bottom_color black, right_color black, left thin, right thin, bottom thin;')
        TB_THIN_CENTER = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz center, vert center, wrap on; borders: bottom_color black, right_color black, right thin, bottom thin;')
        TBR_THIN_CENTER = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz center, vert center, wrap on; borders: bottom_color black, right_color black, right thin, bottom thin;', num_format_str='#,###')
        TR_THIN_CENTER = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz center, vert center, wrap on; borders: bottom_color black, right_color black, right thin, bottom thin;')
        TL_THIN_CENTER = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz center, vert center, wrap on; borders: bottom_color black, left_color black, right_color black, right thin, bottom thin, left thin;')

        ################################### Data Init ###################################### 
        user_id = self.request.user.id
        log_export_operation(user_id, "{} - {}".format(_("Sales contract"), _("Hall sales")))
        contract_form = HallSalesContractForm(self.request.POST)
        contract_id = contract_form.data.get('contract_id', '')

        customer_id = contract_form.data.get('customer_id')
        created_at = contract_form.data.get('created_at', '')
        manager = contract_form.data.get('manager')
        hall_id = contract_form.data.get('hall_id')
        update_hall_id = contract_form.data.get('update_hall_id')

        if not hall_id and update_hall_id:
            hall_id = update_hall_id

        company = address = tel = fax = ''
        if customer_id:
            customer = Hall.objects.get(id=customer_id)
            company = customer.customer_name
            address = customer.address
            tel = customer.tel
            fax = customer.fax
        hall_name = hall_address = hall_tel = ''
        if hall_id:
            hall = Hall.objects.get(id=hall_id)
            company = hall.customer_name
            tel = hall.tel
            fax = hall.fax
            hall_name = hall.name
            address = hall_address = hall.address
            hall_tel = hall.tel

        product_formset = ProductFormSet(
            self.request.POST,
            prefix='product'
        )
        total_quantity = total_price = 0
        num_of_products = product_formset.total_form_count()

        document_formset = DocumentFormSet(
            self.request.POST,
            prefix='document'
        )
        num_of_documents = document_formset.total_form_count()

        document_fee_formset = DocumentFeeFormSet(
            self.request.POST,
            prefix='document_fee'
        )
        num_of_document_fees = document_fee_formset.total_form_count()

        remarks = contract_form.data.get('remarks')
        sub_total = contract_form.data.get('sub_total')
        tax = contract_form.data.get('tax')
        fee = contract_form.data.get('fee')
        total = contract_form.data.get('total')
        shipping_date = contract_form.data.get('shipping_date')
        opening_date = contract_form.data.get('opening_date')
        payment_method = contract_form.data.get('payment_method')
        if payment_method == 'TR': payment_method = '振込'
        elif payment_method == 'CH': payment_method = '小切手'
        elif payment_method == 'BL': payment_method = '手形'
        elif payment_method == 'CA': payment_method = '現金'
        transfer_account = contract_form.data.get('transfer_account')
        person_in_charge = contract_form.data.get('person_in_charge', '')
        confirmor = contract_form.data.get('confirmor')

        total_number = num_of_products + num_of_documents + num_of_document_fees

        try:
            sub_total = int(sub_total.replace(',', '')) if len(sub_total.replace(',', '')) > 0 else 0
            tax = int(tax.replace(',', '')) if len(tax.replace(',', '')) > 0 else 0
            fee = int(fee.replace(',', '')) if len(fee.replace(',', '')) > 0 else 0
            total = int(total.replace(',', '')) if len(total.replace(',', '')) > 0 else 0

        except ValueError:
            sub_total = 0
            tax = 0
            fee = 0
            total = 0

        ######################################### Sheet Init ########################################

        response = HttpResponse(content_type='application/ms-excel')
        response['Content-Disposition'] = 'attachment; filename="hall_sales_contract_{}.xls"'.format(contract_id)
        # response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        # response['Content-Disposition'] = 'attachment; filename="hall_sales_contract_{}.xlsx"'.format(contract_id)
        wb = xlwt.Workbook(encoding='utf-8')
        ws = wb.add_sheet("請求書", cell_overwrite_ok=True)

        ws.set_header_str(str.encode(''))
        ws.set_footer_str(str.encode(''))
        ws.set_left_margin(0.314961)
        ws.set_top_margin(0.393701)
        ws.set_right_margin(0.314961)
        ws.set_bottom_margin(0.19685)

        for i in range(11): 
            ws.col(i).width = int(300 * w_arr[i])

        ######################################### Start Drawing ########################################
        total_items_count = num_of_products + num_of_documents + num_of_document_fees
        
        for i in range(56):
            ws.row(i).height_mismatch = True
            ws.row(i).height = int(20 * h_arr[i])

        date_str = ''
        created_at = getDate(created_at)

        if created_at:
            date_str = '%d年%d月%d日' % (created_at.year, created_at.month, created_at.day)

        ws.write_merge(1, 1, 0, 10, '売買契約書', xlwt.easyxf('font: height 400, name ＭＳ Ｐゴシック, color black; align: horiz center, wrap on;'))
        # ws.write_merge(2, 2, 3, 7, '（お客様控え）', xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック, color black; align: horiz center, wrap on;'))
        ws.write_merge(2, 2, 8, 10, date_str, xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック, color black; align: horiz right, wrap on;'))
        ws.write(4, 0, '買主（甲）', F11)
        ws.write_merge(4, 4, 5, 7, '売主（乙）', F11)
        ws.write_merge(6, 6, 5, 10, 'バッジオ株式会社', xlwt.easyxf('font: height 320, name ＭＳ Ｐゴシック, color black, bold on; align: horiz left, vert center, wrap on;'))
        ws.write_merge(7, 7, 5, 10, '代表取締役　金　昇志', F11)
        ws.write_merge(8, 8, 5, 10, '〒537-0021　大阪府大阪市東成区東中本2丁目4-15', F11)
        ws.write_merge(9, 9, 5, 9, 'TEL: 06-6753-8078', F11)
        ws.write(10, 4, '㊞', F11)
        ws.write_merge(10, 10, 5, 9, 'FAX: 06-6753-8079', F11)
        ws.write(10, 10, '㊞', F11)

        ws.write(12, 0, '設置場所', F11)
        ws.write(12, 1, 'ホール名', F11)
        ws.write_merge(12, 12, 2, 10, hall_name)
        ws.write(13, 1, '住所', F11)
        ws.write_merge(13, 13, 2, 10, hall_address)
        ws.write(14, 1, 'TEL', F11)
        ws.write_merge(14, 14, 2, 10, hall_tel)

        ws.write_merge(16, 17, 0, 10, '買主（以下甲という）と売主（以下乙という）は下記の商品（以下物件という）を裏面記載取引約定を含み売買契約を締結致しました。', F11)

        ####################################### Table #####################################
        ws.write_merge(19, 19, 0, 6, '商　品　名', TTL_CENTER)
        ws.write(19, 7, '数量', TT_CENTER)
        ws.write(19, 8, '単　価', TT_CENTER)
        ws.write_merge(19, 19, 9, 10, '金　額', TTR_CENTER)

        row_no = 20

        ##### Products
        if num_of_products:
            for form in product_formset.forms:
                form.is_valid()

                id = form.cleaned_data.get('product_id')
                product_name = Product.objects.get(id=id).name
                quantity = form.cleaned_data.get('quantity', 0)
                price = form.cleaned_data.get('price', 0)
                amount = quantity * price

                total_quantity += quantity
                total_price += price

                if total_number > 16: 
                    continue

                ws.write_merge(row_no, row_no, 0, 6, product_name, TL)
                ws.write(row_no, 7, quantity, TC_RIGHT)
                ws.write(row_no, 8, price, TC_RIGHT)
                ws.write_merge(row_no, row_no, 9, 10, amount, TR_RIGHT)

                row_no += 1

        ###### Documents
        if num_of_documents:
            for form in document_formset.forms:
                form.is_valid()

                id = form.cleaned_data.get('document_id')
                document_name = Document.objects.get(id=id).name.replace('（売上）', '').replace('（仕入）', '')
                quantity = form.cleaned_data.get('quantity', 0)
                price = form.cleaned_data.get('price', 0)
                amount = quantity * price

                total_quantity += quantity
                total_price += price

                if total_number > 16: 
                    continue

                ws.write_merge(row_no, row_no, 0, 6, document_name, TL)
                ws.write(row_no, 7, quantity, TC_RIGHT)
                ws.write(row_no, 8, price, TC_RIGHT)
                ws.write_merge(row_no, row_no, 9, 10, amount, TR_RIGHT)

                row_no += 1

        ####### Document Fee
        if num_of_document_fees:
            for form in document_fee_formset.forms:
                form.is_valid()

                id = form.cleaned_data.get('document_fee_id')
                document_fee = DocumentFee.objects.get(id=id)
                type = document_fee.type
                model_count = form.cleaned_data.get('model_count', 0)
                unit_count = form.cleaned_data.get('unit_count', 0)
                model_price = document_fee.model_price
                unit_price = document_fee.unit_price
                amount = model_price * model_count + unit_price * unit_count + document_fee.application_fee

                total_quantity += model_count
                total_price += unit_count

                if total_number > 16: 
                    continue

                ws.write_merge(row_no, row_no, 0, 6, str(dict(TYPE_CHOICES)[type]) + ' ' + str(model_count) + '機種' + str(unit_count) + '台', TL)
                ws.write(row_no, 7, None, TC_RIGHT)
                ws.write(row_no, 8, None, TC_RIGHT)
                ws.write_merge(row_no, row_no, 9, 10, amount, TR_RIGHT)

                row_no += 1

        forRange = 16 - total_number if total_number <= 16 else 16
        for i in range(0, forRange):

            ws.write_merge(row_no, row_no, 0, 6, None, TL)
            ws.write(row_no, 7, None, TC_RIGHT)
            ws.write(row_no, 8, None, TC_RIGHT)
            ws.write_merge(row_no, row_no, 9, 10, None, TR_RIGHT)

            row_no += 1

        ws.write_merge(36, 39, 0, 6, '※備考: ' + remarks, xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: vert top, wrap on; borders: top_color black, top medium, left_color black, left thin, bottom_color black, bottom thin;'))
        ws.write_merge(36, 36, 7, 8, '小　計', TTL_CENTER)
        ws.write_merge(36, 36, 9, 10, sub_total, TTR_RIGHT)
        ws.write_merge(37, 37, 7, 8, '消費税（10%）', TL_CENTER)        
        ws.write_merge(37, 37, 9, 10, tax, TR_RIGHT)
        ws.write_merge(38, 38, 7, 8, '保険代（非課税）', TL_CENTER)
        ws.write_merge(38, 38, 9, 10, fee, TR_RIGHT)
        ws.write_merge(39, 39, 7, 8, '合　計', TBL_CENTER)
        ws.write_merge(39, 39, 9, 10, total, TBR_RIGHT_TOTAL)

        ########################################### Below tables ###################################
        ws.write_merge(42, 42, 0, 2, '納品日', TTL_THIN_CENTER)

        shipping_date = getDate(shipping_date)
        date_str = ''
        if shipping_date: date_str = '%d年%d月%d日' % (shipping_date.year, shipping_date.month, shipping_date.day)

        ws.write_merge(43, 43, 0, 2, date_str, TBL_THIN_CENTER)

        ws.write_merge(45, 45, 0, 2, 'お支払方法', TTL_THIN_CENTER)
        ws.write_merge(46, 46, 0, 2, payment_method, TBL_THIN_CENTER)

        ws.write_merge(41, 41, 4, 7, '下記の通り御請求申し上げます。', F11)
        ws.write_merge(42, 46, 4, 5, 'お支払内訳', TTL_CENTER)

        milestone_formset = MilestoneFormSet(
            self.request.POST,
            prefix='milestone'
        )

        ind = 0
        for form in milestone_formset.forms:
            form.is_valid()
            date = str(form.cleaned_data.get('date'))

            date = getDate(date)

            date_str = ''
            if date: date_str = '%d年%d月%d日' % (date.year, date.month, date.day)

            amount = form.cleaned_data.get('amount')

            ws.row(row_no).height = int(20 * 18)
            if ind == 0:
                ws.write(42, 6, '初回', xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz center, vert center, wrap on; borders: top_color black, bottom_color black, right_color black, top medium, right thin, bottom thin;'))
                ws.write_merge(42, 42, 7, 8, date_str, xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック, bold on; align: horiz center, vert center, wrap on; borders: top_color black, bottom_color black, right_color black, top medium, right thin, bottom thin;'))
                ws.write_merge(42, 42, 9, 10, amount, xlwt.easyxf('font: height 240, name ＭＳ Ｐゴシック, bold on; align: horiz right, vert center, wrap on; borders: top_color black, left_color black, bottom_color black, right_color black, top medium, left thin, right medium, bottom thin;', num_format_str='"¥"#,###'))
            else:
                ws.write(42 + ind, 6, str(ind + 1) + '回', xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz center, vert center, wrap on; borders: bottom_color black, right_color black, right thin, bottom thin;'))
                ws.write_merge(42 + ind, 42 + ind, 7, 8, date_str, xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック, bold on; align: horiz center, vert center, wrap on; borders: bottom_color black, right_color black, right thin, bottom thin;', num_format_str='"¥"#,###'))
                ws.write_merge(42 + ind, 42 + ind, 9, 10, amount, xlwt.easyxf('font: height 240, name ＭＳ Ｐゴシック, bold on; align: horiz right, vert center, wrap on; borders: top_color black, left_color black, bottom_color black, right_color black, right medium, bottom thin;', num_format_str='"¥"#,###'))
            
            ind += 1

        ws.write_merge(47, 47, 4, 10, None, LINE_TOP)

        ws.write_merge(48, 48, 0, 5, '※お振込み手数料は貴社ご負担でお願い致します。', F11)

        ################################# Last Table ################################
        ws.write_merge(49, 51, 0, 1, '振込先口座', TTL_THIN_CENTER)
        ws.write_merge(49, 49, 2, 10, 'りそな銀行　船場支店（101）　普通　0530713　バッジオカブシキガイシャ', TTR_THIN)
        ws.write_merge(50, 50, 2, 10, None, TR_THIN)
        ws.write_merge(51, 51, 2, 10, None, TR_THIN)

        ws.write(53, 7, '担当者', TTL_THIN_CENTER)
        ws.write(53, 8, person_in_charge, TT_THIN_CENTER)
        ws.write(53, 9, '確認印', TT_THIN_CENTER)
        ws.write(53, 10, confirmor, TTR_THIN_CENTER)

        ws.write_merge(53, 53, 0, 1, 'No.{}'.format(contract_id), xlwt.easyxf('font: height 240, name ＭＳ Ｐゴシック; align: horiz left, vert center, wrap on;'))


        ################################# Next Page ################################
        for i in range(56,118):
            ws.row(i).height_mismatch = True
            ws.row(i).height = int(20 * 13.5)

        INDENT = "";
        INDENT2 = "   "
        INDENT3 = "       "
        ws.write_merge(59, 59, 1, 2, "取引約定", F11_BOLD)
        ws.write_merge(61, 61, 0, 10, INDENT + "1. (1)物件の所有権は代金完済(売主が受領した小切手又は手形が期日到来し決済されるまで)するまで", F11)
        ws.write_merge(62, 62, 0, 10, INDENT3 + "乙に属し、完済と同時に甲に移転する。", F11)
        ws.write_merge(63, 63, 0, 10, INDENT2 + "(2)甲は、物件の納品を受けたと同時に、乙に対し、占有改定の方法により物件を引き渡し、これにより、", F11)
        ws.write_merge(64, 64, 0, 10, INDENT3 + "甲及び乙は、物件の占有が乙にあることを相互に確認する。", F11)
        ws.write_merge(65, 65, 0, 10, INDENT2 + "(3)甲が上記 (1) に基づいて物件の所有権を取得したと同時に、乙は甲に対し、物件を引き渡すもの", F11)
        ws.write_merge(66, 66, 0, 10, INDENT3 + "とする。", F11)
        ws.write_merge(67, 67, 0, 10, INDENT2 + "(4)甲は、上記(1)に基づいて物件の所有権が甲に移転するまで、善良なる管理者の注意をもって、乙のた", F11)
        ws.write_merge(68, 68, 0, 10, INDENT3 + "めに物件を保管する。", F11)
        ws.write_merge(69, 69, 0, 10, INDENT2 + "(5) 甲は、乙より別段の指示がない限り通常の営業の範囲において、物件を使用することができる。", F11)
        ws.write_merge(70, 70, 0, 10, INDENT2 + "(6) 甲は、代金を完済するまでは、下記各号の行為をしてはならない。", F11)
        ws.write_merge(71, 71, 0, 10, INDENT3 + "①乙に無断で物件を表面記載以外の設置場所に移転すること。", F11)
        ws.write_merge(72, 72, 0, 10, INDENT3 + "②物件を第三者に貸与または譲渡すること。", F11)
        ws.write_merge(73, 73, 0, 10, INDENT3 + "③物件を譲渡担保に供すること。", F11)

        ws.write_merge(74, 74, 0, 10, INDENT + "2. 甲は本契約の締結と同時に手付金を乙に支払うものとする。", F11)
        ws.write_merge(75, 75, 0, 10, INDENT + "3. 甲は物件の引渡を受けた後7日以内にその検査を行い、検査後は物件の引渡の遅延、瑕疵、故障、付属品", F11)
        ws.write_merge(76, 76, 0, 10, INDENT2 + "この不足その他乙の側における契約の履行につき何ら請求、異議を申し立てることはできない。", F11)
        ws.write_merge(77, 77, 0, 10, INDENT + "4. 天災、人災等による物件の毀損、滅失、紛失、盗難、被詐取その一切の損害については物件の引渡を受け", F11)
        ws.write_merge(78, 78, 0, 10, INDENT2 + "た時からすべて甲において負担するものとする。", F11)
        ws.write_merge(79, 79, 0, 10, INDENT + "5. 物件の引渡後は同物件に附される損害保険に関し甲が保険料の支払に任じ代金完済前の事故発生により受", F11)
        ws.write_merge(80, 80, 0, 10, INDENT2 + "領した保険料は未済代金の限度まで乙に帰属するものとする。", F11)
        ws.write_merge(81, 81, 0, 10, INDENT + "6. 甲または甲の保証人が第三者より差押、仮差押、仮処分の実行を受けまたは受けるおそれのある時、ある", F11)
        ws.write_merge(82, 82, 0, 10, INDENT2 + "いは甲が一回でも所定代金支払を延滞した時、甲ならびに甲の保証人は当然に期限の利益を失い、未払代金", F11)
        ws.write_merge(83, 83, 0, 10, INDENT2 + "を一時に乙に支払わなければならない。", F11)
        ws.write_merge(84, 84, 0, 10, INDENT + "7. 本件債務不履行による期限後の延滞損害金は利息期限法所定最高額の率とする。", F11)
        ws.write_merge(85, 85, 0, 10, INDENT + "8. 甲が残金の支払を一回でもなった時、その他本契約の条項の一つでも違反した時、乙は何等の催告を要せ", F11)
        ws.write_merge(86, 86, 0, 10, INDENT2 + "ず本契約を直ちに解除することができる。", F11)
        ws.write_merge(87, 87, 0, 10, INDENT + "9. 乙より本契約を解除された時、甲は物件の使用権を失い、直ちに乙に対し、物件を乙の指定する場所におい", F11)
        ws.write_merge(88, 88, 0, 10, INDENT2 + "て返還しなければならない。甲がこの返還をなさない場合は乙または乙の代理人は予告なく同物件の所在場所 ", F11)
        ws.write_merge(89, 89, 0, 10, INDENT2 + "に赴き、同物件を回収し搬出することができる。甲はこれに対して何等の異議を申し立てることなく、", F11)
        ws.write_merge(90, 90, 0, 10, INDENT2 + "損害賠償等の要求をすることができない。", F11)

        # 10
        # ws.write_merge(62, 62, 0, 9, "", F11)
        ws.write_merge(91, 91, 0, 10, INDENT + "10. 本契約を解除した時は、乙は甲より受領済みの代金を物件の使用料および損害金の一部または全部に充当し", F11)
        ws.write_merge(92, 92, 0, 10, INDENT2 + "物件の使用料および損害金が受領済の代金を超える時は乙は甲に対しその差額を請求することができる。 ", F11)
        ws.write_merge(93, 93, 0, 10, INDENT2 + "尚、使用料は物件の引渡を受けた日から返還の日まで、最初の60日までについては1日1台につき物件単 ", F11)
        ws.write_merge(94, 94, 0, 10, INDENT2 + "価の75分の1,61日以降については同じく100分の1の割合で計算した金額とする。但し使用料は物", F11)
        ws.write_merge(95, 95, 0, 10, INDENT2 + "件の代金総額の範囲内とする。", F11)

        ws.write_merge(96, 96, 0, 10, INDENT + "11.甲は乙の請求ある場合は本契約につき執行承諾文書を付した公正証書の作成手続に異議なく協力するものと", F11)
        ws.write_merge(97, 97, 0, 10, INDENT2 + "する。", F11)
        ws.write_merge(98, 98, 0, 10, INDENT + "12. 連帯保証人は甲と連帯して本契約について甲が乙に対し負担する一切の債務を保証する。", F11)
        ws.write_merge(99, 99, 0, 10, INDENT + "13. 行政措置に関連して機械構造基準の変更ある場合は、甲はこの契約に、一切の不服を申し立てない。", F11)
        ws.write_merge(100, 100, 0, 10, INDENT + "14. 本契約に関し将来当事者間に於いて紛争が生じたときは、売主の本店所在地を管轄する管轄裁判所とする", F11)
        ws.write_merge(101, 101, 0, 10, INDENT2 + "ことを双方合意する。", F11)
        ws.write_merge(102, 102, 0, 10, INDENT + "15. 本契約に定めのない事項または疑義が生じた場合は甲乙双方誠意をもって定めるものとする。", F11)
        ws.write_merge(104, 104, 0, 10, INDENT2 + "以上契約の証として本証2通を作成し、甲乙各1通を所持する。", F11)
        ws.write(106, 8, "以   上", F11)


        ######################################################################################################################################
        if total_items_count > 16: 
            ws = wb.add_sheet('別紙', cell_overwrite_ok=True)

            for i in range(11): 
                ws.col(i).width = int(300 * w_arr[i])

            ws.set_header_str(str.encode(''))
            ws.set_footer_str(str.encode(''))
            ws.set_left_margin(0.314961)
            ws.set_top_margin(0.393701)
            ws.set_right_margin(0.314961)
            ws.set_bottom_margin(0.19685)

            for i in range(56):
                ws.row(i).height_mismatch = True
                ws.row(i).height = int(20 * 18)
            ws.row(44).height = int(20 * 3.75)
            ws.row(45).height = int(20 * 15.75)

            ####################################### Table #####################################
            ws.write_merge(0, 0, 0, 6, '商　品　名', TTL_CENTER)
            ws.write(0, 7, '数量', TT_CENTER)
            ws.write(0, 8, '単　価', TT_CENTER)
            ws.write_merge(0, 0, 9, 10, '金　額', TTR_CENTER)

            row_no = 1

            ##### Products
            if num_of_products:
                for form in product_formset.forms:
                    form.is_valid()

                    id = form.cleaned_data.get('product_id')
                    product_name = Product.objects.get(id=id).name
                    quantity = form.cleaned_data.get('quantity', 0)
                    price = form.cleaned_data.get('price', 0)
                    amount = quantity * price

                    total_quantity += quantity
                    total_price += price

                    ws.write_merge(row_no, row_no, 0, 6, product_name, TL)
                    ws.write(row_no, 7, quantity, TC_RIGHT)
                    ws.write(row_no, 8, price, TC_RIGHT)
                    ws.write_merge(row_no, row_no, 9, 10, amount, TR_RIGHT)

                    row_no += 1

            ###### Documents
            if num_of_documents:
                for form in document_formset.forms:
                    form.is_valid()

                    id = form.cleaned_data.get('document_id')
                    document_name = Document.objects.get(id=id).name.replace('（売上）', '').replace('（仕入）', '')
                    quantity = form.cleaned_data.get('quantity', 0)
                    price = form.cleaned_data.get('price', 0)
                    amount = quantity * price

                    total_quantity += quantity
                    total_price += price

                    ws.write_merge(row_no, row_no, 0, 6, document_name, TL)
                    ws.write(row_no, 7, quantity, TC_RIGHT)
                    ws.write(row_no, 8, price, TC_RIGHT)
                    ws.write_merge(row_no, row_no, 9, 10, amount, TR_RIGHT)

                    row_no += 1

            ####### Document Fee
            if num_of_document_fees:
                for form in document_fee_formset.forms:
                    form.is_valid()

                    id = form.cleaned_data.get('document_fee_id')
                    document_fee = DocumentFee.objects.get(id=id)
                    type = document_fee.type
                    model_count = form.cleaned_data.get('model_count', 0)
                    unit_count = form.cleaned_data.get('unit_count', 0)
                    model_price = document_fee.model_price
                    unit_price = document_fee.unit_price
                    amount = model_price * model_count + unit_price * unit_count + document_fee.application_fee

                    total_quantity += model_count
                    total_price += unit_count

                    ws.write_merge(row_no, row_no, 0, 6, str(dict(TYPE_CHOICES)[type]) + ' ' + str(model_count) + '機種' + str(unit_count) + '台', TL)
                    ws.write(row_no, 7, None, TC_RIGHT)
                    ws.write(row_no, 8, None, TC_RIGHT)
                    ws.write_merge(row_no, row_no, 9, 10, amount, TR_RIGHT)

                    row_no += 1

            total_number = num_of_products + num_of_documents + num_of_document_fees
            if total_number - 1 < 39:
                for i in range(0, 39 - total_number + 1):

                    ws.write_merge(row_no, row_no, 0, 6, None, TL)
                    ws.write(row_no, 7, None, TC_RIGHT)
                    ws.write(row_no, 8, None, TC_RIGHT)
                    ws.write_merge(row_no, row_no, 9, 10, None, TR_RIGHT)

                    row_no += 1

            ws.write_merge(40, 40, 0, 6, '※備考', LINE_TOP)
            ws.write_merge(40, 40, 7, 8, '小　計', TTL_CENTER)
            ws.write_merge(40, 40, 9, 10, sub_total, TTR_RIGHT)
            ws.write_merge(41, 41, 7, 8, '消費税（10%）', TL_CENTER)        
            ws.write_merge(41, 41, 9, 10, tax, TR_RIGHT)
            ws.write_merge(42, 42, 7, 8, '保険代（非課税）', TL_CENTER)
            ws.write_merge(42, 42, 9, 10, fee, TR_RIGHT)
            ws.write_merge(43, 43, 7, 8, '合　計', TBL_CENTER)
            ws.write_merge(43, 43, 9, 10, total, TBR_RIGHT_TOTAL)

            ws.write_merge(45, 45, 0, 3, 'No.{}'.format(contract_id), F11)

        wb.save(response)
        return response

class HallPurchasesInvoiceView(AdminLoginRequiredMixin, View):
    def post(self, *args, **kwargs):

        user_id = self.request.user.id
        log_export_operation(user_id, "{} - {}".format(_("Purchase contract"), _("Hall purchase")))
        contract_form = HallPurchasesContractForm(self.request.POST)
        contract_id = contract_form.data.get('contract_id', '')

        customer_id = contract_form.data.get('customer_id')
        created_at = contract_form.data.get('created_at', '')
        manager = contract_form.data.get('manager')
        hall_id = contract_form.data.get('hall_id')
        update_hall_id = contract_form.data.get('update_hall_id')

        if not hall_id and update_hall_id:
            hall_id = update_hall_id
            
        company = address = tel = fax = ''
        if customer_id:
            customer = Hall.objects.get(id=customer_id)
            company = customer.customer_name
            address = customer.address
            tel = customer.tel
            fax = customer.fax
        hall_name = hall_address = hall_tel = ''
        if hall_id:
            hall = Hall.objects.get(id=hall_id)
            company = hall.customer_name
            tel = hall.tel
            fax = hall.fax
            hall_name = hall.name
            address = hall_address = hall.address
            hall_tel = hall.tel

        response = HttpResponse(content_type='application/ms-excel')
        response['Content-Disposition'] = 'attachment; filename="hall_purchases_contract_{}.xls"'.format(contract_id)
        # response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        # response['Content-Disposition'] = 'attachment; filename="hall_purchases_contract_{}.xlsx"'.format(contract_id)
        wb = xlwt.Workbook(encoding='utf-8')
        ws = wb.add_sheet("請求書", cell_overwrite_ok=True)

        ws.set_header_str(str.encode(''))
        ws.set_footer_str(str.encode(''))
        ws.set_left_margin(0.314961)
        ws.set_top_margin(0.393701)
        ws.set_right_margin(0.314961)
        ws.set_bottom_margin(0.19685)

        ######################################### Drawing ########################################
        c_w = int(256 * 3.6)
        cell_height = int(20 * 23)

        s_dot_rb = xlwt.easyxf('font: height 200, name ＭＳ 明朝; align: vert center, horiz center, wrap on; borders: top_color black, left_color black, bottom_color black, right_color black, bottom dotted, right dotted;')
        s_dot_rb_num = xlwt.easyxf('font: height 200, name ＭＳ 明朝; align: vert center, horiz center, wrap on; borders: top_color black, left_color black, bottom_color black, right_color black, bottom dotted, right dotted;', num_format_str='#,##0')
        s_dot_b = xlwt.easyxf('font: height 200, name ＭＳ 明朝; align: vert center, horiz center, wrap on; borders: bottom_color black, bottom dotted')
        s_thin_r = xlwt.easyxf('font: height 200, name ＭＳ 明朝; align: vert center, horiz center, wrap on; borders: right_color black, right thin')
        s_dot_t = xlwt.easyxf('font: height 200, name ＭＳ 明朝; align: vert center, horiz center, wrap on; borders: top_color black, top dotted')
        s_none = xlwt.easyxf('font: height 200, name ＭＳ 明朝; align: vert center, horiz center, wrap on;')
        s_dot_b_thin_r = xlwt.easyxf('font: height 200, name ＭＳ 明朝; align: vert center, horiz center, wrap on; borders: top_color black, left_color black, bottom_color black, right_color black, bottom dotted, right thin;')
        s_dot_tb_thin_r = xlwt.easyxf('font: height 200, name ＭＳ 明朝; align: vert center, horiz center, wrap on; borders: top_color black, left_color black, bottom_color black, right_color black, top dotted, bottom dotted, right thin;')
        s_dot_rtb = xlwt.easyxf('font: height 200, name ＭＳ 明朝; align: vert center, horiz center, wrap on; borders: top_color black, left_color black, bottom_color black, right_color black, top dotted, bottom dotted, right dotted;')
        s_dot_tb = xlwt.easyxf('font: height 200, name ＭＳ 明朝; align: vert center, horiz center, wrap on; borders: top_color black, bottom_color black, top dotted, bottom dotted')

        TTL_CENTER = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz center, vert center, wrap on; borders: top_color black, left_color black, bottom_color black, right_color black, top medium, left medium, right thin, bottom thin;')
        TT_CENTER = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz center, vert center, wrap on; borders: top_color black, bottom_color black, right_color black, top medium, right thin, bottom thin;')
        TTR_CENTER = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz center, vert center, wrap on; borders: top_color black, left_color black, bottom_color black, right_color black, top medium, left thin, right medium, bottom thin;')
        TC_RIGHT = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz right, wrap on; borders: bottom_color black, right_color black, right thin, bottom thin;', num_format_str='#,###')
        TR_RIGHT = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz right, wrap on; borders: bottom_color black, right_color black, right medium, bottom thin;', num_format_str='#,###')
        F11 = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: wrap on;')
        LINE_TOP = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: wrap on; borders: top_color black, top medium;')
        TTR_RIGHT = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz right, wrap on; borders: top_color black, left_color black, bottom_color black, right_color black, top medium, left thin, right medium, bottom thin;', num_format_str='#,###')
        TL_CENTER = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz center, vert center, wrap on; borders: bottom_color black, left_color black, right_color black, right thin, bottom thin, left medium;')
        TBL_CENTER = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: horiz center, vert center, wrap on; borders: left_color black, bottom_color black, right_color black, left medium, right thin, bottom medium;')
        TBR_RIGHT_TOTAL = xlwt.easyxf('font: height 240, name ＭＳ Ｐゴシック, bold on; align: horiz right, vert center, wrap on; borders: bottom_color black, right_color black, right medium, bottom medium;', num_format_str='"¥"#,###')
        TL = xlwt.easyxf('font: height 220, name ＭＳ Ｐゴシック; align: wrap on; borders: bottom_color black, left_color black, right_color black, right thin, bottom thin, left medium;')

        for i in range(1, 80):
            ws.row(i).height_mismatch = True
            ws.row(i).height = cell_height

        product_formset = ProductFormSet(
            self.request.POST,
            prefix='product'
        )

        total_quantity = total_price = 0

        num_of_products = product_formset.total_form_count()

        document_formset = DocumentFormSet(
            self.request.POST,
            prefix='document'
        )
        num_of_documents = document_formset.total_form_count()

        document_fee_formset = DocumentFeeFormSet(
            self.request.POST,
            prefix='document_fee'
        )
        num_of_document_fees = document_fee_formset.total_form_count()

        total_number = num_of_products + num_of_documents + num_of_document_fees

        remarks = contract_form.data.get('remarks')
        sub_total = contract_form.data.get('sub_total')
        tax = contract_form.data.get('tax')
        fee = contract_form.data.get('fee')
        total = contract_form.data.get('total')
        shipping_date = contract_form.data.get('shipping_date')
        opening_date = contract_form.data.get('opening_date')
        payment_method = contract_form.data.get('payment_method')
        transfer_account = contract_form.data.get('transfer_account')
        person_in_charge = contract_form.data.get('person_in_charge', '')
        confirmor = contract_form.data.get('confirmor')

        try:
            sub_total = int(sub_total.replace(',', '')) if len(sub_total.replace(',', '')) > 0 else 0
            tax = int(tax.replace(',', '')) if len(tax.replace(',', '')) > 0 else 0
            fee = int(fee.replace(',', '')) if len(fee.replace(',', '')) > 0 else 0
            total = int(total.replace(',', '')) if len(total.replace(',', '')) > 0 else 0

        except ValueError:
            sub_total = 0
            tax = 0
            fee = 0
            total = 0

        
        
        for i in range(25):
            ws.col(i).width = c_w
        ws.col(4).width = ws.col(14).width = int(256 * 5.5)

        row_no = 0

        ws.row(row_no).height_mismatch = True
        ws.row(row_no).height = int(20 * 39)

        ws.write_merge(row_no, row_no, 4, 20, '売買契約書', xlwt.easyxf('font: height 400, name ＭＳ 明朝, color black, bold on; align: vert center, horiz center, wrap on; borders: bottom_color black, bottom double;'))
        row_no += 1
        ws.row(row_no).height = int(20 * 15)
        ws.write_merge(row_no, row_no, 23, 24, 'No', xlwt.easyxf('font: height 200, name ＭＳ 明朝; align: vert center, horiz left; borders: bottom_color black, bottom dotted'))
        
        row_no += 1
        ws.row(row_no).height = int(20 * 15)
        row_no += 1
        ws.write_merge(row_no, row_no, 0, 6, '会　　社　　名', xlwt.easyxf('font: height 200, name ＭＳ 明朝; align: vert center, horiz center, wrap on; borders: top_color black, left_color black, bottom_color black, right_color black, top thin, left thin, right dotted, bottom  dotted;'))
        ws.write_merge(row_no, row_no, 7, 17, '住　　　　　　所', xlwt.easyxf('font: height 200, name ＭＳ 明朝; align: vert center, horiz center, wrap on; borders: top_color black, left_color black, bottom_color black, right_color black, top thin, bottom dotted, right dotted'))
        ws.write_merge(row_no, row_no, 18, 24, '電　　　　　話', xlwt.easyxf('font: height 200, name ＭＳ 明朝; align: vert center, horiz center, wrap on; borders: top_color black, left_color black, bottom_color black, right_color black, top thin, bottom dotted, right thin;'))

        row_no += 1
        ws.write_merge(row_no, row_no, 0, 6, company, xlwt.easyxf('font: height 200, name ＭＳ 明朝; align: vert center, horiz center, wrap on; borders: top_color black, left_color black, bottom_color black, right_color black, bottom thin, left thin, right dotted;'))
        ws.write_merge(row_no, row_no, 7, 17, address, xlwt.easyxf('font: height 200, name ＭＳ 明朝; align: vert center, horiz center, wrap on; borders: top_color black, left_color black, bottom_color black, right_color black, bottom thin, right dotted;'))
        ws.write_merge(row_no, row_no, 18, 24, tel, xlwt.easyxf('font: height 200, name ＭＳ 明朝; align: vert center, horiz center, wrap on; borders: top_color black, left_color black, bottom_color black, right_color black, bottom thin, right thin;'))

        row_no += 1
        ws.write_merge(row_no, row_no + 29, 0, 0, '契\n\n\n\n\n約\n\n\n\n\n条\n\n\n\n\n項', xlwt.easyxf('font: height 200, name ＭＳ 明朝; align: vert center, horiz center, wrap on; borders: top_color black, left_color black, bottom_color black, right_color black, top thin, bottom thin, left thin, right dotted;'))
        ws.write_merge(row_no, row_no, 1, 8, '品　　　　　名', xlwt.easyxf('font: height 200, name ＭＳ 明朝; align: vert center, horiz center, wrap on; borders: top_color black, left_color black, bottom_color black, right_color black, top thin, bottom dotted, right dotted;'))
        ws.write_merge(row_no, row_no, 9, 10, '台数', xlwt.easyxf('font: height 200, name ＭＳ 明朝; align: vert center, horiz center, wrap on; borders: top_color black, left_color black, bottom_color black, right_color black, top thin, bottom dotted, right dotted;'))
        ws.write_merge(row_no, row_no, 11, 13, '単　価', xlwt.easyxf('font: height 200, name ＭＳ 明朝; align: vert center, horiz center, wrap on; borders: top_color black, left_color black, bottom_color black, right_color black, top thin, bottom dotted, right dotted;'))
        ws.write_merge(row_no, row_no, 14, 17, '金　　　　額', xlwt.easyxf('font: height 200, name ＭＳ 明朝; align: vert center, horiz center, wrap on; borders: top_color black, left_color black, bottom_color black, right_color black, top thin, bottom dotted, right dotted;'))
        ws.write_merge(row_no, row_no, 18, 24, '備　　　　　考', xlwt.easyxf('font: height 200, name ＭＳ 明朝; align: vert center, horiz center, wrap on; borders: top_color black, left_color black, bottom_color black, right_color black, top thin, bottom dotted, right thin;'))

        row_no += 1

        if num_of_products:
            for form in product_formset.forms:
                form.is_valid()

                id = form.cleaned_data.get('product_id')
                product_name = Product.objects.get(id=id).name
                quantity = form.cleaned_data.get('quantity', 0)
                price = form.cleaned_data.get('price', 0)
                amount = quantity * price

                total_quantity += quantity
                total_price += price

                if total_number > 12: 
                    continue

                ws.write_merge(row_no, row_no, 1, 8, product_name, s_dot_rb)
                ws.write_merge(row_no, row_no, 9, 10, quantity, s_dot_rb_num)
                ws.write_merge(row_no, row_no, 11, 13, price, s_dot_rb_num)
                ws.write_merge(row_no, row_no, 14, 17, amount, s_dot_rb_num)
                ws.write_merge(row_no, row_no, 18, 24, '', s_dot_b_thin_r)

                row_no += 1

        # Document Table
        if num_of_documents:
            for form in document_formset.forms:
                form.is_valid()

                id = form.cleaned_data.get('document_id')
                document_name = Document.objects.get(id=id).name.replace('（売上）', '').replace('（仕入）', '')
                quantity = form.cleaned_data.get('quantity', 0)
                price = form.cleaned_data.get('price', 0)
                amount = quantity * price

                total_quantity += quantity
                total_price += price

                if total_number > 12: 
                    continue

                ws.write_merge(row_no, row_no, 1, 8, document_name, s_dot_rb)
                ws.write_merge(row_no, row_no, 9, 10, quantity, s_dot_rb_num)
                ws.write_merge(row_no, row_no, 11, 13, price, s_dot_rb_num)
                ws.write_merge(row_no, row_no, 14, 17, amount, s_dot_rb_num)
                ws.write_merge(row_no, row_no, 18, 24, '', s_dot_b_thin_r)

                row_no += 1
        
        if num_of_document_fees:
            for form in document_fee_formset.forms:
                form.is_valid()

                id = form.cleaned_data.get('document_fee_id')
                document_fee = DocumentFee.objects.get(id=id)
                type = document_fee.type
                model_count = form.cleaned_data.get('model_count', 0)
                unit_count = form.cleaned_data.get('unit_count', 0)
                model_price = document_fee.model_price
                unit_price = document_fee.unit_price
                amount = model_price * model_count + unit_price * unit_count + document_fee.application_fee

                total_quantity += model_count
                total_price += unit_count

                if total_number > 12: 
                    continue
                
                ws.write_merge(row_no, row_no, 1, 8, str(dict(TYPE_CHOICES)[type]), s_dot_rb)
                ws.write_merge(row_no, row_no, 9, 10, None, s_dot_rb_num)
                ws.write_merge(row_no, row_no, 11, 13, None, s_dot_rb_num)
                ws.write_merge(row_no, row_no, 14, 17, amount, s_dot_rb_num)
                ws.write_merge(row_no, row_no, 18, 24, '', s_dot_b_thin_r)

                row_no += 1

        forRange = 12 - total_number if total_number <= 12 else 12
        for i in range(0, forRange):

            ws.write_merge(row_no, row_no, 1, 8, None, s_dot_rb)
            ws.write_merge(row_no, row_no, 9, 10, None, s_dot_rb)
            ws.write_merge(row_no, row_no, 11, 13, None, s_dot_rb)
            ws.write_merge(row_no, row_no, 14, 17, None, s_dot_rb)
            ws.write_merge(row_no, row_no, 18, 24, None, s_dot_b_thin_r)

            row_no += 1

        ws.write_merge(row_no, row_no, 1, 8, '小　　　　　　　計', s_dot_rb)
        ws.write_merge(row_no, row_no, 9, 10, None, s_dot_rb_num)
        ws.write_merge(row_no, row_no, 11, 13, None, s_dot_rb_num)
        ws.write_merge(row_no, row_no, 14, 17, sub_total, s_dot_rb_num)
        ws.write_merge(row_no, row_no, 18, 24, None, s_dot_b_thin_r)

        row_no += 1
        ws.write_merge(row_no, row_no, 1, 8, '消　　　費　　　税', s_dot_rb)
        ws.write_merge(row_no, row_no, 9, 10, None, s_dot_rb_num)
        ws.write_merge(row_no, row_no, 11, 13, None, s_dot_rb_num)
        ws.write_merge(row_no, row_no, 14, 17, tax, s_dot_rb_num)
        ws.write_merge(row_no, row_no, 18, 24, None, s_dot_b_thin_r)

        row_no += 1
        ws.write_merge(row_no, row_no, 1, 8, '合　　　　　　　計', s_dot_rb)
        ws.write_merge(row_no, row_no, 9, 10, total_quantity, s_dot_rb_num)
        ws.write_merge(row_no, row_no, 11, 13, total_price, s_dot_rb_num)
        ws.write_merge(row_no, row_no, 14, 17, total, xlwt.easyxf('font: height 280, name ＭＳ 明朝; align: vert center, horiz center, wrap on; borders: top_color black, left_color black, bottom_color black, right_color black, bottom dotted, right dotted;', num_format_str='#,##0'))
        ws.write_merge(row_no, row_no, 18, 24, None, s_dot_b_thin_r)

        row_no += 1
        ws.row(row_no).height = int(20 * 6)
        ws.write(row_no, 24, None, s_thin_r)

        row_no += 1
        ws.row(row_no).height = int(cell_height / 2)
        ws.row(row_no + 1).height = int(cell_height / 2)
        ws.write_merge(row_no, row_no + 1, 1, 3, '契　約　日', s_dot_rtb)
        # ws.write_merge(row_no, row_no + 1, 4, 4, '令和', s_dot_tb)
        ws.write_merge(row_no, row_no + 1, 4, 5, None, s_dot_tb)
        ws.write_merge(row_no, row_no + 1, 6, 6, '年', s_dot_tb)
        ws.write_merge(row_no, row_no + 1, 7, 7, None, s_dot_tb)
        ws.write_merge(row_no, row_no + 1, 8, 8, '月', s_dot_tb)
        ws.write_merge(row_no, row_no + 1, 9, 9, None, s_dot_tb)
        ws.write_merge(row_no, row_no + 1, 10, 10, '日', s_dot_tb)
        ws.write_merge(row_no, row_no + 1, 11, 13, '納品予定日', s_dot_tb)
        # ws.write_merge(row_no, row_no + 1, 14, 14, '令和', s_dot_tb)
        shipping_date = getDate(shipping_date)

        ws.write_merge(row_no, row_no + 1, 14, 15, str(shipping_date.year) if shipping_date else None, s_dot_tb)
        ws.write_merge(row_no, row_no + 1, 16, 16, '年', s_dot_tb)
        ws.write_merge(row_no, row_no + 1, 17, 17, str(shipping_date.month) if shipping_date else None, s_dot_tb)
        ws.write_merge(row_no, row_no + 1, 18, 18, '月', s_dot_tb)
        ws.write_merge(row_no, row_no + 1, 19, 19, str(shipping_date.day) if shipping_date else None, s_dot_tb)
        ws.write_merge(row_no, row_no + 1, 20, 20, '日', s_dot_tb)
        ws.write(row_no, 21, 'AM', s_none)
        ws.write(row_no + 1, 21, 'PM', s_dot_b)
        ws.write_merge(row_no, row_no + 1, 22, 24, None, s_dot_tb_thin_r)

        row_no += 2
        ws.row(row_no).height = int(cell_height / 2)
        ws.row(row_no + 1).height = int(cell_height / 2)
        ws.write_merge(row_no, row_no + 1, 1, 3, '集　金　日', s_dot_rb)
        # ws.write_merge(row_no, row_no + 1, 4, 4, '令和', s_dot_b)
        ws.write_merge(row_no, row_no + 1, 4, 5, None, s_dot_b)
        ws.write_merge(row_no, row_no + 1, 6, 6, '年', s_dot_b)
        ws.write_merge(row_no, row_no + 1, 7, 7, None, s_dot_b)
        ws.write_merge(row_no, row_no + 1, 8, 8, '月', s_dot_b)
        ws.write_merge(row_no, row_no + 1, 9, 9, None, s_dot_b)
        ws.write_merge(row_no, row_no + 1, 10, 10, '日', s_dot_b)
        ws.write_merge(row_no, row_no + 1, 11, 13, '開　店　日', s_dot_b)
        # ws.write_merge(row_no, row_no + 1, 14, 14, '令和', s_dot_b)
        opening_date = getDate(opening_date)
        ws.write_merge(row_no, row_no + 1, 14, 15, str(opening_date.year) if opening_date else None, s_dot_b)
        ws.write_merge(row_no, row_no + 1, 16, 16, '年', s_dot_b)
        ws.write_merge(row_no, row_no + 1, 17, 17, str(opening_date.month) if opening_date else None, s_dot_b)
        ws.write_merge(row_no, row_no + 1, 18, 18, '月', s_dot_b)
        ws.write_merge(row_no, row_no + 1, 19, 19, str(opening_date.day) if opening_date else None, s_dot_b)
        ws.write_merge(row_no, row_no + 1, 20, 20, '日', s_dot_b)
        ws.write(row_no, 21, 'AM', s_none)
        ws.write(row_no + 1, 21, 'PM', s_dot_b)
        ws.write_merge(row_no, row_no + 1, 22, 24, None, s_dot_b_thin_r)

        row_no += 2
        ws.write_merge(row_no, row_no, 1, 4, 'お 支 払 方 法', s_dot_rb)
        ws.write_merge(row_no, row_no, 5, 16, '現金一括　・　現金及び手形　・　手形', s_dot_rb)
        ws.write(row_no, 24, None, s_thin_r)
        
        row_no += 1
        ws.row(row_no).height = int(20 * 18)
        ws.write_merge(row_no, row_no + 5, 1, 4, '現金及び手形　　によるお支払', s_dot_rb)
        ws.write_merge(row_no, row_no, 5, 10, '現金・小切手・振込', s_dot_rb)
        ws.write_merge(row_no, row_no, 11, 12, '期日', s_dot_rb)
        ws.write(row_no, 13, ':', s_dot_rb)
        # ws.write(row_no, 14, '令和', s_dot_rb)
        ws.write_merge(row_no, row_no, 14, 15, None, s_dot_rb)
        ws.write(row_no, 16, '年', s_dot_rb)
        ws.write(row_no, 17, None, s_dot_rtb)
        ws.write(row_no, 18, '月', s_dot_rtb)
        ws.write(row_no, 19, None, s_dot_rtb)
        ws.write(row_no, 20, '日', s_dot_rtb)
        ws.write_merge(row_no, row_no, 21, 24, None, s_dot_tb_thin_r)

        milestone_formset = MilestoneFormSet(
            self.request.POST,
            prefix='milestone'
        )

        ind = 0
        for form in milestone_formset.forms:
            form.is_valid()
            date = form.cleaned_data.get('date')
            year = str(date.year)[0:4] if date else None
            month = date.month if date else None
            day = date.day if date else None

            amount = form.cleaned_data.get('amount')

            row_no += 1
            ws.row(row_no).height = int(20 * 18)
            if ind == 0:
                ws.write_merge(row_no, row_no, 5, 6, '手形', s_dot_rb)
                ws.write_merge(row_no, row_no, 7, 8, None, s_dot_rb)
                ws.write_merge(row_no, row_no, 9, 10, '初回', s_dot_rb)
            else:
                ws.write_merge(row_no, row_no, 5, 8, None, s_dot_rb)
                ws.write_merge(row_no, row_no, 9, 10, str(ind + 1) + '回', s_dot_rb)
            ws.write_merge(row_no, row_no, 11, 12, '期日', s_dot_rb)
            ws.write(row_no, 13, ':', s_dot_rb)
            # ws.write(row_no, 14, '令和', s_dot_rb)
            ws.write_merge(row_no, row_no, 14, 15, year, s_dot_rb)
            ws.write(row_no, 16, '年', s_dot_rb)
            ws.write(row_no, 17, month, s_dot_rb)
            ws.write(row_no, 18, '月', s_dot_rb)
            ws.write(row_no, 19, day, s_dot_rb)
            ws.write(row_no, 20, '日', s_dot_rb)
            ws.write_merge(row_no, row_no, 21, 24, amount, xlwt.easyxf('font: height 200, name ＭＳ 明朝; align: vert center, horiz center, wrap on; borders: top_color black, left_color black, bottom_color black, right_color black, bottom dotted, right thin;',  num_format_str='#,##0'))

            ind += 1

        row_no += 1
        ws.write_merge(row_no, row_no, 1, 12, '上記の通り契約いたしました', s_none)
        # ws.write_merge(row_no, row_no, 15, 16, '令和', s_none)
        ws.write_merge(row_no, row_no, 15, 17, None, s_none)
        ws.write(row_no, 18, '年', s_none)
        ws.write(row_no, 19, None, s_none)
        ws.write(row_no, 20, '月', s_none)
        ws.write(row_no, 21, None, s_none)
        ws.write(row_no, 22, '日', s_none)
        ws.write_merge(row_no, row_no, 23, 24, None, s_thin_r)
        
        row_no += 1
        ws.row(row_no).height = int(20 * 8.25)
        ws.write_merge(row_no, row_no, 1, 24, None, xlwt.easyxf('font: height 200, name ＭＳ 明朝; align: vert center, horiz center, wrap on; borders: top_color black, top dotted, right_color black, right thin, bottom_color black, bottom thin;'))

        row_no += 1
        ws.row(row_no).height = int(20 * 6)
        ws.write_merge(row_no, row_no, 0, 24, None, xlwt.easyxf('borders: bottom_color black, bottom thin'))

        row_no += 1
        ws.row(row_no).height = int(20 * 13.5)
        ws.row(row_no + 1).height = int(20 * 13.5)
        ws.row(row_no + 2).height = int(20 * 22.5)
        ws.row(row_no + 3).height = int(20 * 16.5)
        ws.row(row_no + 4).height = int(20 * 15)
        ws.row(row_no + 5).height = int(20 * 15)
        ws.write_merge(row_no, row_no + 5, 0, 0, '買\n\n\n\n\n主', xlwt.easyxf('font: height 200, name ＭＳ 明朝; align: vert center, horiz center, wrap on; borders: top_color black, left_color black, right_color black, left thin, right thin;'))
        ws.write_merge(row_no, row_no + 1, 1, 11, '〒537-0021 \n大阪府大阪市東成区東中本2-4-15', xlwt.easyxf('font: height 180, name ＭＳ 明朝; align: vert center, horiz center, wrap on;'))
        ws.write_merge(row_no + 2, row_no + 2, 1, 11, 'バッジオ株式会社', xlwt.easyxf('font: height 360, name ＭＳ 明朝; align: vert center, horiz center, wrap on;'))
        ws.write_merge(row_no + 3, row_no + 3, 1, 11, '代表取締役　　　金　　昇志', xlwt.easyxf('font: height 240, name ＭＳ 明朝; align: vert center, horiz center, wrap on;'))
        ws.write_merge(row_no + 4, row_no + 4, 1, 11, 'TEL　06-6753-8078', xlwt.easyxf('font: height 220, name ＭＳ 明朝; align: vert center, horiz center, wrap on;'))
        ws.write_merge(row_no + 5, row_no + 5, 1, 11, 'FAX　06-6753-8079', xlwt.easyxf('font: height 220, name ＭＳ 明朝; align: vert center, horiz center, wrap on;'))
        ws.write_merge(row_no, row_no + 4, 12, 12, None, s_none)
        ws.write(row_no + 5, 12, '㊞', s_dot_rb)
        
        ws.write_merge(row_no, row_no + 5, 13, 13, '売\n\n\n\n\n主', xlwt.easyxf('font: height 200, name ＭＳ 明朝; align: vert center, horiz center, wrap on; borders: top_color black, left_color black, right_color black, left thin, right thin;'))
        ws.write_merge(row_no, row_no + 5, 14, 23, None, xlwt.easyxf('font: height 180, name ＭＳ 明朝; align: vert center, horiz center, wrap on;'))
        # ws.write_merge(row_no + 2, row_no + 2, 14, 23, None, xlwt.easyxf('font: height 360, name ＭＳ 明朝; align: vert center, horiz center, wrap on;'))
        # ws.write_merge(row_no + 3, row_no + 3, 14, 23, None, xlwt.easyxf('font: height 240, name ＭＳ 明朝; align: vert center, horiz center, wrap on;'))
        # ws.write_merge(row_no + 4, row_no + 4, 14, 23, None, xlwt.easyxf('font: height 220, name ＭＳ 明朝; align: vert center, horiz center, wrap on;'))
        # ws.write_merge(row_no + 5, row_no + 5, 14, 23, None, xlwt.easyxf('font: height 220, name ＭＳ 明朝; align: vert center, horiz center, wrap on;'))
        ws.write_merge(row_no, row_no + 4, 24, 24, None, xlwt.easyxf('font: height 220, name ＭＳ 明朝; align: vert center, horiz center, wrap on; borders: right_color black, right thin;'))
        ws.write(row_no + 5, 24, '㊞', s_dot_b_thin_r)
        
        row_no += 6
        ws.row(row_no).height = int(20 * 12)
        ws.write_merge(row_no, row_no, 0, 24, None, xlwt.easyxf('borders: top_color black, top thin'))

        row_no += 1
        ws.row(row_no).height = int(20 * 18)
        ws.write_merge(row_no, row_no, 0, 24, '※ お振込み手数料は御社負担でお願い致します。', xlwt.easyxf('font: height 200, name ＭＳ 明朝, bold on; align: vert center, horiz right, wrap on; borders: top_color black, left_color black, bottom_color black, right_color black, top thin, bottom double, left medium, right thin;'))

        row_no += 1
        ws.write_merge(row_no, row_no, 0, 24, ' ＜振込先＞', xlwt.easyxf('font: height 300, name ＭＳ 明朝; align: vert center, horiz left, wrap on; borders: left_color black, bottom_color black, right_color black, bottom thin, left medium, right thin;'))

        row_no += 1
        ws.row(row_no).height = int(20 * 12)

        row_no += 1
        ws.row(row_no).height = int(20 * 20.25)
        ws.write_merge(row_no, row_no, 0, 24, '※売買証明書の弊社到着をもちまして契約が成立したものとします。', xlwt.easyxf('font: height 200, name ＭＳ 明朝, bold on; align: vert center, horiz left, wrap on;'))

        row_no += 1
        ws.write_merge(row_no, row_no, 0, 24, '購入証明書送付後、キャンセルした場合には売買成立代金の全額を賠償するものとします。', xlwt.easyxf('font: height 200, name ＭＳ 明朝, bold on; align: vert center, horiz left, wrap on;'))

        row_no += 1
        ws.write_merge(row_no, row_no, 0, 24, '欠品、破損等のご連絡は商品到着後3日以内とします。それ以降の返品、交換は一切出来かねます。', xlwt.easyxf('font: height 180, name ＭＳ 明朝, bold on; align: vert center, horiz left, wrap on;'))

        # -------------------------------------------------------------------------------------------------------------------------
        if total_number > 12:

            ws = wb.add_sheet('別紙', cell_overwrite_ok=True)

            ws.set_header_str(str.encode(''))
            ws.set_footer_str(str.encode(''))
            ws.set_left_margin(0.314961)
            ws.set_top_margin(0.393701)
            ws.set_right_margin(0.314961)
            ws.set_bottom_margin(0.19685)

            w_arr = [9.15, 8.04, 10.04, 2.04, 5.04, 8.04, 6.71, 7.15, 12.59, 6.71, 6.71]
            for i in range(11): 
                ws.col(i).width = int(300 * w_arr[i])

            for i in range(56):
                ws.row(i).height_mismatch = True
                ws.row(i).height = int(20 * 18)
            ws.row(44).height = int(20 * 3.75)
            ws.row(45).height = int(20 * 15.75)

            ####################################### Table #####################################
            ws.write_merge(0, 0, 0, 6, '商　品　名', TTL_CENTER)
            ws.write(0, 7, '数量', TT_CENTER)
            ws.write(0, 8, '単　価', TT_CENTER)
            ws.write_merge(0, 0, 9, 10, '金　額', TTR_CENTER)

            row_no = 1

            total_quantity = total_price = 0
            ##### Products
            if num_of_products:
                for form in product_formset.forms:
                    form.is_valid()

                    id = form.cleaned_data.get('product_id')
                    product_name = Product.objects.get(id=id).name
                    quantity = form.cleaned_data.get('quantity', 0)
                    price = form.cleaned_data.get('price', 0)
                    amount = quantity * price

                    total_quantity += quantity
                    total_price += price

                    ws.write_merge(row_no, row_no, 0, 6, product_name, TL)
                    ws.write(row_no, 7, quantity, TC_RIGHT)
                    ws.write(row_no, 8, price, TC_RIGHT)
                    ws.write_merge(row_no, row_no, 9, 10, amount, TR_RIGHT)

                    row_no += 1

            ###### Documents
            if num_of_documents:
                for form in document_formset.forms:
                    form.is_valid()

                    id = form.cleaned_data.get('document_id')
                    document_name = Document.objects.get(id=id).name.replace('（売上）', '').replace('（仕入）', '')
                    quantity = form.cleaned_data.get('quantity', 0)
                    price = form.cleaned_data.get('price', 0)
                    amount = quantity * price

                    total_quantity += quantity
                    total_price += price

                    ws.write_merge(row_no, row_no, 0, 6, document_name, TL)
                    ws.write(row_no, 7, quantity, TC_RIGHT)
                    ws.write(row_no, 8, price, TC_RIGHT)
                    ws.write_merge(row_no, row_no, 9, 10, amount, TR_RIGHT)

                    row_no += 1

            if num_of_document_fees:
                for form in document_fee_formset.forms:
                    form.is_valid()

                    id = form.cleaned_data.get('document_fee_id')
                    document_fee = DocumentFee.objects.get(id=id)
                    type = document_fee.type
                    model_count = form.cleaned_data.get('model_count', 0)
                    unit_count = form.cleaned_data.get('unit_count', 0)
                    model_price = document_fee.model_price
                    unit_price = document_fee.unit_price
                    amount = model_price * model_count + unit_price * unit_count + document_fee.application_fee

                    total_quantity += quantity
                    total_price += price

                    ws.write_merge(row_no, row_no, 0, 6, str(dict(TYPE_CHOICES)[type]), TL)
                    ws.write(row_no, 7, None, TC_RIGHT)
                    ws.write(row_no, 8, None, TC_RIGHT)
                    ws.write_merge(row_no, row_no, 9, 10, amount, TR_RIGHT)

                    row_no += 1

            if total_number + 1 < 39:
                for i in range(0, 39 - total_number + 1):

                    ws.write_merge(row_no, row_no, 0, 6, None, TL)
                    ws.write(row_no, 7, None, TC_RIGHT)
                    ws.write(row_no, 8, None, TC_RIGHT)
                    ws.write_merge(row_no, row_no, 9, 10, None, TR_RIGHT)

                    row_no += 1

            ws.write_merge(40, 40, 0, 6, '※備考', LINE_TOP)
            ws.write_merge(40, 40, 7, 8, '小　計', TTL_CENTER)
            ws.write_merge(40, 40, 9, 10, sub_total, TTR_RIGHT)
            ws.write_merge(41, 41, 7, 8, '消費税（10%）', TL_CENTER)        
            ws.write_merge(41, 41, 9, 10, tax, TR_RIGHT)
            ws.write_merge(42, 42, 7, 8, '保険代（非課税）', TL_CENTER)
            ws.write_merge(42, 42, 9, 10, fee, TR_RIGHT)
            ws.write_merge(43, 43, 7, 8, '合　計', TBL_CENTER)
            ws.write_merge(43, 43, 9, 10, total, TBR_RIGHT_TOTAL)

            ws.write_merge(45, 45, 0, 3, 'No.{}'.format(contract_id), F11)

        wb.save(response)
        return response