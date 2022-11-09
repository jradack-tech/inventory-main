
import xlwt
import json
import datetime
from django.shortcuts import render
from django.http import HttpResponse
from django.views.generic.base import TemplateView
from django.views.generic.list import ListView
from django.utils.translation import gettext as _
from django.core.paginator import Paginator
from users.views import AdminLoginRequiredMixin
from masterdata.models import NO_FEE_PURCHASES, NO_FEE_SALES, FEE_PURCHASES, FEE_SALES
from contracts.models import TraderSalesContract, HallSalesContract, TraderPurchasesContract, HallPurchasesContract
from contracts.utilities import generate_random_number, log_export_operation
from .forms import SearchForm

paginate_by = 10

cell_width = 256 * 20
wide_cell_width = 256 * 45
cell_height = int(256 * 1.5)
header_height = int(cell_height * 1.5)
font_size = 20 * 10 # pt

common_style = xlwt.easyxf('font: height 200; align: vert center, horiz left, wrap on;\
                            borders: top_color black, bottom_color black, right_color black, left_color black,\
                            left thin, right thin, top thin, bottom thin;')
date_style = xlwt.easyxf('font: height 200; align: vert center, horiz left, wrap on;\
                            borders: top_color black, bottom_color black, right_color black, left_color black,\
                            left thin, right thin, top thin, bottom thin;')
bold_style = xlwt.easyxf('font: bold on, height 200; align: vert center, horiz left, wrap on;\
                            borders: top_color black, bottom_color black, right_color black, left_color black,\
                            left thin, right thin, top thin, bottom thin;')




    




        
            
            
          
            
           
            
    def post(self, request, *args, **kwargs):

        income_items = json.loads(self.request.POST.get('printable-items'))

        user_id = self.request.user.id
        log_export_operation(user_id, "{} - {}".format(_("Accounting software CSV"), _("Sale")))

        response = HttpResponse(content_type='application/ms-excel')
        response['Content-Disposition'] = 'attachment; filename="accounting_sales_{}.xls"'.format(generate_random_number())
        wb = xlwt.Workbook(encoding='utf-8')

        
        
        if self.request.LANGUAGE_CODE == 'en':
            ws = wb.add_sheet('Accounting CSV - Sales')
            ws.set_header_str(str.encode(''))
            ws.set_footer_str(str.encode(''))
            ws.set_left_margin(0.314961)
            ws.set_top_margin(0.393701)
            ws.set_right_margin(0.314961)
            ws.set_bottom_margin(0.19685)
        else:
            ws = wb.add_sheet("{} - {}".format(_('Accounting software CSV'), _('Sale')))
            ws.set_header_str(str.encode(''))
            ws.set_footer_str(str.encode(''))
            ws.set_left_margin(0.314961)
            ws.set_top_margin(0.393701)
            ws.set_right_margin(0.314961)
            ws.set_bottom_margin(0.19685)
        
        ws.row(0).height_mismatch = True
        ws.row(0).height = header_height

        ws.write(0, 0, _('No'), bold_style)
        ws.write(0, 1, _('Contract ID'), bold_style)
        ws.write(0, 2, _('Contract date'), bold_style)
        ws.write(0, 3, _('Type'), bold_style)
        ws.write(0, 4, _('Taxation'), bold_style)
        ws.write(0, 5, _('Amount'), bold_style)
        ws.write(0, 6, _('Customer'), bold_style)

        for i in range(10):
            ws.col(i).width = cell_width
        ws.col(6).width = wide_cell_width

        row_no = 1
        contract_list = self.get_contract_list()[2]
        date_style.num_format_str = 'yyyy/mm/dd' if self.request.LANGUAGE_CODE == 'ja' else 'mm/dd/yyyy'

        for contract in contract_list:
            if contract.contract_id in income_items.keys() and income_items[contract.contract_id]:

                customer_name = ''
                try:
                    customer_name = contract.customer.excel
                    # if customer_name == None or customer_name == '':
                    #     customer_name = contract.hall.customer_name
                except:
                    try:
                        customer_name = contract.hall.payee
                    except:
                        customer_name = ''

                ws.write_merge(row_no, row_no + 1, 0, 0, "{}".format(int(row_no / 2 + 1)), common_style)
                ws.write(row_no, 1, contract.contract_id, common_style)
                ws.write(row_no, 2, contract.created_at, date_style)
                ws.write(row_no, 3, _('Income'), common_style)
                ws.write(row_no, 4, FEE_SALES, common_style)
                ws.write(row_no, 5, round(int(contract.taxed_total)), common_style)
                ws.write(row_no, 6, customer_name, common_style)

                row_no += 1

                # ws.write(row_no, 0, "{}".format(row_no), common_style)
                ws.write(row_no, 1, contract.contract_id, common_style)
                ws.write(row_no, 2, contract.created_at, date_style)
                ws.write(row_no, 3, None, common_style)
                ws.write(row_no, 4, '非課売上', common_style)
                ws.write(row_no, 5, contract.fee, common_style)
                ws.write(row_no, 6, customer_name, common_style)
                
                contract.sale_print_date = datetime.date.today()
                contract.save()
                
                row_no += 1
        
        wb.save(response)
        return response
  
