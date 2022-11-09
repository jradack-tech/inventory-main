from masterdata.models import (
    POSTAL_CODE,
    ADDRESS,
    COMPANY_NAME,
    CEO,
    TEL,
    FAX,
    TRANSFER_ACCOUNT,
    REFAX,
    P_SENSOR_NUMBER,

    FEE_SALES, FEE_PURCHASES, NO_FEE_PURCHASES, NO_FEE_SALES
)

def seller(request):
    return {
        "seller_postal_code": POSTAL_CODE,
        'seller_address': ADDRESS,
        'seller_company': COMPANY_NAME,
        'seller_ceo': CEO,
        'seller_tel': TEL,
        'seller_fax': FAX,

        'seller_payee_account': TRANSFER_ACCOUNT,
        'seller_reply_fax': REFAX,
        'seller_p_sensor_number': P_SENSOR_NUMBER,

        'fee_sales': FEE_SALES,
        'no_fee_sales': NO_FEE_SALES,
        'fee_purchases': FEE_PURCHASES,
        'no_fee_purchases': NO_FEE_PURCHASES,
    }