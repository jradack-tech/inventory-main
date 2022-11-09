from django import template
from contracts.utilities import date_dump

register = template.Library()

@register.simple_tag(takes_context=False)
def ordinal(num):
    return "%d%s" % (num,"tsnrhtdd"[(num//10%10!=1)*(num%10<4)*num%10::4])

@register.simple_tag(takes_context=True)
def date_conversion(context, date):
    if date:
        return date_dump(date, context['request'].LANGUAGE_CODE)
    return ""

@register.simple_tag(takes_context=False)
def first_item(value):
    return value * 2 - 1

@register.simple_tag(takes_context=False)
def second_item(value):
    return value * 2
