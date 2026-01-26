from datetime import datetime
import re


def bdate(date_str):
    if date_str == '':
        return ''
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", date_str):
        date_obj = datetime.strptime(date_str, '%Y-%m-%d')
        return date_obj.strftime("%d %B %Y")
    elif re.fullmatch(r"\d{4}-\d{2}", date_str):
        date_obj = datetime.strptime(date_str, '%Y-%m')
        return date_obj.strftime("%B %Y")
    return "Invalid date format"
