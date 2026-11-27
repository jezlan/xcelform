"""Turns Python and Frappe values into values Excel can store safely."""

import datetime
from decimal import Decimal

from frappe.utils import get_datetime, getdate, strip_html_tags
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

EXCEL_MAX_TEXT = 32767


def excel_value(value, fieldtype=None):
    """Return `value` as something openpyxl can write: a number, date, bool, text or None.

    Text has HTML tags and characters Excel rejects removed. With `fieldtype` "Date" or
    "Datetime", date strings become real dates so Excel can sort and format them.
    """
    if value is None or value == "":
        return None
    if isinstance(value, bool | int | float):
        return value
    if isinstance(value, Decimal):
        return float(value)
    if isinstance(value, datetime.datetime):
        return value.replace(tzinfo=None)
    if isinstance(value, datetime.date | datetime.time | datetime.timedelta):
        return value

    text = value if isinstance(value, str) else str(value)
    if fieldtype in ("Date", "Datetime"):
        try:
            return getdate(text) if fieldtype == "Date" else get_datetime(text)
        except Exception:
            pass
    if "<" in text:
        text = strip_html_tags(text)
    return ILLEGAL_CHARACTERS_RE.sub("", text)[:EXCEL_MAX_TEXT]


def set_cell_value(cell, value, fieldtype=None):
    """Write a value to a cell. Text that starts with "=" stays text and never becomes a formula."""
    cell.value = excel_value(value, fieldtype)
    if cell.data_type == "f":
        cell.data_type = "s"
    return cell
