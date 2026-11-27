"""Sends a generated workbook to the browser as a download or as preview HTML."""

import re
from io import BytesIO

import frappe

from xcelform.core.preview import generate_excel_preview

XLSX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def workbook_bytes(wb):
    output = BytesIO()
    wb.save(output)
    return output


def preview_response(wb):
    return {"preview": generate_excel_preview(workbook_bytes(wb))}


def download_response(wb, file_name):
    frappe.local.response.filename = safe_file_name(file_name)
    frappe.local.response.filecontent = workbook_bytes(wb).getvalue()
    frappe.local.response.type = "binary"
    frappe.local.response.mime_type = XLSX_MIME_TYPE


def safe_file_name(file_name):
    """Drop characters that are not allowed in file names on Windows, macOS or Linux."""
    name = re.sub(r'[\\/:*?"<>|\r\n\t]+', "-", file_name).strip(" .-")
    return name or "export.xlsx"
