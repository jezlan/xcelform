import frappe
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image
from frappe.utils.file_manager import save_file


@frappe.whitelist()
def export_to_xf_excel(
    doctype, doc_name, excel_format=None, attach=False, download=False
):

    doc = frappe.get_doc(doctype, doc_name)

    excelFormat = None
    if excel_format:

        excelFormat = frappe.get_value(
            "XF Excel Format",
            {"doc_type": doctype, "name": excel_format, "enable": 1},
            "python_script",
        )
    else:

        excelFormat = frappe.get_value(
            "XF Excel Format",
            {"doc_type": doctype, "is_default": 1, "enable": 1},
            "python_script  ",
        )
        if not excelFormat:
            frappe.throw("Default enabled Excel Format not defined for this Doctype")

    if not excelFormat:
        frappe.throw("Enabled Excel Format not defined for this Doctype")

    context = {
        "wb": openpyxl.Workbook(),
        "ws": None,
        "Font": Font,
        "doc": doc,
        "Alignment": Alignment,
        "get_column_letter": get_column_letter,
        "Border": Border,
        "Side": Side,
        "PatternFill": PatternFill,
        "Image": Image,
    }

    exec(excelFormat, context)

    ws = context["ws"]

    from io import BytesIO

    output = BytesIO()
    context["wb"].save(output)
    output.seek(0)

    file_name = f"XF_Excel_{doctype}_{doc_name}.xlsx"

    if attach:
        file_doc = save_file(file_name, output.read(), doctype, doc_name, is_private=1)
        return file_doc.file_url

    if download:
        frappe.local.response.filename = file_name
        frappe.local.response.filecontent = output.getvalue()
        frappe.local.response.type = "binary"
        frappe.local.response.mime_type = (
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        return


@frappe.whitelist()
def get_default_excel_format(doctype):
    default_format = frappe.db.get_value(
        "XF Excel Format", {"doc_type": doctype, "is_default": 1}, "name"
    )
    return default_format
