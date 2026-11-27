"""Document Excel export (the "XF Excel Export" button on forms)."""

import frappe
from frappe import _
from frappe.utils import cint
from frappe.utils.file_manager import save_file

from xcelform.core.output import download_response, preview_response, safe_file_name, workbook_bytes
from xcelform.core.preview import generate_excel_preview, rgb_to_hex  # noqa: F401 - kept for existing imports
from xcelform.core.sandbox import run_format_script


@frappe.whitelist()
def export_to_xf_excel(doctype, doc_name, excel_format=None, attach=False, download=False, is_preview=False):
    attach, download, is_preview = cint(attach), cint(download), cint(is_preview)
    if not (attach or download or is_preview):
        frappe.throw(_("Choose to attach, download or preview the Excel file."))

    doc = frappe.get_doc(doctype, doc_name)
    doc.check_permission("read")
    if attach:
        doc.check_permission("write")

    script = get_format_script(doctype, excel_format)
    wb = run_format_script(script, {"doc": doc}, script_name=excel_format or doctype)
    file_name = f"XF_Excel_{doctype}_{doc_name}.xlsx"

    if is_preview:
        return preview_response(wb)

    if attach:
        file_doc = save_file(safe_file_name(file_name), workbook_bytes(wb).getvalue(), doctype, doc_name, is_private=1)
        return file_doc.file_url

    download_response(wb, file_name)


@frappe.whitelist()
def get_default_excel_format(doctype):
    return frappe.db.get_value("XF Excel Format", {"doc_type": doctype, "is_default": 1, "enable": 1}, "name")


def get_format_script(doctype, excel_format=None):
    filters = {"doc_type": doctype, "enable": 1}
    if excel_format:
        filters["name"] = excel_format
    else:
        filters["is_default"] = 1

    script = frappe.db.get_value("XF Excel Format", filters, "python_script")
    if not script:
        if excel_format:
            frappe.throw(_("Excel Format {0} is not enabled for {1}.").format(excel_format, doctype))
        frappe.throw(_("No enabled default Excel Format is set for {0}.").format(doctype))
    return script
