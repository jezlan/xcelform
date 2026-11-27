"""Document Excel export (the "XF Excel Export" button on forms)."""

import frappe
from frappe import _
from frappe.utils import cint

from xcelform.core.output import download_response, preview_response, workbook_bytes
from xcelform.core.pipeline import FORMAT_FIELDS, build_workbook, render_file_name
from xcelform.core.preview import generate_excel_preview, rgb_to_hex  # noqa: F401 - kept for existing imports
from xcelform.core.starter import build_starter_template
from xcelform.sources.document import get_document, get_tables


@frappe.whitelist()
def export_to_xf_excel(doctype, doc_name, excel_format=None, attach=False, download=False, is_preview=False):
    attach, download, is_preview = cint(attach), cint(download), cint(is_preview)
    if not (attach or download or is_preview):
        frappe.throw(_("Choose to attach, download or preview the Excel file."))

    doc = get_document(doctype, doc_name, for_write=attach)
    fmt = get_document_format(doctype, excel_format)
    context = {"doc": doc}
    wb = build_workbook(fmt, context, get_tables(doc))

    if is_preview:
        return preview_response(wb)

    file_name = render_file_name(fmt.file_name_pattern, context, f"XF_Excel_{doctype}_{doc_name}")
    if attach:
        file_doc = frappe.get_doc(
            {
                "doctype": "File",
                "file_name": file_name,
                "attached_to_doctype": doctype,
                "attached_to_name": doc_name,
                "is_private": 1,
                "content": workbook_bytes(wb).getvalue(),
            }
        )
        file_doc.save(ignore_permissions=True)
        return file_doc.file_url

    download_response(wb, file_name)


@frappe.whitelist()
def get_default_excel_format(doctype):
    return frappe.db.get_value("XF Excel Format", {"doc_type": doctype, "is_default": 1, "enable": 1}, "name")


@frappe.whitelist()
def download_starter_template(doctype):
    """An .xlsx with placeholders for the main fields and child tables of `doctype`, ready to restyle."""
    frappe.has_permission("XF Excel Format", "create", throw=True)
    download_response(build_starter_template(doctype), f"{doctype} template.xlsx")


def get_document_format(doctype, excel_format=None):
    filters = {"doc_type": doctype, "enable": 1}
    if excel_format:
        filters["name"] = excel_format
    else:
        filters["is_default"] = 1

    fmt = frappe.db.get_value("XF Excel Format", filters, FORMAT_FIELDS, as_dict=True)
    if not fmt:
        if excel_format:
            frappe.throw(_("Excel Format {0} is not enabled for {1}.").format(excel_format, doctype))
        frappe.throw(_("No enabled default Excel Format is set for {0}.").format(doctype))
    return fmt
