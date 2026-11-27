"""One path from an Excel format to a finished workbook, shared by document and report export."""

import html
import os

import frappe
from frappe import _

from xcelform.core.output import safe_file_name
from xcelform.core.sandbox import run_format_script, validate_script_syntax
from xcelform.core.template import fill_template, template_env
from xcelform.core.template_check import check_template

TEMPLATE = "Template"
SCRIPT = "Script"
FORMAT_FIELDS = ["name", "format_type", "template_file", "python_script", "post_process_script", "file_name_pattern"]


def build_workbook(excel_format, context, repeat_sources=None):
    """Build the workbook for a format (a dict with FORMAT_FIELDS) from `context`."""
    if excel_format.format_type == TEMPLATE:
        wb = fill_template(read_template(excel_format.template_file), context, repeat_sources)
        if (excel_format.post_process_script or "").strip():
            wb = run_format_script(excel_format.post_process_script, context, script_name=excel_format.name, wb=wb)
        return wb

    if not (excel_format.python_script or "").strip():
        frappe.throw(_("Excel Format {0} has no script.").format(frappe.bold(excel_format.name)))
    return run_format_script(excel_format.python_script, context, script_name=excel_format.name)


def render_file_name(pattern, context, default):
    """File name from the format's pattern (Jinja, like `{{ doc.name }}`), or `default`."""
    name = ""
    if (pattern or "").strip():
        try:
            name = template_env().from_string(pattern).render(context).strip()
        except Exception as e:
            frappe.throw(_("The file name pattern has an error: {0}").format(html.escape(str(e))))
    name = (name or default).rstrip(" .")
    if not name.lower().endswith(".xlsx"):
        name += ".xlsx"
    return safe_file_name(name)


def read_template(file_url):
    if not file_url:
        frappe.throw(_("This Excel Format has no template file. Upload one first."))
    file_name = frappe.db.get_value("File", {"file_url": file_url}, "name")
    if not file_name:
        frappe.throw(_("The template file {0} was not found. Upload it again.").format(html.escape(file_url)))

    file_doc = frappe.get_doc("File", file_name)
    path = file_doc.get_full_path()
    if not os.path.exists(path):
        frappe.throw(_("The template file {0} is missing on the server. Upload it again.").format(html.escape(file_url)))
    with open(path, "rb") as f:
        return f.read()


def validate_format(doc, repeat_names=(), field_checker=None):
    """Checks shared by XF Excel Format and XF Report Excel Format on save."""
    if doc.format_type == SCRIPT:
        if not (doc.python_script or "").strip():
            frappe.throw(_("Write the Python script, or set Format Type to Template."))
        validate_script_syntax(doc.python_script)
    else:
        if not doc.template_file:
            frappe.throw(_("Upload the Excel template, or set Format Type to Script."))
        make_template_private(doc)
        summary = check_template(doc.template_file, read_template(doc.template_file), repeat_names, field_checker)
        if doc.is_new() or doc.has_value_changed("template_file") or summary.warnings:
            frappe.msgprint(summary.message, title=_("Template checked"), indicator="orange" if summary.warnings else "green")
        if (doc.post_process_script or "").strip():
            validate_script_syntax(doc.post_process_script)

    if (doc.file_name_pattern or "").strip():
        try:
            template_env().parse(doc.file_name_pattern)
        except Exception as e:
            frappe.throw(_("The file name pattern has an error: {0}").format(html.escape(str(e))))


def make_template_private(doc):
    """Templates often hold client details, so keep the uploaded file private."""
    file_name = frappe.db.get_value("File", {"file_url": doc.template_file, "is_private": 0}, "name")
    if not file_name:
        return
    file_doc = frappe.get_doc("File", file_name)
    file_doc.is_private = 1
    file_doc.save(ignore_permissions=True)
    doc.template_file = file_doc.file_url
