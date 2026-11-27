"""Checks an uploaded template before it is saved, and explains what was found."""

import html
import re
import zipfile
from io import BytesIO

import frappe
from frappe import _
from jinja2 import TemplateSyntaxError

from xcelform.core.template import (
    attribute_lookups,
    find_placeholders,
    load_template,
    parse_cell,
    repeat_sources_used,
    template_env,
)

MAX_TEMPLATE_SIZE = 5 * 1024 * 1024
MAX_LISTED_ERRORS = 20

# Parts of an .xlsx that openpyxl does not keep when it saves the filled file.
LOST_FEATURES = {
    "xl/vbaProject.bin": "macros",
    "xl/pivotTables/": "pivot tables",
    "xl/slicers/": "slicers",
    "xl/ctrlProps/": "form controls (buttons, check boxes)",
    "xl/diagrams/": "SmartArt",
    "xl/threadedComments/": "threaded comments (kept as plain notes)",
}
SHAPE_TAG = re.compile(rb"<(?:\w+:)?sp[\s>]")


def check_template(file_url, content, repeat_names=(), field_checker=None):
    """Validate a template and return a summary. Throws if it cannot be used."""
    file_url = file_url or ""
    if file_url.lower().endswith(".xls"):
        frappe.throw(_("Old .xls files are not supported. Open the file in Excel and save it as .xlsx."))
    if file_url.lower().endswith(".xlsm"):
        frappe.throw(_("Files with macros (.xlsm) are not supported. Save the file as .xlsx (macros are removed)."))
    if not file_url.lower().endswith(".xlsx"):
        frappe.throw(_("The template must be an Excel workbook (.xlsx)."))
    if len(content) > MAX_TEMPLATE_SIZE:
        frappe.throw(_("The template is larger than {0} MB.").format(MAX_TEMPLATE_SIZE // (1024 * 1024)))

    warnings = _lost_features(content)
    wb = load_template(content)
    env = template_env()
    errors, placeholders, repeating = [], 0, {}

    for ws in wb.worksheets:
        row_sources = {}
        for cell, expression, text in find_placeholders(ws):
            placeholders += 1
            where = f"{ws.title}!{cell.coordinate}"
            try:
                ast = parse_cell(env, expression, text)
            except TemplateSyntaxError as e:
                errors.append(_("{0}: {1}").format(where, e.message or e))
                continue

            sources = repeat_sources_used(ast, repeat_names)
            for source in sources:
                if row_sources.setdefault(cell.row, source) != source:
                    errors.append(
                        _("{0}: row {1} uses both {2} and {3}. A row can repeat for one table only.").format(
                            where, cell.row, row_sources[cell.row], source
                        )
                    )
            if field_checker:
                for name, attribute in attribute_lookups(ast):
                    problem = field_checker(name, attribute)
                    if problem:
                        errors.append(_("{0}: {1}").format(where, problem))

        for row, source in row_sources.items():
            repeating.setdefault(source, []).append(f"{ws.title}!{row}")

    if errors:
        listed = errors[:MAX_LISTED_ERRORS]
        if len(errors) > MAX_LISTED_ERRORS:
            listed.append(_("...and {0} more").format(len(errors) - MAX_LISTED_ERRORS))
        frappe.throw(
            _("Fix these cells in the template and upload it again:") + _html_list(listed),
            title=_("Template has errors"),
        )
    if not placeholders:
        warnings.append(_("No {{ }} placeholders were found, so every export will be the file as uploaded."))

    lines = [_("{0} placeholders found.").format(placeholders)]
    for source, rows in repeating.items():
        lines.append(_("Rows repeating for {0}: {1}").format(html.escape(source), html.escape(", ".join(rows))))
    if warnings:
        lines.append(_("Not kept in the exported file:") + _html_list(warnings))

    return frappe._dict(
        placeholders=placeholders,
        repeating=repeating,
        warnings=warnings,
        message="<br>".join(lines),
    )


def _lost_features(content):
    found = []
    try:
        with zipfile.ZipFile(BytesIO(content)) as archive:
            names = archive.namelist()
            for prefix, label in LOST_FEATURES.items():
                if any(name.startswith(prefix) for name in names):
                    found.append(label)
            drawings = [n for n in names if n.startswith("xl/drawings/") and n.endswith(".xml")]
            if any(SHAPE_TAG.search(archive.read(name)) for name in drawings):
                found.append(_("shapes and text boxes (use a picture instead)"))
    except zipfile.BadZipFile:
        pass  # load_template reports this
    return found


def _html_list(items):
    return "<ul>" + "".join(f"<li>{html.escape(str(item))}</li>" for item in items) + "</ul>"
