import html
import importlib
import types
from io import BytesIO

import frappe
import openpyxl
from frappe import _
from frappe.utils import cint
from frappe.utils import safe_exec as frappe_safe_exec
from frappe.utils.file_manager import save_file
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.workbook import Workbook
from RestrictedPython import safe_builtins

XLSX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
ALLOWED_TEXT_ALIGN = {"left", "center", "right", "justify"}


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
    wb = run_format_script(
        script,
        {"doc": doc},
        script_name=excel_format or doctype,
    )

    output = BytesIO()
    wb.save(output)
    file_name = f"XF_Excel_{doctype}_{doc_name}.xlsx"

    if is_preview:
        return {"preview": generate_excel_preview(output)}

    if attach:
        file_doc = save_file(file_name, output.getvalue(), doctype, doc_name, is_private=1)
        return file_doc.file_url

    frappe.local.response.filename = file_name
    frappe.local.response.filecontent = output.getvalue()
    frappe.local.response.type = "binary"
    frappe.local.response.mime_type = XLSX_MIME_TYPE


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


def run_format_script(script, context, script_name=None):
    """Run an Excel format script in Frappe's sandbox and return the workbook it built.

    Scripts get `wb`, `ws`, the openpyxl style helpers and the safe `frappe` namespace,
    plus everything in `context` (for example `doc`).
    """
    if not frappe_safe_exec.is_safe_exec_enabled():
        frappe.throw(
            _(
                "XcelForm runs Excel format scripts in Frappe's sandbox, which is turned off on this server. "
                "Ask your administrator to run: bench set-config -g server_script_enabled 1"
            ),
            title=_("Excel format scripts are disabled"),
        )

    wb = openpyxl.Workbook()
    script_globals = {
        "__builtins__": {**safe_builtins, "__import__": _import_openpyxl_names},
        "wb": wb,
        "ws": wb.active,
        "Font": Font,
        "Alignment": Alignment,
        "Border": Border,
        "Side": Side,
        "PatternFill": PatternFill,
        "Image": Image,
        "get_column_letter": get_column_letter,
        **context,
    }

    try:
        exec_globals, _locals = frappe_safe_exec.safe_exec(
            script,
            script_globals,
            restrict_commit_rollback=True,
            script_filename=f"XF Excel Format {script_name or ''}",
        )
    except (frappe.ValidationError, frappe.PermissionError):
        raise
    except Exception as e:
        frappe.log_error(title=f"XF Excel Format script failed: {script_name}")
        frappe.throw(
            _("The Excel format script failed: {0}").format(html.escape(f"{type(e).__name__}: {e}")),
            title=_("Excel format error"),
        )

    result = exec_globals.get("wb")
    if not isinstance(result, Workbook):
        frappe.throw(_("The Excel format script must leave an openpyxl workbook in the variable 'wb'."))
    return result


def _import_openpyxl_names(name, globals=None, locals=None, fromlist=(), level=0):
    """Import hook for format scripts: allows only `from openpyxl[.x] import Name`.

    Returns a plain namespace holding just the requested names, never a module object,
    so scripts cannot reach other modules through it.
    """
    if level or not fromlist or "*" in fromlist or not (name == "openpyxl" or name.startswith("openpyxl.")):
        raise ImportError(
            f"Excel format scripts can only use 'from openpyxl... import Name' (got import of '{name}')."
        )

    module = importlib.import_module(name)
    names = {}
    for attr in fromlist:
        value = getattr(module, attr, None) if not attr.startswith("_") else None
        if value is None or isinstance(value, types.ModuleType):
            raise ImportError(f"Cannot import '{attr}' from '{name}' in an Excel format script.")
        names[attr] = value
    return types.SimpleNamespace(**names)


def rgb_to_hex(rgb):
    """Convert an Excel ARGB colour to a CSS hex colour, falling back to white."""
    if not isinstance(rgb, str) or len(rgb) < 6 or rgb == "00000000":
        return "#FFFFFF"
    hex_part = rgb[-6:]
    if any(c not in "0123456789abcdefABCDEF" for c in hex_part):
        return "#FFFFFF"
    return f"#{hex_part}"


def generate_excel_preview(output):
    """Render the active sheet of a workbook as an HTML table. All cell content is escaped."""
    output.seek(0)
    ws = load_workbook(BytesIO(output.read())).active

    spans = {}
    covered = set()
    for merged_range in ws.merged_cells.ranges:
        spans[(merged_range.min_row, merged_range.min_col)] = (
            merged_range.max_row - merged_range.min_row + 1,
            merged_range.max_col - merged_range.min_col + 1,
        )
        for row in range(merged_range.min_row, merged_range.max_row + 1):
            for col in range(merged_range.min_col, merged_range.max_col + 1):
                if (row, col) != (merged_range.min_row, merged_range.min_col):
                    covered.add((row, col))

    parts = ["<table border='1' style='border-collapse: collapse; width: 100%; text-align: center;'>"]
    for row_idx, row in enumerate(ws.iter_rows(), start=1):
        parts.append("<tr>")
        for col_idx, cell in enumerate(row, start=1):
            if (row_idx, col_idx) in covered:
                continue

            rowspan, colspan = spans.get((row_idx, col_idx), (1, 1))
            span_attrs = (f' rowspan="{rowspan}"' if rowspan > 1 else "") + (
                f' colspan="{colspan}"' if colspan > 1 else ""
            )

            bold = "font-weight: bold;" if cell.font and cell.font.bold else ""
            bg_color = rgb_to_hex(cell.fill.start_color.rgb) if cell.fill and cell.fill.fill_type else "#FFFFFF"
            horizontal = cell.alignment.horizontal if cell.alignment else None
            text_align = horizontal if horizontal in ALLOWED_TEXT_ALIGN else "center"
            value = html.escape(str(cell.value)) if cell.value is not None else ""

            style = f"background-color: {bg_color}; text-align: {text_align}; padding: 5px; border: 1px solid black; {bold}"
            parts.append(f'<td style="{style.strip()}"{span_attrs}>{value}</td>')
        parts.append("</tr>")
    parts.append("</table>")

    return "".join(parts)
