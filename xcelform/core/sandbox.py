"""Runs Excel format scripts in Frappe's server-script sandbox."""

import html
import importlib
import types

import frappe
import openpyxl
from frappe import _
from frappe.utils import safe_exec as frappe_safe_exec
from openpyxl.drawing.image import Image
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.workbook import Workbook
from RestrictedPython import compile_restricted, safe_builtins

from xcelform.core import helpers


def run_format_script(script, context, script_name=None):
    """Run an Excel format script in Frappe's sandbox and return the workbook it built.

    Scripts get `wb`, `ws`, the openpyxl style helpers, the `xf` helpers and the safe
    `frappe` namespace, plus everything in `context` (for example `doc` or `data`).
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
        "xf": helpers.namespace(),
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


def validate_script_syntax(script):
    """Throw a readable error if the script will not compile in the sandbox."""
    try:
        compile_restricted(script or "", filename="<format script>", mode="exec", policy=frappe_safe_exec.FrappeTransformer)
    except SyntaxError as e:
        message = e.msg if isinstance(e.msg, str) else "; ".join(e.msg)
        if e.lineno:
            message = _("Line {0}: {1}").format(e.lineno, message)
        frappe.throw(
            _("The script has an error. {0}").format(html.escape(message)),
            title=_("Invalid script"),
        )


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
