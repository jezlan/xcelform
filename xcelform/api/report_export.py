"""Whitelisted endpoints for exporting query and script reports to Excel."""

import frappe
from frappe import _
from frappe.desk.query_report import get_report_doc
from frappe.utils import cint, nowdate

from xcelform.core.output import download_response, preview_response
from xcelform.core.sandbox import run_format_script
from xcelform.sources.report import get_report_data


@frappe.whitelist()
def get_report_formats(report_name):
    """Enabled formats for a report, default first. Empty if the user cannot open the report."""
    if not can_open_report(report_name):
        return []
    return frappe.get_all(
        "XF Report Excel Format",
        filters={"report": report_name, "enabled": 1},
        fields=["name", "is_default"],
        order_by="is_default desc, name asc",
    )


@frappe.whitelist()
def export_report(report_name, filters=None, excel_format=None, is_preview=False):
    """Run the report with `filters` and build the Excel file with the chosen format.

    Returns preview HTML when `is_preview` is set, otherwise sends the .xlsx file.
    """
    script = get_report_format_script(report_name, excel_format)
    report = get_report_data(report_name, filters)

    wb = run_format_script(
        script,
        {
            "report_name": report.report_name,
            "filters": report.filters,
            "columns": report.columns,
            "data": report.data,
            "total_row": report.total_row,
        },
        script_name=excel_format or report_name,
    )

    if cint(is_preview):
        return preview_response(wb)

    download_response(wb, f"{report_name} {nowdate()}.xlsx")


def can_open_report(report_name):
    """True if the current user may run the report. Checks quietly, without error popups."""
    muted = frappe.flags.mute_messages
    frappe.flags.mute_messages = True
    try:
        get_report_doc(report_name)
        return True
    except (frappe.PermissionError, frappe.ValidationError):
        return False
    finally:
        frappe.flags.mute_messages = muted


def get_report_format_script(report_name, excel_format=None):
    filters = {"report": report_name, "enabled": 1}
    if excel_format:
        filters["name"] = excel_format
    else:
        filters["is_default"] = 1

    script = frappe.db.get_value("XF Report Excel Format", filters, "python_script")
    if not script:
        if excel_format:
            frappe.throw(_("Excel Format {0} is not enabled for report {1}.").format(excel_format, report_name))
        frappe.throw(_("No enabled default Excel Format is set for report {0}.").format(report_name))
    return script
