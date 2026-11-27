"""Runs a report for export and returns its columns and rows in one consistent shape."""

import json

import frappe
from frappe import _
from frappe.desk.query_report import get_column_as_dict, normalize_result, run

SUPPORTED_REPORT_TYPES = ("Script Report", "Query Report", "Custom Report")


def parse_filters(filters):
    if not filters:
        return {}
    if isinstance(filters, str):
        filters = json.loads(filters)
    if not isinstance(filters, dict):
        frappe.throw(_("Report filters must be a JSON object."))
    return filters


def get_report_data(report_name, filters=None):
    """Run `report_name` as the current user and return columns, rows and the total row.

    Frappe's `run` checks that the user may open the report and use the given filters.
    Columns come back as dicts (fieldname, label, fieldtype, options, width) and rows as
    dicts keyed by fieldname, whatever format the report itself uses.
    """
    report_type = frappe.db.get_value("Report", report_name, "report_type")
    if not report_type:
        frappe.throw(_("Report {0} not found").format(report_name), frappe.DoesNotExistError)
    if report_type not in SUPPORTED_REPORT_TYPES:
        frappe.throw(
            _("Excel formats cannot export a {0}. Supported types: {1}.").format(
                report_type, ", ".join(SUPPORTED_REPORT_TYPES)
            )
        )

    filters = parse_filters(filters)
    result = run(report_name, filters=dict(filters), are_default_filters=False)

    if result.get("prepared_report"):
        if not result.get("doc"):
            frappe.throw(
                _(
                    "{0} runs in the background. Open the report, click Generate New Report with these "
                    "filters, wait for it to finish, then export again."
                ).format(frappe.bold(report_name)),
                title=_("Report not generated yet"),
            )

    columns = [get_column_as_dict(column) for column in result.get("columns") or []]
    data = list(result.get("result") or [])

    total_row = None
    if result.get("add_total_row") and data:
        total_row = data.pop()
        if isinstance(total_row, list | tuple):
            total_row = dict(zip([c["fieldname"] for c in columns], total_row))

    data = normalize_result(data, columns)

    return frappe._dict(
        report_name=report_name,
        report_type=report_type,
        filters=filters,
        columns=columns,
        data=data,
        total_row=total_row,
    )
