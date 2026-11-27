"""Helpers available to format scripts as `xf`, so common layouts take a few lines."""

import types

import frappe
from openpyxl.styles import Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from xcelform.core.values import set_cell_value

HEADER_FILL = PatternFill("solid", start_color="F2F2F2")
THIN = Side(style="thin", color="BFBFBF")
CELL_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
TOTAL_BORDER = Border(left=THIN, right=THIN, top=Side(style="medium"), bottom=THIN)


def namespace():
    return types.SimpleNamespace(
        write_table=write_table,
        autofit_columns=autofit_columns,
        number_format=number_format,
    )


def number_format(fieldtype):
    """Excel number format for a Frappe fieldtype, or None to keep Excel's default."""
    date_format = frappe.utils.get_user_date_format().lower()
    return {
        "Currency": "#,##0.00",
        "Float": "#,##0.00",
        "Int": "#,##0",
        "Percent": '0.00"%"',
        "Date": date_format,
        "Datetime": f"{date_format} hh:mm:ss",
    }.get(fieldtype)


def write_table(ws, columns, data, start_row=1, start_col=1, total_row=None, header=True):
    """Write report-style columns and rows to `ws`, styled and with number formats.

    `columns` are dicts with `fieldname`, `label` and `fieldtype` (as reports provide them);
    `data` is a list of dicts keyed by fieldname. Hidden columns are skipped.
    Returns the number of the first empty row after the table.
    """
    columns = [c for c in columns if not c.get("hidden")]
    row = start_row

    if header:
        for offset, column in enumerate(columns):
            cell = ws.cell(row=row, column=start_col + offset, value=column.get("label") or column.get("fieldname"))
            cell.font = Font(bold=True)
            cell.fill = HEADER_FILL
            cell.border = CELL_BORDER
        row += 1

    for record in data or []:
        _write_row(ws, columns, record, row, start_col, CELL_BORDER)
        row += 1

    if total_row:
        _write_row(ws, columns, total_row, row, start_col, TOTAL_BORDER, bold=True)
        row += 1

    return row


def autofit_columns(ws, min_width=8, max_width=60):
    """Set each column's width from the longest value in it."""
    widths = {}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None:
                length = max(len(line) for line in str(cell.value).splitlines() or [""])
                widths[cell.column] = max(widths.get(cell.column, 0), length)
    for column, length in widths.items():
        ws.column_dimensions[get_column_letter(column)].width = min(max(length + 2, min_width), max_width)


def _write_row(ws, columns, record, row, start_col, border, bold=False):
    for offset, column in enumerate(columns):
        fieldtype = column.get("fieldtype")
        cell = set_cell_value(ws.cell(row=row, column=start_col + offset), record.get(column.get("fieldname")), fieldtype)
        cell.border = border
        if bold:
            cell.font = Font(bold=True)
        fmt = number_format(fieldtype)
        if fmt:
            cell.number_format = fmt
