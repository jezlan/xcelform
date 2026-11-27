"""Builds a starter template for a doctype: its main fields and child tables, already
filled with placeholders, so a format can start from a file instead of an empty sheet."""

import frappe
from frappe import _
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from xcelform.core.helpers import number_format

VALUE_FIELDTYPES = {
    "Data", "Link", "Dynamic Link", "Select", "Date", "Datetime", "Time", "Currency", "Float", "Int",
    "Percent", "Check", "Small Text", "Text", "Long Text", "Read Only", "Phone", "Duration", "Rating",
}  # fmt: skip
TOTAL_FIELDTYPES = {"Currency", "Float", "Int"}
TOTAL_NAME_HINTS = ("amount", "qty", "total")  # sum these columns, not rates or factors
SKIP_FIELDS = {"naming_series", "amended_from"}
MAX_HEADER_FIELDS = 12
MAX_TABLE_COLUMNS = 8
MAX_TABLES = 3

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEADER_FILL = PatternFill("solid", start_color="F2F2F2")


def build_starter_template(doctype):
    meta = frappe.get_meta(doctype)
    if meta.istable or meta.issingle:
        frappe.throw(_("Pick a regular doctype, not a child table or a single doctype."))

    wb = Workbook()
    ws = wb.active
    ws.title = _safe_sheet_title(doctype)

    tables = [(df, frappe.get_meta(df.options)) for df in meta.get_table_fields() if not df.hidden][:MAX_TABLES]
    width = max([2] + [len(_pick_fields(child, MAX_TABLE_COLUMNS)) for _df, child in tables])

    ws["A1"] = _(doctype)
    ws["A1"].font = Font(bold=True, size=14)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=width)
    ws["A2"] = "{{ doc.name }}"
    ws["A2"].font = Font(bold=True)

    row = 4
    for df in _pick_fields(meta, MAX_HEADER_FIELDS):
        ws.cell(row=row, column=1, value=_(df.label)).font = Font(bold=True)
        _placeholder(ws.cell(row=row, column=2), f"doc.{df.fieldname}", df.fieldtype)
        row += 1

    for table_df, child in tables:
        columns = _pick_fields(child, MAX_TABLE_COLUMNS)
        if not columns:
            continue
        row += 1
        ws.cell(row=row, column=1, value=_(table_df.label)).font = Font(bold=True, size=12)
        row += 1
        for col, df in enumerate(columns, start=1):
            cell = ws.cell(row=row, column=col, value=_(df.label))
            cell.font = Font(bold=True)
            cell.fill = HEADER_FILL
            cell.border = BORDER
            cell.alignment = Alignment(horizontal="center")
        repeat_row = row + 1
        for col, df in enumerate(columns, start=1):
            _placeholder(ws.cell(row=repeat_row, column=col), f"{table_df.fieldname}.{df.fieldname}", df.fieldtype)
            ws.cell(row=repeat_row, column=col).border = BORDER

        total_row = repeat_row + 1
        totals = [
            (col, df)
            for col, df in enumerate(columns, start=1)
            if df.fieldtype in TOTAL_FIELDTYPES and any(hint in df.fieldname for hint in TOTAL_NAME_HINTS)
        ]
        if totals:
            ws.cell(row=total_row, column=1, value=_("Total")).font = Font(bold=True)
            for col, df in totals:
                letter = get_column_letter(col)
                cell = ws.cell(row=total_row, column=col, value=f"=SUM({letter}{repeat_row}:{letter}{repeat_row})")
                cell.font = Font(bold=True)
                cell.number_format = number_format(df.fieldtype) or "General"
        row = total_row + 1

    ws.column_dimensions["A"].width = 24
    for col in range(2, width + 1):
        ws.column_dimensions[get_column_letter(col)].width = 18
    return wb


def _pick_fields(meta, limit):
    """Up to `limit` fields in form order: important ones first, then the rest of the form."""
    fields = [
        df
        for df in meta.fields
        if df.fieldtype in VALUE_FIELDTYPES and not df.hidden and df.label and df.fieldname not in SKIP_FIELDS
    ]
    important = [
        df
        for df in fields
        if df.in_list_view or df.reqd or df.bold or df.in_standard_filter or df.fieldname == meta.title_field
    ]
    picked = important[:limit]
    picked += [df for df in fields if df not in picked][: limit - len(picked)]
    return sorted(picked, key=fields.index)


def _placeholder(cell, expression, fieldtype):
    cell.value = "{{ " + expression + " }}"
    fmt = number_format(fieldtype)
    if fmt:
        cell.number_format = fmt
    return cell


def _safe_sheet_title(title):
    for char in "[]:*?/\\":
        title = title.replace(char, " ")
    return title[:31] or "Sheet"
