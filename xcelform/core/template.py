"""Fills an uploaded Excel template.

A cell holding `{{ expression }}` gets the value of the expression. A row whose
placeholders use a repeating source (a child table such as `items`, or `data` for report
rows) is written once per record. Everything else in the file is kept as designed:
styles, merged cells, row heights, formulas, conditional formatting, dropdowns, images,
charts and print settings. When rows repeat, anything below them moves down and ranges
that cover the repeating row (like `=SUM(D5:D5)`) stretch to cover every copy.

Expressions run in Frappe's sandboxed Jinja environment, the same one print formats use.
"""

import datetime
import html
import re
from bisect import bisect_left
from copy import copy
from io import BytesIO

import frappe
from frappe import _
from jinja2 import ChainableUndefined, TemplateSyntaxError, nodes
from openpyxl import load_workbook
from openpyxl.cell.cell import Cell, MergedCell
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.formula.tokenizer import Token, Tokenizer
from openpyxl.styles.numbers import is_date_format
from openpyxl.utils.cell import column_index_from_string, coordinate_from_string
from openpyxl.worksheet.cell_range import CellRange, MultiCellRange
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.merge import MergedCellRange

from xcelform.core.helpers import number_format
from xcelform.core.values import set_cell_value

PLACEHOLDER = re.compile(r"\{\{(.*?)\}\}", re.S)
SINGLE_PLACEHOLDER = re.compile(r"^\s*\{\{(.*?)\}\}\s*$", re.S)
REF_PART = re.compile(r"^(?P<col>\$?[A-Za-z]{1,3})?(?P<row_abs>\$)?(?P<row>[1-9]\d{0,6})?$")
DATE_TYPES = ("Date", "Datetime")


def fill_template(content, context, repeat_sources=None):
    """Return an openpyxl Workbook built from the template bytes `content`.

    `context` holds the names placeholders can use (for example `doc`). `repeat_sources`
    maps a name to a list of records; a row using that name repeats once per record,
    with the name bound to the current record.
    """
    repeat_sources = repeat_sources or {}
    wb = load_template(content)
    env = template_env()
    base_context = {**repeat_sources, **context}

    sheets = [_SheetPlan(ws, env, repeat_sources) for ws in wb.worksheets]
    row_maps = {plan.ws.title: plan.row_map for plan in sheets}

    for plan in sheets:
        plan.rewrite(row_maps)
    _remap_defined_names(wb.defined_names, None, row_maps)

    for plan in sheets:
        plan.render(base_context, repeat_sources)
    return wb


def load_template(content):
    try:
        return load_workbook(BytesIO(content), rich_text=False)
    except Exception:
        frappe.throw(
            _("The template could not be opened. Save it as an Excel Workbook (.xlsx) and upload it again."),
            title=_("Invalid template"),
        )


def template_env():
    """Frappe's sandboxed Jinja environment, where missing values come out blank."""
    if not getattr(frappe.local, "xf_template_env", None):
        frappe.local.xf_template_env = frappe.get_jenv().overlay(
            undefined=ChainableUndefined, finalize=lambda value: "" if value is None else value
        )
    return frappe.local.xf_template_env


def find_placeholders(ws):
    """Yield (cell, expression or None, text) for every cell with a placeholder.

    `expression` is set when the cell holds a single placeholder and nothing else; such a
    cell gets the expression's real value (number, date...) instead of text.
    """
    for cell in sorted(ws._cells.values(), key=lambda c: (c.row, c.column)):
        value = cell.value
        if isinstance(value, str) and "{{" in value:
            single = SINGLE_PLACEHOLDER.match(value)
            expression = single.group(1).strip() if single and "{{" not in single.group(1) else None
            yield cell, expression, value


def parse_cell(env, expression, text):
    """Parse a placeholder cell and return its Jinja AST. Raises TemplateSyntaxError."""
    return env.parse("{{ " + expression + " }}" if expression is not None else text)


def attribute_lookups(ast):
    """Yield (name, attribute) for every `name.attribute` or `name["attribute"]` in the expression."""
    for node in ast.find_all((nodes.Getattr, nodes.Getitem)):
        if not isinstance(node.node, nodes.Name):
            continue
        if isinstance(node, nodes.Getattr):
            yield node.node.name, node.attr
        elif isinstance(node.arg, nodes.Const) and isinstance(node.arg.value, str):
            yield node.node.name, node.arg.value


def repeat_sources_used(ast, repeat_names):
    """The tables whose rows this cell reads, like `items` in `{{ items.qty }}`.

    Using a table as a whole (`{{ items|length }}`, `{{ items[0].qty }}`) does not make
    the row repeat.
    """
    return {name for name, _attribute in attribute_lookups(ast) if name in repeat_names}


class _RowMap:
    """Where each template row ends up once repeating rows are expanded."""

    def __init__(self, counts):
        self.rows = sorted(counts)
        self.counts = counts
        self.extra_before = []  # extra rows added above each repeating row
        extra = 0
        for row in self.rows:
            self.extra_before.append(extra)
            extra += counts[row] - 1

    def first(self, row):
        index = bisect_left(self.rows, row)
        if index < len(self.rows) and self.rows[index] == row:
            return row + self.extra_before[index]
        return row + (self.extra_before[index - 1] + self.counts[self.rows[index - 1]] - 1 if index else 0)

    def count(self, row):
        return self.counts.get(row, 1)

    def last(self, row):
        return self.first(row) + self.count(row) - 1

    def targets(self, row):
        first = self.first(row)
        return range(first, first + self.count(row))

    def shifts(self):
        return any(count != 1 for count in self.counts.values())


class _SheetPlan:
    def __init__(self, ws, env, repeat_sources):
        self.ws = ws
        self.cells = []  # (template coordinate, expression or None, text, compiled, source, row)
        self.row_sources = {}
        compiled = {}

        for cell, expression, text in find_placeholders(ws):
            try:
                ast = parse_cell(env, expression, text)
            except TemplateSyntaxError as e:
                frappe.throw(
                    _("{0}: {1}").format(cell_label(ws, cell.coordinate), html.escape(e.message or str(e))),
                    title=_("Template error"),
                )

            sources = repeat_sources_used(ast, repeat_sources)
            if len(sources) > 1:
                frappe.throw(
                    _("{0} uses {1} together. A row can repeat for one table only.").format(
                        cell_label(ws, cell.coordinate), ", ".join(sorted(sources))
                    )
                )
            source = next(iter(sources), None)
            if source:
                other = self.row_sources.setdefault(cell.row, source)
                if other != source:
                    frappe.throw(
                        _("Row {0} on sheet {1} uses both {2} and {3}. A row can repeat for one table only.").format(
                            cell.row, ws.title, other, source
                        )
                    )

            key = (expression, text)
            if key not in compiled:
                compiled[key] = (
                    env.compile_expression(expression, undefined_to_none=True)
                    if expression is not None
                    else env.from_string(text)
                )
            self.cells.append([cell.coordinate, expression, text, compiled[key], source, cell.row])

        # A repeating row is written once per record; with no records it stays as one blank row,
        # so formulas that point at it keep working.
        self.row_map = _RowMap(
            {row: max(len(repeat_sources[source] or []), 1) for row, source in self.row_sources.items()}
        )

    # Layout

    def rewrite(self, row_maps):
        ws, row_map = self.ws, self.row_map
        moved = {}
        self.copies = {}  # template row -> list of output rows

        for (row, column), cell in sorted(ws._cells.items()):
            original_value, original_link = cell._value, cell.hyperlink
            for index, target_row in enumerate(row_map.targets(row)):
                target = cell if index == 0 else _copy_cell(ws, cell, target_row)
                target.row = target_row
                if original_link:
                    target._hyperlink = copy(original_link) if index else original_link
                    target._hyperlink.ref = target.coordinate
                self._remap_cell_formula(target, original_value, row, index, row_maps)
                moved[(target_row, column)] = target

        ws._cells = moved
        if not row_map.shifts():
            return

        self._remap_merges()
        self._remap_row_heights()
        self._remap_conditional_formatting(row_maps)
        self._remap_data_validations(row_maps)
        self._remap_drawings(row_maps)
        self._remap_print_settings()
        _remap_defined_names(ws.defined_names, ws.title, row_maps)

    def _remap_cell_formula(self, cell, value, template_row, index, row_maps):
        own_row = template_row if template_row in self.row_sources else None
        if isinstance(value, ArrayFormula):
            text = remap_formula(value.text, self.ws.title, row_maps, own_row, index)
            ref = remap_range(CellRange(value.ref), self.row_map)
            cell.value = ArrayFormula(ref.coord, text)
        elif cell.data_type == "f" and isinstance(value, str):
            cell.value = remap_formula(value, self.ws.title, row_maps, own_row, index)

    def _remap_merges(self):
        ws, row_map = self.ws, self.row_map
        ranges = []
        for merged in ws.merged_cells.ranges:
            if merged.min_row == merged.max_row and merged.min_row in self.row_sources:
                # A merge inside a repeating row is repeated with it.
                for target_row in row_map.targets(merged.min_row):
                    ranges.append(
                        CellRange(min_col=merged.min_col, max_col=merged.max_col, min_row=target_row, max_row=target_row)
                    )
            else:
                ranges.append(remap_range(merged, row_map))
        ws.merged_cells = MultiCellRange()
        for cell_range in ranges:
            ws.merged_cells.add(MergedCellRange(ws, cell_range.coord))

    def _remap_row_heights(self):
        dimensions = self.ws.row_dimensions
        new = {}
        for row, dimension in list(dimensions.items()):
            for target_row in self.row_map.targets(row):
                target = copy(dimension)
                target.index = target_row
                new[target_row] = target
        dimensions.clear()
        dimensions.update(new)

    def _remap_conditional_formatting(self, row_maps):
        old = self.ws.conditional_formatting
        new = ConditionalFormattingList()
        for formatting in old:
            sqref = remap_multi_range(formatting.sqref, self.row_map)
            for rule in formatting.rules:
                if rule.formula:
                    rule.formula = [remap_formula(f, self.ws.title, row_maps) for f in rule.formula]
                new.add(sqref, rule)
        self.ws.conditional_formatting = new

    def _remap_data_validations(self, row_maps):
        for validation in self.ws.data_validations.dataValidation:
            validation.sqref = MultiCellRange(remap_multi_range(validation.sqref, self.row_map))
            for attr in ("formula1", "formula2"):
                formula = getattr(validation, attr)
                if formula:
                    setattr(validation, attr, remap_formula(formula, self.ws.title, row_maps))

    def _remap_drawings(self, row_maps):
        for drawing in [*self.ws._images, *self.ws._charts]:
            _move_anchor(drawing, self.row_map)
        for chart in self.ws._charts:
            _remap_chart_references(chart, self.ws.title, row_maps)

    def _remap_print_settings(self):
        ws, row_map = self.ws, self.row_map
        if ws.freeze_panes:
            column, row = coordinate_from_string(ws.freeze_panes)
            ws.freeze_panes = f"{column}{row_map.first(row)}"
        if ws.auto_filter.ref:
            ws.auto_filter.ref = remap_range(CellRange(ws.auto_filter.ref), row_map).coord
        for table in ws.tables.values():
            table.ref = remap_range(CellRange(table.ref), row_map).coord
            if table.autoFilter is not None:
                table.autoFilter.ref = table.ref
        if ws._print_area:
            ws._print_area = type(ws._print_area)([remap_range(r, row_map).coord for r in ws._print_area.ranges])
        if ws._print_rows:
            ws._print_rows.min_row = row_map.first(ws._print_rows.min_row)
            ws._print_rows.max_row = row_map.last(ws._print_rows.max_row)

    # Values

    def render(self, base_context, repeat_sources):
        ws = self.ws
        for template_coordinate, expression, text, compiled, source, template_row in self.cells:
            records = (repeat_sources.get(source) or [None]) if source else [None]
            column = column_index_from_string(coordinate_from_string(template_coordinate)[0])
            for target_row, record in zip(self.row_map.targets(template_row), records, strict=False):
                context = {**base_context, source: record} if source else base_context
                cell = ws._cells[(target_row, column)]
                try:
                    if source and record is None:
                        value = None  # the blank row left when a table has no records
                    elif expression is not None:
                        value = compiled(context)
                    else:
                        value = compiled.render(context)
                except Exception as e:
                    frappe.throw(
                        _("{0}: {1}").format(
                            cell_label(ws, template_coordinate), html.escape(f"{type(e).__name__}: {e}")
                        ),
                        title=_("Template error"),
                    )
                _write_value(cell, value)


def _write_value(cell, value):
    template_format = cell.number_format
    set_cell_value(cell, value)
    if isinstance(cell.value, datetime.date) and not is_date_format(template_format):
        # The designer left the cell as General: show dates the way the user sees them in Frappe.
        cell.number_format = number_format("Datetime" if isinstance(cell.value, datetime.datetime) else "Date")


def _copy_cell(ws, cell, row):
    if isinstance(cell, MergedCell):
        target = MergedCell(ws, row=row, column=cell.column)
    else:
        target = Cell(ws, row=row, column=cell.column)
        target._value = cell._value
        target.data_type = cell.data_type
    if cell.has_style:
        target._style = copy(cell._style)
    return target


def _move_anchor(drawing, row_map):
    anchor = drawing.anchor
    if isinstance(anchor, str):
        column, row = coordinate_from_string(anchor)
        drawing.anchor = f"{column}{row_map.first(row)}"
        return
    start = getattr(anchor, "_from", None)
    if start is None:
        return  # absolute anchor, placed by position rather than by cell
    shift = row_map.first(start.row + 1) - 1 - start.row
    start.row += shift
    end = getattr(anchor, "to", None)
    if end is not None:
        end.row += shift  # keep the picture's size


def _remap_chart_references(chart, sheet, row_maps):
    for series in getattr(chart, "series", []):
        for attr in ("val", "cat", "xVal", "yVal", "tx"):
            data = getattr(series, attr, None)
            for ref_attr in ("numRef", "strRef"):
                reference = getattr(data, ref_attr, None) if data is not None else None
                if reference is not None and reference.f:
                    reference.f = remap_formula(reference.f, sheet, row_maps)
                    for cache in ("numCache", "strCache"):
                        if hasattr(reference, cache):
                            setattr(reference, cache, None)  # Excel rebuilds it from the cells


def _remap_defined_names(defined_names, sheet, row_maps):
    for defined_name in defined_names.values():
        if defined_name.attr_text and not defined_name.attr_text.startswith("#"):
            defined_name.attr_text = remap_formula(defined_name.attr_text, sheet, row_maps)


# References


def remap_formula(formula, sheet, row_maps, own_row=None, index=0):
    """Move the row numbers in a formula to where those rows ended up.

    A reference to the formula's own repeating row (like `=B5*C5` in row 5) points at the
    same copy of that row. Ranges stretch over repeated rows.
    """
    has_equals = formula.startswith("=")
    try:
        tokens = Tokenizer(formula if has_equals else "=" + formula)
    except Exception:
        return formula

    changed = False
    for token in tokens.items:
        if token.type == Token.OPERAND and token.subtype == Token.RANGE:
            new_value = _remap_reference(token.value, sheet, row_maps, own_row, index)
            if new_value != token.value:
                token.value = new_value
                changed = True
    if not changed:
        return formula
    rendered = tokens.render()
    return rendered if has_equals else rendered[1:]


def _remap_reference(reference, sheet, row_maps, own_row, index):
    prefix, address = "", reference
    if "!" in reference:
        prefix, address = reference.rsplit("!", 1)
        target_sheet = prefix.strip("'").replace("''", "'")
    else:
        target_sheet = sheet
    row_map = row_maps.get(target_sheet)
    if row_map is None or "[" in address:
        return reference

    parts = address.split(":")
    if len(parts) > 2:
        return reference
    matches = [REF_PART.match(part) for part in parts]
    if not all(matches) or not all(m.group("col") or m.group("row") for m in matches):
        return reference
    if not any(m.group("row") for m in matches):
        return reference  # whole columns such as A:A
    if len(parts) == 1 and not matches[0].group("col"):
        return reference

    own = own_row is not None and target_sheet == sheet
    new_parts = []
    for position, match in enumerate(matches):
        row = match.group("row")
        if row:
            row = int(row)
            if own and row == own_row and not match.group("row_abs"):
                row = row_map.targets(row)[index] if index < row_map.count(row) else row_map.first(row)
            elif position == len(matches) - 1 and len(matches) > 1:
                row = row_map.last(row)
            else:
                row = row_map.first(row)
        new_parts.append(f"{match.group('col') or ''}{match.group('row_abs') or ''}{row or ''}")
    return (prefix + "!" if prefix else "") + ":".join(new_parts)


def remap_range(cell_range, row_map):
    return CellRange(
        min_col=cell_range.min_col,
        max_col=cell_range.max_col,
        min_row=row_map.first(cell_range.min_row),
        max_row=row_map.last(cell_range.max_row),
    )


def remap_multi_range(multi_range, row_map):
    ranges = multi_range.ranges if hasattr(multi_range, "ranges") else MultiCellRange(str(multi_range)).ranges
    return " ".join(remap_range(cell_range, row_map).coord for cell_range in ranges)


def cell_label(ws, coordinate):
    return html.escape(f"{ws.title}!{coordinate}")
