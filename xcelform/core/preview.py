"""Turns a generated workbook into an HTML table for the in-app preview."""

import html
from io import BytesIO

from openpyxl import load_workbook

ALLOWED_TEXT_ALIGN = {"left", "center", "right", "justify"}


def rgb_to_hex(rgb):
    """Convert an Excel ARGB colour to a CSS hex colour, falling back to white."""
    if not isinstance(rgb, str) or len(rgb) < 6 or rgb == "00000000":
        return "#FFFFFF"
    hex_part = rgb[-6:]
    if any(c not in "0123456789abcdefABCDEF" for c in hex_part):
        return "#FFFFFF"
    return f"#{hex_part}"


def generate_excel_preview(output):
    """Render the active sheet of a saved workbook as an HTML table. All cell content is escaped."""
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
