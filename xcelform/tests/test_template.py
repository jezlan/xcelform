# Copyright (c) 2026, Muhamed Jezlan and Contributors
# See license.txt

import datetime
import time
import zipfile
from io import BytesIO

import frappe
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

try:
    from frappe.tests import IntegrationTestCase
except ImportError:  # Frappe v15
    from frappe.tests.utils import FrappeTestCase as IntegrationTestCase

from xcelform.core.helpers import number_format, write_table
from xcelform.core.template import fill_template
from xcelform.core.template_check import check_template

YELLOW = "00FFFF00"


def invoice_template():
    """The invoice from the roadmap: header fields, a repeating item row and a total."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Invoice"
    ws["A1"] = "TAX INVOICE"
    ws["A1"].font = Font(bold=True, size=16)
    ws.merge_cells("A1:D1")
    ws["A2"], ws["B2"] = "Customer", "{{ doc.customer_name }}"
    ws["A3"], ws["B3"] = "Date", "{{ doc.posting_date }}"
    ws["B3"].number_format = "dd-mm-yyyy"
    ws["C3"], ws["D3"] = "No.", "{{ doc.name }}"
    for column, label in zip("ABCDE", ["Item", "Qty", "Rate", "Amount", "Check"], strict=True):
        ws[f"{column}4"] = label

    ws["A5"], ws["B5"], ws["C5"], ws["D5"] = (
        "{{ items.item_name }}",
        "{{ items.qty }}",
        "{{ items.rate }}",
        "{{ items.amount }}",
    )
    ws["E5"] = "=B5*C5-D5+$B$5*0"
    ws["A5"].fill = PatternFill("solid", start_color="FFFF00")
    ws["D5"].number_format = "#,##0.00"
    ws.row_dimensions[5].height = 22

    ws["A6"], ws["D6"] = "Total", "=SUM(D5:D5)"
    ws.merge_cells("A6:C6")
    ws.row_dimensions[6].height = 30
    ws["A7"] = "Thank you, {{ doc.customer_name }}. {{ doc.missing_field }}"

    ws.conditional_formatting.add("D5:D5", CellIsRule(operator="greaterThan", formula=["$D$6"], fill=PatternFill("solid", start_color="00FF00")))
    validation = DataValidation(type="list", formula1='"Paid,Unpaid"')
    ws.add_data_validation(validation)
    validation.add("E7")
    ws.freeze_panes = "A5"
    ws.print_area = "A1:E7"

    summary = wb.create_sheet("Summary")
    summary["A1"], summary["B1"] = "Invoice total", "=Invoice!D6"
    summary["A2"], summary["B2"] = "Items", "{{ items|length }}"
    summary["A3"], summary["B3"] = "All amounts", "=SUM(Invoice!D5:D5)"
    return save(wb)


def save(wb):
    output = BytesIO()
    wb.save(output)
    return output.getvalue()


def make_invoice(item_count):
    items = [
        frappe._dict(idx=i, item_name=f"Item {i}", qty=i, rate=10.5, amount=i * 10.5) for i in range(1, item_count + 1)
    ]
    doc = frappe._dict(
        name="ACC-SINV-0042",
        customer_name="Al Noor Trading",
        posting_date=datetime.date(2026, 11, 24),
        items=items,
    )
    return doc, items


def fill(content, doc, items):
    wb = fill_template(content, {"doc": doc}, {"items": items})
    return load_workbook(BytesIO(save(wb)))


def values(ws, max_col=5):
    return [[cell.value for cell in row] for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=max_col)]


class TestTemplateEngine(IntegrationTestCase):
    def setUp(self):
        frappe.set_user("Administrator")
        self.template = invoice_template()

    def test_invoice_with_three_items(self):
        doc, items = make_invoice(3)
        wb = fill(self.template, doc, items)
        ws = wb["Invoice"]

        self.assertEqual(
            values(ws),
            [
                ["TAX INVOICE", None, None, None, None],
                ["Customer", "Al Noor Trading", None, None, None],
                ["Date", datetime.datetime(2026, 11, 24), "No.", "ACC-SINV-0042", None],
                ["Item", "Qty", "Rate", "Amount", "Check"],
                ["Item 1", 1, 10.5, 10.5, "=B5*C5-D5+$B$5*0"],
                ["Item 2", 2, 10.5, 21, "=B6*C6-D6+$B$5*0"],
                ["Item 3", 3, 10.5, 31.5, "=B7*C7-D7+$B$5*0"],
                ["Total", None, None, "=SUM(D5:D7)", None],
                ["Thank you, Al Noor Trading. ", None, None, None, None],
            ],
        )
        self.assertEqual(ws["B3"].number_format, "dd-mm-yyyy")  # the designer's format is kept
        self.assertEqual({str(r) for r in ws.merged_cells.ranges}, {"A1:D1", "A8:C8"})
        self.assertEqual([ws.cell(row=r, column=1).fill.fgColor.rgb for r in (5, 6, 7, 8)], [YELLOW] * 3 + ["00000000"])
        self.assertEqual([ws.cell(row=r, column=4).number_format for r in (5, 6, 7)], ["#,##0.00"] * 3)
        self.assertEqual({r: d.height for r, d in ws.row_dimensions.items() if d.height}, {5: 22, 6: 22, 7: 22, 8: 30})

        rule_ranges = {str(cf.sqref): cf.rules[0].formula for cf in ws.conditional_formatting}
        self.assertEqual(rule_ranges, {"D5:D7": ["$D$8"]})
        self.assertEqual(str(ws.data_validations.dataValidation[0].sqref), "E9")
        self.assertEqual(ws.freeze_panes, "A5")
        self.assertEqual(ws.print_area, "'Invoice'!$A$1:$E$9")

        summary = wb["Summary"]
        self.assertEqual([summary["B1"].value, summary["B2"].value, summary["B3"].value], ["=Invoice!D8", 3, "=SUM(Invoice!D5:D7)"])

    def test_invoice_with_one_item_is_unchanged_in_layout(self):
        doc, items = make_invoice(1)
        ws = fill(self.template, doc, items)["Invoice"]
        self.assertEqual(ws["A5"].value, "Item 1")
        self.assertEqual(ws["D6"].value, "=SUM(D5:D5)")
        self.assertEqual({str(r) for r in ws.merged_cells.ranges}, {"A1:D1", "A6:C6"})

    def test_invoice_with_no_items_keeps_one_blank_row(self):
        doc, items = make_invoice(0)
        ws = fill(self.template, doc, items)["Invoice"]
        self.assertEqual(values(ws)[4], [None, None, None, None, "=B5*C5-D5+$B$5*0"])
        self.assertEqual(ws["D6"].value, "=SUM(D5:D5)")
        self.assertEqual(ws["A5"].fill.fgColor.rgb, YELLOW)

    def test_invoice_with_500_items(self):
        doc, items = make_invoice(500)
        ws = fill(self.template, doc, items)["Invoice"]
        self.assertEqual(ws["A504"].value, "Item 500")
        self.assertEqual(ws["E504"].value, "=B504*C504-D504+$B$5*0")
        self.assertEqual(ws["D505"].value, "=SUM(D5:D504)")
        self.assertIn("A505:C505", {str(r) for r in ws.merged_cells.ranges})

    def test_1000_rows_fill_quickly(self):
        doc, items = make_invoice(1000)
        start = time.monotonic()
        save(fill_template(self.template, {"doc": doc}, {"items": items}))
        self.assertLess(time.monotonic() - start, 3)

    def test_merges_and_images_in_moved_rows(self):
        wb = Workbook()
        ws = wb.active
        ws["A1"], ws["C1"] = "{{ items.item_name }}", "{{ items.qty }}"
        ws.merge_cells("A1:B1")
        ws["A3"] = "Signed"
        ws.merge_cells("A3:B4")
        image = Image(BytesIO(_png()))
        ws.add_image(image, "D3")

        doc, items = make_invoice(3)
        ws = fill(save(wb), doc, items).active
        self.assertEqual({str(r) for r in ws.merged_cells.ranges}, {"A1:B1", "A2:B2", "A3:B3", "A5:B6"})
        self.assertEqual(ws["A5"].value, "Signed")
        self.assertEqual(ws._images[0].anchor._from.row, 4)  # 0-based: row 5

    def test_general_cells_get_the_user_date_format(self):
        wb = Workbook()
        wb.active["A1"] = "{{ doc.posting_date }}"
        doc, items = make_invoice(0)
        ws = fill(save(wb), doc, items).active
        self.assertEqual(ws["A1"].number_format, number_format("Date"))

    def test_values_are_written_safely(self):
        wb = Workbook()
        ws = wb.active
        ws["A1"], ws["A2"], ws["A3"] = "{{ doc.formula }}", "{{ doc.html }}", "{{ doc.__class__ }}{{ ''.__class__.__mro__ }}"
        doc = frappe._dict(formula='=HYPERLINK("http://example.com","x")', html="<b>bold</b> & co")
        ws = fill(save(wb), doc, []).active

        self.assertEqual(ws["A1"].value, '=HYPERLINK("http://example.com","x")')
        self.assertEqual(ws["A1"].data_type, "s")  # text, not a live formula
        self.assertEqual(ws["A2"].value, "bold & co")
        self.assertIsNone(ws["A3"].value)  # the sandbox hides internals

    def test_report_table_writer_keeps_formula_like_text_as_text(self):
        wb = Workbook()
        write_table(wb.active, [{"fieldname": "note", "label": "Note", "fieldtype": "Data"}], [{"note": "=1+1"}])
        ws = load_workbook(BytesIO(save(wb))).active
        self.assertEqual((ws["A2"].value, ws["A2"].data_type), ("=1+1", "s"))

    def test_one_row_cannot_repeat_for_two_tables(self):
        wb = Workbook()
        wb.active["A1"], wb.active["B1"] = "{{ items.qty }}", "{{ taxes.rate }}"
        with self.assertRaises(frappe.ValidationError) as ctx:
            fill_template(save(wb), {"doc": {}}, {"items": [], "taxes": []})
        self.assertIn("one table", str(ctx.exception))

    def test_errors_name_the_cell(self):
        wb = Workbook()
        wb.active.title = "Main"
        wb.active["C2"] = "{{ doc.name ( }}"
        with self.assertRaises(frappe.ValidationError) as ctx:
            fill_template(save(wb), {"doc": {}})
        self.assertIn("Main!C2", str(ctx.exception))

        wb.active["C2"] = "{{ 1 / 0 }}"
        with self.assertRaises(frappe.ValidationError) as ctx:
            fill_template(save(wb), {"doc": {}})
        self.assertIn("Main!C2", str(ctx.exception))


class TestTemplateCheck(IntegrationTestCase):
    def setUp(self):
        frappe.set_user("Administrator")

    def test_summary_lists_placeholders_and_repeating_rows(self):
        summary = check_template("/private/files/invoice.xlsx", invoice_template(), ["items"])
        self.assertEqual(summary.placeholders, 9)
        self.assertEqual(summary.repeating, {"items": ["Invoice!5"]})
        self.assertEqual(summary.warnings, [])

    def test_unknown_fields_are_listed_by_cell(self):
        from xcelform.sources.document import field_checker

        wb = Workbook()
        wb.active["A1"] = "{{ doc.first_name }} {{ doc.frist_name }}"
        wb.active["A2"] = "{{ email_ids.email_id }}{{ email_ids.emial }}{{ doc.get_formatted('first_name') }}"
        with self.assertRaises(frappe.ValidationError) as ctx:
            check_template("t.xlsx", save(wb), ["email_ids"], field_checker("Contact"))
        message = str(ctx.exception)
        self.assertIn("Sheet!A1: doc.frist_name is not a field of Contact", message)
        self.assertIn("Sheet!A2: email_ids.emial is not a field of Contact Email", message)
        self.assertNotIn("first_name is not", message)
        self.assertNotIn("get_formatted", message)

    def test_wrong_file_types_are_rejected(self):
        for file_url in ("t.xls", "t.xlsm", "t.csv"):
            with self.subTest(file_url=file_url), self.assertRaises(frappe.ValidationError):
                check_template(file_url, invoice_template())
        with self.assertRaises(frappe.ValidationError):
            check_template("t.xlsx", b"not a zip file")

    def test_features_excel_drops_are_warned_about(self):
        output = BytesIO(invoice_template())
        with zipfile.ZipFile(output, "a") as archive:
            archive.writestr("xl/pivotTables/pivotTable1.xml", "<pivotTableDefinition/>")
            archive.writestr("xl/drawings/drawing9.xml", '<xdr:wsDr><xdr:sp macro=""/></xdr:wsDr>')
        summary = check_template("t.xlsx", output.getvalue())
        self.assertEqual(len(summary.warnings), 2)
        self.assertIn("pivot tables", summary.message)


def _png():
    from PIL import Image as PILImage

    output = BytesIO()
    PILImage.new("RGB", (20, 10), "red").save(output, format="PNG")
    return output.getvalue()
