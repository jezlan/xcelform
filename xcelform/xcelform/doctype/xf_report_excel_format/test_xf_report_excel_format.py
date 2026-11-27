# Copyright (c) 2026, Muhamed Jezlan and Contributors
# See license.txt

import os
from io import BytesIO
from unittest.mock import patch

import frappe
from openpyxl import Workbook, load_workbook

try:
    from frappe.tests import IntegrationTestCase
except ImportError:  # Frappe v15
    from frappe.tests.utils import FrappeTestCase as IntegrationTestCase

from xcelform.api.report_export import export_report, get_report_formats
from xcelform.core.helpers import number_format
from xcelform.sources.report import get_report_data

QUERY_REPORT = "_Test XF ToDo Query"
FORMAT_NAME = "_Test XF Query Format"
QUERY = """
select
    name as `ID:Link/ToDo:120`,
    description as `Description:Data:200`,
    priority as `Priority::80`,
    1 as `Count:Int:60`,
    date as `Date:Date:100`
from `tabToDo`
where description like '%%_Test XF row%%'
order by description
"""
SCRIPT = """
ws["A1"] = report_name
xf.write_table(ws, columns, data, start_row=3, total_row=total_row)
xf.autofit_columns(ws)
"""

# Only this test's own records are needed. Following the Report link would make Frappe
# create sample records from every installed app (ERPNext, HRMS...) and commit them.
IGNORE_TEST_RECORD_DEPENDENCIES = ["Report"]


class TestXFReportExcelFormat(IntegrationTestCase):
    def setUp(self):
        self.enterContext(patch("frappe.utils.safe_exec.is_safe_exec_enabled", return_value=True))
        self.enterContext(patch("frappe.log_error"))  # scripts here fail on purpose
        frappe.set_user("Administrator")

        self.todos = [
            frappe.get_doc(
                {"doctype": "ToDo", "description": description, "date": "2026-11-20", "priority": "High"}
            ).insert()
            for description in ("_Test XF row one", "_Test XF row <b>two</b>")
        ]
        self.report = frappe.get_doc(
            {
                "doctype": "Report",
                "report_name": QUERY_REPORT,
                "ref_doctype": "ToDo",
                "report_type": "Query Report",
                "is_standard": "No",
                "module": "Desk",
                "add_total_row": 1,
                "query": QUERY,
            }
        ).insert()
        self.format = self.make_format(FORMAT_NAME, is_default=1)

    def tearDown(self):
        frappe.set_user("Administrator")
        frappe.db.rollback()

    def make_format(self, name, report=QUERY_REPORT, is_default=0, enabled=1, script=SCRIPT, **kwargs):
        return frappe.get_doc(
            {
                "doctype": "XF Report Excel Format",
                "name1": name,
                "report": report,
                "enabled": enabled,
                "is_default": is_default,
                "format_type": "Script",
                "python_script": script,
                **kwargs,
            }
        ).insert()

    def download(self, **kwargs):
        frappe.local.response = frappe._dict()
        export_report(QUERY_REPORT, filters="{}", **kwargs)
        return frappe.local.response

    # Report data

    def test_query_report_columns_and_rows_are_normalised(self):
        report = get_report_data(QUERY_REPORT, "{}")

        self.assertEqual([c["fieldname"] for c in report.columns], ["id", "description", "priority", "count", "date"])
        self.assertEqual(report.columns[0]["fieldtype"], "Link")
        self.assertEqual(report.columns[0]["options"], "ToDo")
        self.assertEqual(report.columns[3]["fieldtype"], "Int")
        self.assertCountEqual([row["id"] for row in report.data], [t.name for t in self.todos])

    def test_total_row_is_separated_from_data(self):
        report = get_report_data(QUERY_REPORT, "{}")

        self.assertEqual(len(report.data), 2)
        self.assertIsInstance(report.total_row, dict)
        self.assertEqual(report.total_row["count"], 2)

    def test_script_report_with_text_columns_and_list_rows(self):
        report = get_report_data("ToDo", {})

        self.assertEqual(report.columns[0]["fieldname"], "id")
        self.assertEqual(report.columns[2]["fieldtype"], "Date")
        ids = {row["id"] for row in report.data}
        self.assertTrue({t.name for t in self.todos} <= ids)

    def test_prepared_report_without_result_gives_clear_message(self):
        self.report.db_set("prepared_report", 1)
        with self.assertRaises(frappe.ValidationError) as ctx:
            get_report_data(QUERY_REPORT, {})
        self.assertIn("runs in the background", str(ctx.exception))

    # Export

    def test_download_builds_formatted_workbook(self):
        response = self.download(excel_format=FORMAT_NAME)

        self.assertEqual(response.type, "binary")
        self.assertTrue(response.filename.startswith(QUERY_REPORT) and response.filename.endswith(".xlsx"))

        ws = load_workbook(BytesIO(response.filecontent)).active
        self.assertEqual(ws["A1"].value, QUERY_REPORT)
        self.assertEqual([ws.cell(row=3, column=c).value for c in range(1, 6)], ["ID", "Description", "Priority", "Count", "Date"])
        self.assertTrue(ws["A3"].font.bold)
        # ToDo names are random, so don't rely on which row comes first. HTML tags are stripped.
        self.assertCountEqual([ws["B4"].value, ws["B5"].value], ["_Test XF row one", "_Test XF row two"])
        self.assertEqual(ws["E4"].number_format, number_format("Date"))
        self.assertEqual(ws["D6"].value, 2)  # total row
        self.assertTrue(ws["D6"].font.bold)

    def test_default_format_is_used_when_none_given(self):
        response = self.download()
        self.assertEqual(response.type, "binary")

    def test_preview_returns_escaped_html(self):
        html = export_report(QUERY_REPORT, "{}", FORMAT_NAME, is_preview=1)["preview"]
        self.assertIn(self.todos[0].name, html)
        self.assertNotIn("<b>", html)

    def test_user_without_report_access_cannot_export(self):
        frappe.set_user("Guest")
        self.assertEqual(get_report_formats(QUERY_REPORT), [])
        with self.assertRaises(frappe.PermissionError):
            export_report(QUERY_REPORT, "{}", FORMAT_NAME)

    def test_disabled_format_is_not_offered_or_used(self):
        self.make_format("_Test XF Second Format")
        self.make_format("_Test XF Disabled Format", enabled=0)

        self.assertEqual([f.name for f in get_report_formats(QUERY_REPORT)], [FORMAT_NAME, "_Test XF Second Format"])
        with self.assertRaises(frappe.ValidationError):
            export_report(QUERY_REPORT, "{}", "_Test XF Disabled Format")

    def test_missing_default_gives_clear_message(self):
        self.format.db_set("is_default", 0)
        with self.assertRaises(frappe.ValidationError) as ctx:
            export_report(QUERY_REPORT, "{}")
        self.assertIn("No enabled default", str(ctx.exception))

    # Format validation

    def test_only_one_default_per_report(self):
        with self.assertRaises(frappe.ValidationError):
            self.make_format("_Test XF Another Default", is_default=1)

    def test_report_builder_reports_are_rejected(self):
        with self.assertRaises(frappe.ValidationError):
            self.make_format("_Test XF Builder Format", report="Document Share Report")

    def test_template_format_repeats_report_rows(self):
        wb = Workbook()
        ws = wb.active
        ws["A1"], ws["B1"] = "{{ report_name }}", "{{ data|length }} rows"
        ws["A3"], ws["B3"], ws["C3"] = "{{ data.description }}", "{{ data.count }}", "{{ data.date }}"
        ws["C3"].number_format = "dd/mm/yyyy"
        ws["A4"], ws["B4"], ws["C4"] = "Total", "=SUM(B3:B3)", "{{ total_row.count }}"
        output = BytesIO()
        wb.save(output)
        template = frappe.get_doc(
            {"doctype": "File", "file_name": "_test_xf_report_template.xlsx", "is_private": 1, "content": output.getvalue()}
        ).insert()
        self.addCleanup(lambda: os.path.exists(template.get_full_path()) and os.remove(template.get_full_path()))

        self.make_format(
            "_Test XF Template Format",
            format_type="Template",
            template_file=template.file_url,
            python_script=None,
            file_name_pattern="{{ report_name }} export",
        )
        response = self.download(excel_format="_Test XF Template Format")

        self.assertEqual(response.filename, f"{QUERY_REPORT} export.xlsx")
        ws = load_workbook(BytesIO(response.filecontent)).active
        self.assertEqual([ws["A1"].value, ws["B1"].value], [QUERY_REPORT, "2 rows"])
        self.assertCountEqual([ws["A3"].value, ws["A4"].value], ["_Test XF row one", "_Test XF row two"])
        self.assertEqual([ws["B3"].value, ws["C3"].number_format], [1, "dd/mm/yyyy"])
        self.assertEqual([ws["A5"].value, ws["B5"].value, ws["C5"].value], ["Total", "=SUM(B3:B4)", 2])

    def test_script_syntax_is_checked_on_save(self):
        with self.assertRaises(frappe.ValidationError) as ctx:
            self.make_format("_Test XF Broken Format", script="ws['A1'] = (")
        self.assertIn("line 1", str(ctx.exception).lower())
