# Copyright (c) 2025, Muhamed Jezlan and Contributors
# See license.txt

import re
from unittest.mock import patch

import frappe

try:
    from frappe.tests import IntegrationTestCase
except ImportError:  # Frappe v15
    from frappe.tests.utils import FrappeTestCase as IntegrationTestCase

from xcelform.xcelform.hooks.xf_excel_utils import (
    export_to_xf_excel,
    get_default_excel_format,
    run_format_script,
)

FORMAT_NAME = "_Test XF ToDo Format"
FORMAT_SCRIPT = """
from openpyxl.styles import Font
ws["A1"] = "<script>alert('xf')</script>"
ws["A1"].font = Font(bold=True)
ws.merge_cells("A2:C2")
ws["A2"] = doc.name
"""


class TestXFExcelFormat(IntegrationTestCase):
    def setUp(self):
        self.enterContext(patch("frappe.utils.safe_exec.is_safe_exec_enabled", return_value=True))
        frappe.set_user("Administrator")
        self.todo = frappe.get_doc({"doctype": "ToDo", "description": "XcelForm test"}).insert()
        self.format = frappe.get_doc(
            {
                "doctype": "XF Excel Format",
                "name1": FORMAT_NAME,
                "doc_type": "ToDo",
                "enable": 1,
                "is_default": 1,
                "custom_format": 1,
                "python_script": FORMAT_SCRIPT,
            }
        ).insert()

    def tearDown(self):
        frappe.set_user("Administrator")
        frappe.db.rollback()

    def preview(self, **kwargs):
        return export_to_xf_excel("ToDo", self.todo.name, FORMAT_NAME, is_preview=1, **kwargs)["preview"]

    def test_preview_escapes_cell_values(self):
        html = self.preview()
        self.assertNotIn("<script>", html)
        self.assertIn("&lt;script&gt;alert(&#x27;xf&#x27;)&lt;/script&gt;", html)

    def test_preview_renders_merged_cells_once(self):
        rows = re.findall(r"<tr>(.*?)</tr>", self.preview())
        self.assertIn('colspan="3"', rows[1])
        self.assertEqual(rows[1].count("<td"), 1)
        self.assertIn(self.todo.name, rows[1])

    def test_export_requires_read_permission(self):
        frappe.set_user("Guest")
        with self.assertRaises(frappe.PermissionError):
            self.preview()

    def test_export_requires_an_output_mode(self):
        with self.assertRaises(frappe.ValidationError):
            export_to_xf_excel("ToDo", self.todo.name, FORMAT_NAME)

    def test_disabled_format_is_rejected(self):
        self.format.db_set("enable", 0)
        with self.assertRaises(frappe.ValidationError):
            self.preview()
        self.assertIsNone(get_default_excel_format("ToDo"))

    def test_default_format_is_used_when_none_given(self):
        self.assertEqual(get_default_excel_format("ToDo"), FORMAT_NAME)
        html = export_to_xf_excel("ToDo", self.todo.name, is_preview=1)["preview"]
        self.assertIn(self.todo.name, html)

    def test_scripts_can_import_only_openpyxl_names(self):
        for script in ("import os", "from os import path", "import openpyxl", "from openpyxl import styles"):
            with self.subTest(script=script), self.assertRaises(frappe.ValidationError):
                run_format_script(script, {})

        wb = run_format_script("from openpyxl.styles import Font\nws['A1'] = 'ok'", {})
        self.assertEqual(wb.active["A1"].value, "ok")

    def test_scripts_cannot_reach_private_attributes(self):
        with self.assertRaises(frappe.ValidationError):
            run_format_script("ws['A1'] = wb.__class__", {})

    def test_clear_error_when_server_scripts_are_disabled(self):
        with patch("frappe.utils.safe_exec.is_safe_exec_enabled", return_value=False):
            with self.assertRaises(frappe.ValidationError) as ctx:
                run_format_script("ws['A1'] = 1", {})
        self.assertIn("server_script_enabled", str(ctx.exception))
