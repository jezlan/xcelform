# Copyright (c) 2025, Muhamed Jezlan and Contributors
# See license.txt

import os
import re
from io import BytesIO
from unittest.mock import patch

import frappe
from openpyxl import Workbook, load_workbook

try:
    from frappe.tests import IntegrationTestCase
except ImportError:  # Frappe v15
    from frappe.tests.utils import FrappeTestCase as IntegrationTestCase

from xcelform.core.sandbox import run_format_script
from xcelform.xcelform.hooks.xf_excel_utils import (
    download_starter_template,
    export_to_xf_excel,
    get_default_excel_format,
)

FORMAT_NAME = "_Test XF ToDo Format"
FORMAT_SCRIPT = """
from openpyxl.styles import Font
ws["A1"] = "<script>alert('xf')</script>"
ws["A1"].font = Font(bold=True)
ws.merge_cells("A2:C2")
ws["A2"] = doc.name
"""

# Only this test's own records are needed. Following the DocType link would make Frappe
# create sample records from every installed app (ERPNext, HRMS...) and commit them.
IGNORE_TEST_RECORD_DEPENDENCIES = ["DocType"]


class TestXFExcelFormat(IntegrationTestCase):
    def setUp(self):
        self.enterContext(patch("frappe.utils.safe_exec.is_safe_exec_enabled", return_value=True))
        self.enterContext(patch("frappe.log_error"))  # scripts here fail on purpose
        frappe.set_user("Administrator")
        self.todo = frappe.get_doc({"doctype": "ToDo", "description": "XcelForm test"}).insert()
        self.format = frappe.get_doc(
            {
                "doctype": "XF Excel Format",
                "name1": FORMAT_NAME,
                "doc_type": "ToDo",
                "enable": 1,
                "is_default": 1,
                "format_type": "Script",
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


TEMPLATE_FORMAT = "_Test XF Contact Template"
TEMPLATE_FILE = "_test_xf_contact_template.xlsx"


class TestXFExcelFormatTemplate(IntegrationTestCase):
    """Template formats on Contact, which has a child table of email addresses."""

    def setUp(self):
        frappe.set_user("Administrator")
        self.contact = frappe.get_doc(
            {
                "doctype": "Contact",
                "first_name": "_Test XF",
                "last_name": "Contact",
                "email_ids": [{"email_id": "one@example.com", "is_primary": 1}, {"email_id": "two@example.com"}],
            }
        ).insert()

    def tearDown(self):
        frappe.set_user("Administrator")
        frappe.db.rollback()  # Frappe moves or deletes uploaded files on rollback; remove what's left
        for folder in ("private", "public"):
            path = frappe.get_site_path(folder, "files", TEMPLATE_FILE)
            if os.path.exists(path):
                os.remove(path)

    def upload(self, wb, is_private=1):
        output = BytesIO()
        wb.save(output)
        return frappe.get_doc(
            {"doctype": "File", "file_name": TEMPLATE_FILE, "is_private": is_private, "content": output.getvalue()}
        ).insert()

    def make_format(self, file_url, **kwargs):
        return frappe.get_doc(
            {
                "doctype": "XF Excel Format",
                "name1": TEMPLATE_FORMAT,
                "doc_type": "Contact",
                "enable": 1,
                "format_type": "Template",
                "template_file": file_url,
                **kwargs,
            }
        ).insert()

    def contact_template(self):
        wb = Workbook()
        ws = wb.active
        ws["A1"] = "{{ doc.first_name }} {{ doc.last_name }}"
        ws["A3"], ws["B3"] = "{{ email_ids.idx }}", "{{ email_ids.email_id }}"
        ws["A4"], ws["B4"] = "Addresses", "=COUNTA(B3:B3)"
        return wb

    def download(self):
        frappe.local.response = frappe._dict()
        export_to_xf_excel("Contact", self.contact.name, TEMPLATE_FORMAT, download=1)
        return frappe.local.response

    def test_template_download_repeats_child_rows(self):
        self.make_format(self.upload(self.contact_template()).file_url, file_name_pattern="{{ doc.first_name }} card")
        response = self.download()

        self.assertEqual(response.filename, "_Test XF card.xlsx")
        ws = load_workbook(BytesIO(response.filecontent)).active
        self.assertEqual(ws["A1"].value, "_Test XF Contact")
        self.assertEqual([[c.value for c in row] for row in ws["A3:B5"]], [
            [1, "one@example.com"],
            [2, "two@example.com"],
            ["Addresses", "=COUNTA(B3:B4)"],
        ])

    def test_template_can_be_attached_and_previewed(self):
        self.make_format(self.upload(self.contact_template()).file_url)
        html = export_to_xf_excel("Contact", self.contact.name, TEMPLATE_FORMAT, is_preview=1)["preview"]
        self.assertIn("two@example.com", html)

        file_url = export_to_xf_excel("Contact", self.contact.name, TEMPLATE_FORMAT, attach=1)
        attached = frappe.get_doc("File", {"file_url": file_url})
        self.assertEqual((attached.attached_to_doctype, attached.is_private), ("Contact", 1))
        self.assertTrue(attached.file_name.startswith(f"XF_Excel_Contact_{self.contact.name}"))

    def test_post_process_script_runs_after_filling(self):
        with patch("frappe.utils.safe_exec.is_safe_exec_enabled", return_value=True):
            self.make_format(self.upload(self.contact_template()).file_url, post_process_script="ws['D1'] = doc.last_name")
            ws = load_workbook(BytesIO(self.download().filecontent)).active
        self.assertEqual(ws["D1"].value, "Contact")

    def test_template_mode_works_without_server_scripts(self):
        with patch("frappe.utils.safe_exec.is_safe_exec_enabled", return_value=False):
            self.make_format(self.upload(self.contact_template()).file_url)
            self.assertEqual(self.download().type, "binary")

    def test_unknown_field_is_rejected_on_save(self):
        wb = self.contact_template()
        wb.active["A2"] = "{{ doc.frist_name }}"
        with self.assertRaises(frappe.ValidationError) as ctx:
            self.make_format(self.upload(wb).file_url)
        self.assertIn("Sheet!A2: doc.frist_name is not a field of Contact", str(ctx.exception))

    def test_public_upload_is_made_private(self):
        public = self.upload(self.contact_template(), is_private=0)
        fmt = self.make_format(public.file_url)
        self.assertTrue(fmt.template_file.startswith("/private/files/"))
        self.assertTrue(frappe.db.get_value("File", public.name, "is_private"))

    def test_template_needs_a_file(self):
        with self.assertRaises(frappe.ValidationError):
            self.make_format(None)

    def test_starter_template_has_placeholders_for_fields_and_tables(self):
        frappe.local.response = frappe._dict()
        download_starter_template("Contact")
        ws = load_workbook(BytesIO(frappe.local.response.filecontent)).active
        cells = {c.value for row in ws.iter_rows() for c in row if isinstance(c.value, str)}

        self.assertIn("{{ doc.first_name }}", cells)
        self.assertIn("{{ email_ids.email_id }}", cells)

        # The starter file is itself a valid template for the doctype.
        self.make_format(self.upload(load_workbook(BytesIO(frappe.local.response.filecontent))).file_url)
        self.assertEqual(self.download().type, "binary")

    def test_starter_template_needs_format_permission(self):
        frappe.set_user("Guest")
        with self.assertRaises(frappe.PermissionError):
            download_starter_template("Contact")
