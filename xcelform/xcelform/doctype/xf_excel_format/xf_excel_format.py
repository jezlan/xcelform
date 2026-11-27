# Copyright (c) 2025, Muhamed Jezlan and contributors
# For license information, please see license.txt

import frappe
from frappe import _
from frappe.model.document import Document

from xcelform.core.pipeline import validate_format
from xcelform.sources.document import field_checker, get_table_names


class XFExcelFormat(Document):
    def validate(self):
        self.validate_single_default()
        validate_format(self, get_table_names(self.doc_type), field_checker(self.doc_type))

    def validate_single_default(self):
        if not self.is_default:
            return
        other_default = frappe.db.get_value(
            "XF Excel Format",
            {"doc_type": self.doc_type, "is_default": 1, "name": ("!=", self.name)},
            "name",
        )
        if other_default:
            frappe.throw(
                _("{0} is already the default Excel Format for {1}. Untick it there first.").format(
                    frappe.bold(other_default), frappe.bold(self.doc_type)
                )
            )
