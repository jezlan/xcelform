# Copyright (c) 2025, Muhamed Jezlan and contributors
# For license information, please see license.txt

import frappe
from frappe.model.document import Document


class XFExcelFormat(Document):
    def validate(self):
        if self.is_default:
            existing_default = frappe.db.exists(
                "XF Excel Format",
                {
                    "doc_type": self.doc_type,
                    "is_default": 1,
                    "name": ["!=", self.name1],
                },
            )
            if existing_default:
                frappe.throw(
                    (
                        "There can only be one default Excel Format for Doctype {0}."
                    ).format(self.doc_type)
                )
