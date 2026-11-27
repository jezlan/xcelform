# Copyright (c) 2026, Muhamed Jezlan and contributors
# For license information, please see license.txt

import frappe
from frappe import _
from frappe.model.document import Document

from xcelform.core.pipeline import validate_format
from xcelform.sources.report import SUPPORTED_REPORT_TYPES


class XFReportExcelFormat(Document):
	def validate(self):
		self.validate_report_type()
		self.validate_single_default()
		validate_format(self, repeat_names=("data",))

	def validate_report_type(self):
		report_type = frappe.db.get_value("Report", self.report, "report_type")
		if report_type not in SUPPORTED_REPORT_TYPES:
			frappe.throw(
				_("{0} is a {1}. Excel formats work with: {2}.").format(
					frappe.bold(self.report), report_type, ", ".join(SUPPORTED_REPORT_TYPES)
				)
			)

	def validate_single_default(self):
		if not self.is_default:
			return
		other_default = frappe.db.get_value(
			"XF Report Excel Format",
			{"report": self.report, "is_default": 1, "name": ("!=", self.name)},
			"name",
		)
		if other_default:
			frappe.throw(
				_("{0} is already the default format for {1}. Untick it there first.").format(
					frappe.bold(other_default), frappe.bold(self.report)
				)
			)
