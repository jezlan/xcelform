// Copyright (c) 2026, Muhamed Jezlan and contributors
// For license information, please see license.txt

frappe.ui.form.on("XF Report Excel Format", {
	setup(frm) {
		frm.set_query("report", () => ({
			filters: { report_type: ["in", ["Script Report", "Query Report", "Custom Report"]] },
		}));
	},

	refresh(frm) {
		const field = frm.get_field("template_file");
		if (field) {
			field.df.options = { restrictions: { allowed_file_types: [".xlsx"] } };
		}
		if (!frm.is_new() && frm.doc.report) {
			frm.add_custom_button(__("Open Report"), () => {
				frappe.set_route("query-report", frm.doc.report);
			});
		}
	},
});
