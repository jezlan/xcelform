// Copyright (c) 2025, Muhamed Jezlan and contributors
// For license information, please see license.txt

frappe.ui.form.on("XF Excel Format", {
	setup(frm) {
		frm.set_query("doc_type", () => ({ filters: { istable: 0, issingle: 0 } }));
	},

	refresh(frm) {
		const field = frm.get_field("template_file");
		if (field) {
			field.df.options = { restrictions: { allowed_file_types: [".xlsx"] } };
		}
		frm.events.add_starter_button(frm);
	},

	doc_type(frm) {
		frm.events.add_starter_button(frm);
	},

	add_starter_button(frm) {
		const label = __("Download Starter Template");
		frm.remove_custom_button(label);
		if (!frm.doc.doc_type) return;
		frm.add_custom_button(label, () => {
			const params = new URLSearchParams({ doctype: frm.doc.doc_type });
			window.open(`/api/method/xcelform.xcelform.hooks.xf_excel_utils.download_starter_template?${params}`);
		});
	},
});
