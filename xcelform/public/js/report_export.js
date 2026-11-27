// XcelForm: "XF Excel Export" menu item on query and script reports.
//
// Frappe uses one shared report page for every report and de-duplicates menu items by
// label, so the item is added once and then shown or hidden each time a report loads,
// depending on whether that report has an enabled XF Report Excel Format.

frappe.provide("xcelform.report_export");

xcelform.report_export = {
	API: "xcelform.api.report_export",

	label() {
		return __("XF Excel Export");
	},

	patch() {
		const QueryReport = frappe.views && frappe.views.QueryReport;
		if (!QueryReport || QueryReport.prototype.__xf_patched) return;
		QueryReport.prototype.__xf_patched = true;

		const load_report = QueryReport.prototype.load_report;
		QueryReport.prototype.load_report = function (...args) {
			const result = load_report.apply(this, args);
			xcelform.report_export.refresh_menu(this);
			return result;
		};

		// A report may already be open if the desk started on a report URL.
		if (frappe.query_report && frappe.query_report.report_name) {
			this.refresh_menu(frappe.query_report);
		}
	},

	find_menu_item(report) {
		const label = this.label();
		return report.page.menu
			.find("li > a.grey-link > span")
			.filter((i, el) => $(el).text().trim() === label)
			.closest("li");
	},

	refresh_menu(report) {
		const report_name = report.report_name;
		this.find_menu_item(report).hide();

		frappe
			.xcall(`${this.API}.get_report_formats`, { report_name })
			.then((formats) => {
				if (report.report_name !== report_name || !formats || !formats.length) return;
				if (!this.find_menu_item(report).length) {
					report.page.add_menu_item(this.label(), () => this.show_dialog(frappe.query_report));
				}
				this.find_menu_item(report).show();
			});
	},

	show_dialog(report) {
		const report_name = report.report_name;
		let filters;
		try {
			filters = report.get_filter_values(true);
		} catch (e) {
			frappe.show_alert({ message: __("Set the required filters first"), indicator: "orange" });
			return;
		}

		frappe.xcall(`${this.API}.get_report_formats`, { report_name }).then((formats) => {
			if (!formats.length) {
				frappe.msgprint(__("No enabled Excel format exists for this report."));
				return;
			}

			const dialog = new frappe.ui.Dialog({
				title: __("Export {0} to Excel", [__(report_name)]),
				fields: [
					{
						label: __("Excel Format"),
						fieldname: "excel_format",
						fieldtype: "Select",
						options: formats.map((f) => f.name),
						default: formats[0].name,
						reqd: 1,
					},
					{
						fieldtype: "HTML",
						options: `<p class="text-muted small">${__(
							"The export uses the filters currently applied to the report."
						)}</p>`,
					},
				],
				primary_action_label: __("Download"),
				primary_action: (values) => {
					dialog.hide();
					this.download(report_name, filters, values.excel_format);
				},
				secondary_action_label: __("Preview"),
				secondary_action: () => {
					const values = dialog.get_values();
					if (values) this.preview(report_name, filters, values.excel_format);
				},
			});
			dialog.show();
		});
	},

	preview(report_name, filters, excel_format) {
		frappe.call({
			method: `${this.API}.export_report`,
			args: { report_name, filters, excel_format, is_preview: 1 },
			freeze: true,
			freeze_message: __("Building preview..."),
			callback: (r) => {
				if (!r.message || !r.message.preview) return;
				// The server escapes every cell value before building this table.
				new frappe.ui.Dialog({
					title: __("Excel Preview"),
					size: "extra-large",
					fields: [
						{
							fieldtype: "HTML",
							options: `<div style="overflow:auto; max-height:500px;">${r.message.preview}</div>`,
						},
					],
				}).show();
			},
		});
	},

	download(report_name, filters, excel_format) {
		// POST, so long filter sets are not cut off by URL limits, and fetch() rather than a
		// form submit, so a server error shows as a message instead of a raw JSON page.
		const body = new URLSearchParams({
			report_name,
			filters: JSON.stringify(filters || {}),
			excel_format,
		});

		frappe.dom.freeze(__("Preparing Excel..."));
		fetch(`/api/method/${this.API}.export_report`, {
			method: "POST",
			headers: { "X-Frappe-CSRF-Token": frappe.csrf_token, Accept: "application/json" },
			body,
		})
			.then(async (response) => {
				const type = response.headers.get("Content-Type") || "";
				if (response.ok && !type.includes("application/json")) {
					const blob = await response.blob();
					this.save_blob(blob, this.file_name(response, report_name));
					frappe.show_alert({ message: __("Excel downloaded"), indicator: "green" });
					return;
				}
				this.show_server_error(await response.json().catch(() => ({})));
			})
			.catch(() => {
				frappe.msgprint({
					title: __("Export failed"),
					message: __("Could not reach the server. Check your connection and try again."),
					indicator: "red",
				});
			})
			.finally(() => frappe.dom.unfreeze());
	},

	file_name(response, report_name) {
		const header = response.headers.get("Content-Disposition") || "";
		const encoded = header.match(/filename\*=UTF-8''([^;]+)/i);
		if (encoded) return decodeURIComponent(encoded[1]);
		const plain = header.match(/filename="?([^";]+)"?/i);
		return plain ? plain[1] : `${report_name}.xlsx`;
	},

	save_blob(blob, file_name) {
		const url = URL.createObjectURL(blob);
		const link = document.createElement("a");
		link.href = url;
		link.download = file_name;
		document.body.appendChild(link);
		link.click();
		link.remove();
		setTimeout(() => URL.revokeObjectURL(url), 1000);
	},

	show_server_error(data) {
		let messages = [];
		try {
			messages = JSON.parse(data._server_messages || "[]").map((m) => JSON.parse(m));
		} catch (e) {
			messages = [];
		}

		if (messages.length) {
			messages.forEach((m) => frappe.msgprint(m));
		} else {
			frappe.msgprint({
				title: __("Export failed"),
				message: frappe.utils.escape_html(data.exception || __("The server could not build the Excel file.")),
				indicator: "red",
			});
		}
	},
};

$(document).on("app_ready", () => xcelform.report_export.patch());
