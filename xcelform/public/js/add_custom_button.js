const XF_EXPORT_METHOD = "xcelform.xcelform.hooks.xf_excel_utils.export_to_xf_excel";

$(document).on("app_ready", function () {
  frappe.call({
    method: "xcelform.xcelform.hooks.get_allowed_doctypes.get_doctypes",
    callback: function (r) {
      if (r.message) {
        $.each(r.message, function (i, doctype) {
          frappe.ui.form.on(doctype.doc_type, {
            refresh: function (frm) {
              if (frm.is_new()) return;
              frm.add_custom_button(__("XF Excel Export"), function () {
                show_excel_format_dialog(frm);
              });
            },
          });
        });
      }
    },
  });
});

function validate_excel_format(values) {
  if (!values.excel_format) {
    frappe.msgprint({
      title: __("Validation Error"),
      message: __("Select an Excel format!"),
      indicator: "red",
    });
    return false;
  }
  return true;
}

function show_excel_format_dialog(frm) {
  frappe.call({
    method: "xcelform.xcelform.hooks.xf_excel_utils.get_default_excel_format",
    args: {
      doctype: frm.doc.doctype,
    },
    callback: function (r) {
      let default_format = r.message || "";

      let d = new frappe.ui.Dialog({
        title: __("Select Excel Format"),
        fields: [
          {
            label: __("Excel Format"),
            fieldname: "excel_format",
            fieldtype: "Link",
            options: "XF Excel Format",
            default: default_format,
            get_query: () => ({
              filters: {
                doc_type: frm.doc.doctype,
                enable: 1,
              },
            }),
          },
          {
            label: __("Attach to Document"),
            fieldname: "attach",
            fieldtype: "Check",
          },
          {
            label: __("Download Excel"),
            fieldname: "download",
            fieldtype: "Check",
          },
        ],
        primary_action_label: __("Export"),
        primary_action(values) {
          if (!validate_excel_format(values)) return;
          if (!values.attach && !values.download) {
            frappe.msgprint({
              title: __("Validation Error"),
              message: __("Select either 'Attach to Document' or 'Download Excel'."),
              indicator: "red",
            });
            return;
          }

          const args = {
            doctype: frm.doc.doctype,
            doc_name: frm.doc.name,
            excel_format: values.excel_format,
          };

          if (values.attach) {
            frappe.call({
              method: XF_EXPORT_METHOD,
              args: { ...args, attach: 1 },
              freeze: true,
              callback: function (r) {
                if (r.message) {
                  frappe.show_alert({ message: __("Excel file attached"), indicator: "green" });
                  frm.reload_doc();
                  d.hide();
                }
              },
            });
          }

          if (values.download) {
            const params = new URLSearchParams({ ...args, download: 1 });
            window.open(frappe.urllib.get_full_url(`/api/method/${XF_EXPORT_METHOD}?${params.toString()}`));
            d.hide();
            frappe.show_alert({ message: __("Downloading Excel..."), indicator: "blue" });
          }
        },
        secondary_action_label: __("Preview"),
        secondary_action() {
          const values = d.get_values();
          if (!values || !validate_excel_format(values)) return;

          frappe.call({
            method: XF_EXPORT_METHOD,
            args: {
              doctype: frm.doc.doctype,
              doc_name: frm.doc.name,
              excel_format: values.excel_format,
              is_preview: 1,
            },
            freeze: true,
            callback: function (r) {
              if (r.message && r.message.preview) {
                // The server escapes every cell value before building this table.
                let dialog = new frappe.ui.Dialog({
                  title: __("Excel Preview"),
                  size: "extra-large",
                  fields: [
                    {
                      fieldtype: "HTML",
                      fieldname: "preview",
                      options: `<div style="overflow:auto; max-height:500px;">${r.message.preview}</div>`,
                    },
                  ],
                });
                dialog.show();
              }
            },
          });
        },
      });

      d.show();
    },
  });
}
