$(document).on("app_ready", function () {
  frappe.call({
    method: "xcelform.xcelform.hooks.get_allowed_doctypes.get_doctypes",
    callback: function (r) {
      if (r.message) {
        $.each(r.message, function (i, doctype) {
          frappe.ui.form.on(doctype.doc_type, {
            setup: function (frm) {
              frm.add_custom_button(__("XF Excel Export"), function () {
                show_excel_format_dialog(frm);
              });
            },
            refresh: function (frm) {
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
          if (!values.excel_format) {
            frappe.msgprint({
              title: __("Validation Error"),
              message: __("Select an Excel format!"),
              indicator: "red",
            });
            return;
          }
          if (!values.attach && !values.download) {
            frappe.msgprint({
              title: __("Validation Error"),
              message: __(
                "Select either 'Attach to Document' or 'Download Excel'."
              ),
              indicator: "red",
            });
            return;
          }

          const args = {
            doctype: frm.doc.doctype,
            doc_name: frm.doc.name,
            excel_format: values.excel_format,
          };

          // Attach option
          if (values.attach) {
            args.attach = 1;
            frappe.call({
              method:
                "xcelform.xcelform.hooks.xf_excel_utils.export_to_xf_excel",
              args: args,
              callback: function (r) {
                if (r.message) {
                  frappe.msgprint(__("Excel file attached successfully!"));
                  d.hide();
                }
              },
            });
          }

          // Download option
          if (values.download) {
            frappe.msgprint(__("Downloading Excel..."));
            window.open(
              frappe.urllib.get_full_url(
                `/api/method/xcelform.xcelform.hooks.xf_excel_utils.export_to_xf_excel?doctype=${frm.doc.doctype}&doc_name=${frm.doc.name}&excel_format=${values.excel_format}&download=1`
              )
            );
            d.hide();
            frappe.show_alert({
              message: __("Excel downloaded successfully!"),
              indicator: "blue",
            });
          }
        },
        secondary_action_label: __("Preview"),
        secondary_action(values) {
          console.log("hiii")
          frappe.call({
            method: "xcelform.xcelform.hooks.xf_excel_utils.export_to_xf_excel",
            args: {
                doctype: frm.doc.doctype,
                doc_name: frm.doc.name,
                excel_format: values.excel_format,
                is_preview: 1
            },
            callback: function (r) {
              if (r.message) {
                let previewHtml = r.message.preview;
                let dialog = new frappe.ui.Dialog({
                    title: "Excel Preview",
                    fields: [
                        {
                            fieldtype: "HTML",
                            fieldname: "preview",
                            options: `<div style="overflow:auto; max-height:500px;">${previewHtml}</div>`
                        }
                    ]
                });
                dialog.show();
            }
            }
        });
        }
      });

      d.show();
    },
  });
}
