import frappe


def execute():
    """Formats made before templates existed are all Python scripts."""
    for doctype in ("XF Excel Format", "XF Report Excel Format"):
        if not frappe.db.table_exists(doctype) or not frappe.db.has_column(doctype, "format_type"):
            continue
        frappe.db.set_value(
            doctype,
            {"template_file": ("is", "not set")},
            "format_type",
            "Script",
            update_modified=False,
        )
