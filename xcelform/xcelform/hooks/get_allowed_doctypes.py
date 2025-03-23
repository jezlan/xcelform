import frappe


@frappe.whitelist()
def get_doctypes():
    settings = frappe.get_single("XF Settings")
    if settings:
        return settings.get("allowed_doctypes") or []
    return []
