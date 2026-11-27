"""Documents as a data source: permission check, child tables and field names."""

import frappe
from frappe import _
from frappe.model.document import Document

# Attributes every document has, in addition to its fields.
DOCUMENT_ATTRIBUTES = {"doctype", "name", "idx", "parent", "parentfield", "parenttype"}


def get_document(doctype, name, for_write=False):
    doc = frappe.get_doc(doctype, name)
    doc.check_permission("read")
    if for_write:
        doc.check_permission("write")
    return doc


def get_tables(doc):
    """Child tables by fieldname, such as {"items": [...]}. These rows repeat in a template."""
    return {df.fieldname: doc.get(df.fieldname) or [] for df in doc.meta.get_table_fields()}


def get_table_names(doctype):
    return [df.fieldname for df in frappe.get_meta(doctype).get_table_fields()]


def field_checker(doctype):
    """Return a function that explains a `doc.x` or `items.x` placeholder naming no real field."""
    meta = frappe.get_meta(doctype)
    known = {"doc": (doctype, _fields(meta))}
    for df in meta.get_table_fields():
        known[df.fieldname] = (df.options, _fields(frappe.get_meta(df.options)))

    def check(name, attribute):
        if name not in known:
            return None
        label, fields = known[name]
        if attribute in fields or hasattr(Document, attribute):
            return None
        return _("{0}.{1} is not a field of {2}").format(name, attribute, label)

    return check


def _fields(meta):
    return set(meta.get_valid_columns()) | {df.fieldname for df in meta.fields} | DOCUMENT_ATTRIBUTES
