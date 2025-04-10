import frappe
import io
from openpyxl import load_workbook

@frappe.whitelist()
def export_quotation_to_excel_template(quotation_name):
    quotation = frappe.get_doc("Quotation", quotation_name)
    customer = frappe.get_doc("Customer", quotation.party_name)

    file_path = frappe.get_site_path("public", "files", "mẫu báo giá.xlsx")
    wb = load_workbook(file_path)
    ws = wb.active

    ws["B9"] = customer.customer_name or ""
    ws["I9"] = customer.phone or ""

    address_links = frappe.get_all("Dynamic Link", filters={
        "link_doctype": "Customer",
        "link_name": customer.name,
        "parenttype": "Address"
    }, fields=["parent"])

    if address_links:
        address_doc = frappe.get_doc("Address", address_links[0].parent)
        full_address = ", ".join(filter(None, [
            address_doc.address_line1,
            address_doc.address_line2,
            address_doc.city,
            address_doc.state,
            address_doc.country
        ]))
        ws["B10"] = full_address
    else:
        ws["B10"] = ""

    start_row = 14
    for i, item in enumerate(quotation.items):
        row = start_row + i
        ws[f"A{row}"] = i + 1
        ws[f"B{row}"] = item.item_name
        ws[f"E{row}"] = item.description or ""
        ws[f"G{row}"] = item.item_code
        ws[f"H{row}"] = item.qty

    ws["C17"] = quotation.total or 0
    ws["C18"] = 0
    ws["C19"] = 0
    ws["C20"] = quotation.total or 0

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    frappe.local.response.filename = f"Bao_gia_{quotation.name}.xlsx"
    frappe.local.response.filecontent = output.read()
    frappe.local.response.type = "binary"
