import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import xml.etree.ElementTree as ET
import re


standard_columns = df.columns[:7]
jobPart_columns = df.columns[21:]
contact_columns = df.columns[7:21]


# JobShipment Section
def row_to_xml(parent, row, contact_id):
    element = ET.SubElement(parent, "JobShipment")

    job = ET.SubElement(element, "job")
    job.text = str(row["job"])

    for col in standard_columns:
        # Create XML tag
        child = ET.SubElement(element, col)
        # Get text value, handling NaN values
        value = row.get(col)
        # Set text, converting NaN to empty string
        child.text = "" if pd.isna(value) else str(value) 
    
    shipment_contact = ET.SubElement(element, "contactNumber", id=contact_id)

    cartons = ET.SubElement(element, "Cartons")
    carton = ET.SubElement(cartons, "Carton")
    count = ET.SubElement(carton, "count")
    count.text = "1"
    carton_qty = ET.SubElement(carton, "quantity")
    carton_qty.text = "1"
    add_default_content = ET.SubElement(carton, "addDefaultContent")
    add_default_content.text = "false"
    carton_contents = ET.SubElement(carton, "CartonContents")
    for col in jobPart_columns:
        carton_content = ET.SubElement(carton_contents, "CartonContent")
        value = col 
        qty = row[value]
        # if pd.isna(qty):
        #     continue
        jobpart = ET.SubElement(carton_content, "jobPart")
        jobpart.text = str(value)
        quantity = ET.SubElement(carton_content, "quantity")
        quantity.text = str(qty)
        jobpart_job = ET.SubElement(carton_content, "jobPartJob")
        jobpart_job.text = str(row["job"])
    return element

    # jobcontacts = ET.SubElement(element, "JobContacts")
    # jobcontact = ET.SubElement(jobcontacts, "JobContact", id="contact01")
    # jobcontact_job = ET.SubElement(jobcontact, "job")
    # jobcontact_job.text = str(row["job"])
    # billto = ET.SubElement(jobcontact, "billTo")
    # billto.text = str(row["billTo"])
    # shipto = ET.SubElement(jobcontact, "shipTo")
    # shipto.text = str(row["shipTo"])
    # jobcontact_contact = ET.SubElement(jobcontact, "contact", id="1")

# Contacts Section
contacts = ET.SubElement(outerRoot, "Contacts")
def contact_to_xml(parent, row, countact_id):
    contact = ET.SubElement(parent, "Contact", refId=countact_id)
    company_legal_name = ET.SubElement(contact, "companyName")
    company_legal_name.text = str(row.get("company"))  
    title = ET.SubElement(contact, "title")
    title.text = str(row["title"])
    first_name = ET.SubElement(contact, "firstName")
    first_name.text = "" if pd.isna(row.get("contactFirstName")) else str(row.get("contactFirstName"))
    last_name = ET.SubElement(contact, "lastName")
    last_name.text = "" if pd.isna(row.get("contactLastName")) else str(row.get("contactLastName"))
    addres1 = ET.SubElement(contact, "address1")
    addres1.text = str(row["address1"])
    address2 = ET.SubElement(contact, "address2")
    address2.text = "" if pd.isna(row.get("address2")) else str(row.get("address2"))
    address3 = ET.SubElement(contact, "address3")
    address3.text = "" if pd.isna(row.get("address3")) else str(row.get("address3"))
    city = ET.SubElement(contact, "city")
    city.text = str(row["city"])
    state = ET.SubElement(contact, "state")
    state.text = str(row["state"])
    zip=ET.SubElement(contact, "zip")
    zip.text = str(row["zip"])
    country = ET.SubElement(contact, "country")
    country.text = str(row["country"])
    phone = ET.SubElement(contact, "businessPhoneNumber")
    phone.text = "" if pd.isna(row.get("phone")) else str(row.get("phone"))
    email = ET.SubElement(contact, "email")
    email.text = "" if pd.isna(row.get("email")) else str(row.get("email"))
    return contact

# Contact ID generator
def contact_id_generator():
    counter = 1
    while True:
        yield f"contact{counter:02d}"
        counter += 1
contact_ids = contact_id_generator()


def run_program():
    file_path = file_entry.get()
    if not file_path:
        messagebox.showerror("Error", "Please select an Excel file.")
        return

    try:
        df = pd.read_excel(file_path)
        df.columns = df.columns.str.strip()
        # (Rest of the XML generation code goes here, using the selected file)
        messagebox.showinfo("Success", "XML file generated successfully.")

        outerRoot = ET.Element("import")
        root = ET.SubElement(outerRoot, "JobShipments")
        # Iterate over DataFrame rows and build XML
        for _, row in df.iterrows():
            contact_id = next(contact_ids)
            row_to_xml(root, row, contact_id)
            contact_to_xml(contacts, row, contact_id)

        # Write XML to file
        tree = ET.ElementTree(outerRoot)
        output_file = re.sub(r"[\/\\:*?\"<>|]", "_", str(df.iloc[0]["job"]))
        try:
            tree.write(f"Shipments for job {output_file}.xml", encoding="utf-8", xml_declaration=True)
            print("XML file generated successfully.")
        except Exception as e:
            print(f"Error writing XML file: {e}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read Excel file: {e}")