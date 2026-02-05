import tkinter as tk
import os
from tkinter import filedialog, messagebox
import pandas as pd
import xml.etree.ElementTree as ET
import re

# JobShipment Section
def shipment_to_xml(parent, row, contact_id, shipment_columns, content_columns):
    element = ET.SubElement(parent, "JobShipment")

    job = ET.SubElement(element, "job")
    job.text = str(row["job"])

    for col in shipment_columns:
        child = ET.SubElement(element, col)
        value = row.get(col)
        child.text = "" if pd.isna(value) else str(value) 

    if not pd.isna(row.get("accountNumber")) or str(row.get("accountNumber")).strip() != "":
        forced_account_number = ET.SubElement(element, "forcedAccountNumber")
        forced_account_number.text = "true"

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

    for col in content_columns:
        value = col 
        qty = row[value]
        # Extract word and number from the 22nd column header
        match = re.search(r"\s*([A-Za-z]+)\s*(\d+)\s*", value)
        if match:
            word = match.group(1).lower()
            number = match.group(2)
        
        word = word[:3] + word[3].upper() + word[4:] if len(word) > 3 else word # Capitalize 4th letter if exists

        if not pd.isna(qty) and qty != 0:
            carton_content = ET.SubElement(carton_contents, "CartonContent")
            content = ET.SubElement(carton_content, word)
            content.text = str(number)
            if word == "jobPart":
                jobpart_job = ET.SubElement(carton_content, "jobPartJob")
                jobpart_job.text = str(row["job"])

            quantity = ET.SubElement(carton_content, "quantity")
            quantity.text = "" if pd.isna(qty) else str(qty)
    return element

# Contacts Section
def contact_to_xml(parent, row, contact_id, contact_columns):
    contact = ET.SubElement(parent, "Contact", refId=contact_id)
    # Global contact false
    jobContact = ET.SubElement(contact, "jobContact")
    jobContact.text = "true"    

    for col in contact_columns:
        child = ET.SubElement(contact, col)
        value = row.get(col)
        if str(col) != "phone":
            child.text = "" if pd.isna(value) else str(value) 
        else:
            clean_phone = "" if pd.isna(value) else str(value).strip()
            child.text = clean_phone
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

        shipment_columns = df.columns[:9]
        content_columns = df.columns[23:]
        contact_columns = df.columns[9:22]


        outerRoot = ET.Element("import")
        root = ET.SubElement(outerRoot, "JobShipments")
        contacts = ET.SubElement(outerRoot, "Contacts")

        # Iterate over DataFrame rows and build XML
        for _, row in df.iterrows():
            contact_id = next(contact_ids)
            shipment_to_xml(root, row, contact_id, shipment_columns, content_columns)
            contact_to_xml(contacts, row, contact_id, contact_columns)

        # Write XML to file
        tree = ET.ElementTree(outerRoot)
        job_number = re.sub(r"[\/\\:*?\"<>|]", "_", str(df.iloc[0]["job"]))
        output_file = os.path.join(os.path.dirname(file_path), f"{job_number}.xml") # Save to same directory as input file
        tree.write(output_file, encoding="utf-8", xml_declaration=True) 
        # Success message
        result_label.config(text=f"✅ XML file saved to: {output_file}", fg="green")
    except Exception as e:
        result_label.config(text=f"❌ Error: {str(e)}", fg="red")

# GUI Setup
root = tk.Tk()
root.title("Excel to XML Converter for JobShipment import")
root.geometry("450x200")

file_label = tk.Label(root, text="Select Excel File:")
file_label.pack(pady=5)

#File path
file_entry = tk.Entry(root, width=60)
file_entry.pack(pady=5)

#Browse file function
def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    file_entry.delete(0, tk.END) #Clear existing text
    file_entry.insert(0, file_path) #Insert selected file path
    result_label.config(text="") #Clear previous result message

#Browse button
browse_button = tk.Button(root, text="Browse", command=browse_file)
browse_button.pack(pady=5)

#Run button
run_button = tk.Button(root, text="Convert to XML", command=run_program)
run_button.pack(pady=10)

#Result display
result_label = tk.Label(root, text="", wraplength=450)
result_label.pack(pady=5)

root.mainloop()