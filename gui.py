import tkinter as tk
from tkinter import filedialog, messagebox
import os
import pandas as pd
import xml.etree.ElementTree as ET
import re
import traceback

from helper import safe_text, validate_columns_rows, REQUIRED_COLUMNS
from shipment import shipment_to_xml
from contact import contact_to_xml, contact_id_generator


# Main program logic
def run_program():
    global contact_ids
    contact_ids = contact_id_generator() # Reset contact ID generator for each run
    file_path = file_entry.get()
    if not file_path:
        messagebox.showerror("Error", "Please select an Excel file.")
        return

    try:
        df = pd.read_excel(file_path)
        df.columns = df.columns.str.strip()
        validate_columns_rows(df, REQUIRED_COLUMNS)
        df.columns = df.columns.str.replace(r"\s+", "_", regex=True) # Replace spaces with underscores in column names
        shipment_columns = df.columns[:9]
        contact_columns = df.columns[9:22]
        content_columns = df.columns[23:]

        outerRoot = ET.Element("import")
        root = ET.SubElement(outerRoot, "JobShipments")
        contacts = ET.SubElement(outerRoot, "Contacts")

        # Iterate over DataFrame rows and build XML
        for row in df.itertuples(index=False):
            contact_id = next(contact_ids)
            shipment_to_xml(root, row, contact_id, shipment_columns, content_columns)
            contact_to_xml(contacts, row, contact_id, contact_columns)

        # Write XML to file
        tree = ET.ElementTree(outerRoot)
        job_number = re.sub(r"[\/\:*?\"<>|]", "_", safe_text(df.iloc[0]["job"]))
        output_file = os.path.join(os.path.dirname(file_path), f"{job_number}.xml") # Save to same directory as input file
        tree.write(output_file, encoding="utf-8", xml_declaration=True) 
        # Success message
        result_label.config(text=f"✅ XML file saved to: {output_file}", fg="green")
    except Exception as e:
        tb = traceback.format_exc()
        result_label.config(text=f"❌ Error: {str(e)}", fg="red")
        print(f"Error: {str(e)}\n{tb}")

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
