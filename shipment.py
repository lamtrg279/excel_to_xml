import xml.etree.ElementTree as ET
import pandas as pd
import re
from helper import safe_text, REQUIRED_COLUMNS

# JobShipment Section
def shipment_to_xml(parent, row, shipment_columns, content_columns):
    element = ET.SubElement(parent, "JobShipment")

    # ET.SubElement(element, "job").text = safe_text(getattr(row, "job"))

    for col in REQUIRED_COLUMNS:
        ET.SubElement(element, col).text = safe_text(getattr(row, col))

    for col in shipment_columns:
        if col in REQUIRED_COLUMNS:
            continue  # Already handled
        ET.SubElement(element, col).text = safe_text(getattr(row, col, None))

    # Handle forcedAccountNumber
    account_number = getattr(row, "accountNumber", None)
    if not pd.isna(account_number) and str(account_number).strip() != "":
        ET.SubElement(element, "forcedAccountNumber").text = "true"

    # shipment_contact = ET.SubElement(element, "contactNumber", id=contact_id)
    cartons = ET.SubElement(element, "Cartons")
    carton = ET.SubElement(cartons, "Carton")
    ET.SubElement(carton, "count").text = "1"
    ET.SubElement(carton, "quantity").text="1"
    ET.SubElement(carton, "addDefaultContent").text = "false"
    carton_contents = ET.SubElement(carton, "CartonContents")

    for col in content_columns:
        qty = getattr(row, col) 
        # ** Logic for future requests ship by JobPart or JobProduct or JobMaterial only
        # if not pd.isna(qty) and qty != 0:
            # ET.SubElement(carton_contents, "CartonContent") 
            # ET.SubElement(carton_contents, "jobPart").text = safe_text(getattr(row, "jobPart"))   
            # ET.SubElement(carton_contents, "jobPartJob").text = safe_text(getattr(row, "job"))
            # ET.SubElement(carton_contents, "quantity").text = safe_text(qty)
        #or
            # ET.SubElement(carton_contents, "CartonContent") 
            # ET.SubElement(carton_contents, "jobProduct").text = safe_text(getattr(row, "jobProduct"))   
            # ET.SubElement(carton_contents, "quantity").text = safe_text(qty)

        # Extract word and number from the 22nd column header
        match = re.search(r"\s*([A-Za-z]+)_(\d+)\s*", col)
        if match:
            word = match.group(1).lower()
            number = match.group(2)
        else:            continue # Skip if format is unexpected
        
        word = word[:3] + word[3].upper() + word[4:] if len(word) > 3 else word # Capitalize 4th letter if exists

        if not pd.isna(qty) and qty != 0:
            carton_content = ET.SubElement(carton_contents, "CartonContent")
            content = ET.SubElement(carton_content, word)
            content.text = safe_text(number)
            # Special handling for jobPart
            if word == "jobPart":
                jobpart_job = ET.SubElement(carton_content, "jobPartJob")
                jobpart_job.text = safe_text(getattr(row, "job"))   

            quantity = ET.SubElement(carton_content, "quantity")
            quantity.text = safe_text(qty)
    return element
