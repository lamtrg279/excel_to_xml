import xml.etree.ElementTree as ET
from helper import safe_text


# Contacts Section
def contact_to_xml(parent, row, contact_id, contact_columns):
    contact = ET.SubElement(parent, "Contact", refId=contact_id)
    ET.SubElement(contact, "jobContact").text = "true"

    for col in contact_columns:
        ET.SubElement(contact, col).text = safe_text(getattr(row, col, None))
        # if str(col) != "phone":
        #     child.text = "" if pd.isna(value) else str(value) 
        # else:
        #     child.text = safe_text(value) 
    return contact

# Contact ID generator
def contact_id_generator():
    counter = 1
    while True:
        yield f"contact{counter:02d}"
        counter += 1
