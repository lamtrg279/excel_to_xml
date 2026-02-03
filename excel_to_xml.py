import pandas as pd
import xml.etree.ElementTree as ET

# Step 1: Load Excel data
df = pd.read_excel('data.xlsx')
# Strip whitespace from column names
df.columns = df.columns.str.strip()
# Optional: Display DataFrame info for debugging
DEBUG = False
if DEBUG:
    print(df.shape)
    print(df.dtypes)
    print(df.head())
# Step 2: Convert DataFrame to XML
standard_columns = df.columns[:24]
jobPart_columns = df.columns[24:]

outerRoot = ET.Element("import")
root = ET.SubElement(outerRoot, "JobShipments")

# Function to convert a DataFrame row to XML
def row_to_xml(parent, row):
    element = ET.SubElement(parent, "JobShipment")
    for col in standard_columns:
        # Create XML tag
        child = ET.SubElement(element, col)
        # Get text value, handling NaN values
        value = row.get(col)
        # Set text, converting NaN to empty string
        child.text = "" if pd.isna(value) else str(value) 

    cartons = ET.SubElement(element, "Cartons")
    carton = ET.SubElement(cartons, "Carton")
    count = ET.SubElement(carton, "count")
    count.text = "1"
    carton_contents = ET.SubElement(carton, "CartonContents")

    for col in jobPart_columns:
        carton_content = ET.SubElement(carton_contents, "CartonContent")
        value = col 
        qty = row[value]

        print(value, qty)
        # if pd.isna(qty):
        #     continue
        jobpart = ET.SubElement(carton_content, "jobPart")
        jobpart.text = str(value)

        quantity = ET.SubElement(carton_content, "quantity")
        quantity.text = str(qty)

        jobpart_job = ET.SubElement(carton_content, "jobPartJob")
        jobpart_job.text = str(row["job"])

    jobcontacts = ET.SubElement(element, "JobContacts")
    jobcontact = ET.SubElement(jobcontacts, "JobContact")
    jobcontact_job = ET.SubElement(jobcontact, "job")
    jobcontact_job.text = str(row["job"])
    billto = ET.SubElement(jobcontact, "billTo")
    billto.text = str(row["billTo"])
    shipto = ET.SubElement(jobcontact, "shipTo")
    shipto.text = str(row["shipTo"])
    
    return element

# Iterate over DataFrame rows and build XML
for _, row in df.iterrows():
    row_to_xml(root, row)
# Write XML to file
tree = ET.ElementTree(outerRoot)
tree.write("output.xml", encoding="utf-8", xml_declaration=True)