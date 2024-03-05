import zipfile
import os
import base64
from lxml import etree
from io import BytesIO

def find_power_query_files(excel_path):
    # List to store the paths of Power Query files found within the Excel file
    power_query_files = []
    
    # Check if the provided file is an Excel file based on its extension
    if not excel_path.endswith(('.xlsx', '.xlsm', '.xlsb')):
        raise ValueError("Unsupported file format. Please use .xlsx, .xlsm, or .xlsb files.")
    
    # Open the Excel file as a zip archive
    with zipfile.ZipFile(excel_path, 'r') as zip_ref:
        # Iterate through each file in the zip archive
        for file_info in zip_ref.infolist():
            # Check if the file is located within the customXml directory and has an xml extension
            if file_info.filename.startswith('customXml/') and file_info.filename.endswith('.xml'):
                # Open and read the content of the XML file
                with zip_ref.open(file_info) as xml_file:
                    xml_content_bytes = xml_file.read()
                    try:
                        # Parse the XML content
                        root = etree.fromstring(xml_content_bytes)
                        # Define the namespace used in the DataMashup elements
                        namespace = {'d': 'http://schemas.microsoft.com/DataMashup'}
                        # Search for DataMashup elements within the XML document
                        data_mashup_elements = root.xpath('//d:DataMashup', namespaces=namespace)
                        if data_mashup_elements:
                            # If found, decode the base64 content of the DataMashup element
                            base64_content = data_mashup_elements[0].text
                            decoded_content = base64.b64decode(base64_content)
                            # Look for the ZIP archive signatures to find the embedded ZIP archive
                            zip_start = decoded_content.find(b'PK\x03\x04') # Start of ZIP archive
                            zip_end = decoded_content.find(b'PK\x05\x06') # End of ZIP archive (end of central directory record)
                            if zip_start == -1 or zip_end == -1:
                                print("ZIP archive start or end signature not found.")
                            else:
                                # Extract the ZIP archive from the decoded content
                                zip_data = BytesIO(decoded_content[zip_start:zip_end + 22]) # Include the EOCD size
                                with zipfile.ZipFile(zip_data) as archive:
                                    print("ZIP Archive Contents:", archive.namelist())
                                    # Check for the presence of the 'Formulas/Section1.m' file, which contains Power Query formulas
                                    if 'Formulas/Section1.m' in archive.namelist():
                                        # Read and print the content of 'Formulas/Section1.m'
                                        section1_m_content = archive.read('Formulas/Section1.m').decode('utf-8')
                                        print("Content of Formulas/Section1.m:", section1_m_content)
                        else:
                            print("DataMashup content not found.")
                    except etree.XMLSyntaxError as e:
                        # Handle any XML parsing errors
                        print(f"XML parsing error: {e}")

    # Return the list of Power Query files found
    return power_query_files
