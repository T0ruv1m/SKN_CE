import os
import xml.etree.ElementTree as ET
import csv
from config_tools import DIR
import re
import pandas as pd

def extract_data_from_xml(xml_file):
    """Extract texts from XML elements."""
    try:
        tree = ET.parse(xml_file)
        root = tree.getroot()

        # Define namespaces
        namespaces = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}

        # Extract data
        chNTR_element = root.find('.//nfe:protNFe/nfe:infProt/nfe:chNFe', namespaces)
        chNTR_text = chNTR_element.text if chNTR_element is not None else None

        xMun_element = root.find('.//nfe:dest/nfe:enderDest/nfe:xMun', namespaces)
        xMun_text = xMun_element.text if xMun_element is not None else None

        infCpl_element = root.find('.//nfe:infAdic/nfe:infCpl', namespaces)
        infCpl_text = infCpl_element.text if infCpl_element is not None else ''

        number_match = re.search(r'\b\d{44}\b', infCpl_text)
        extracted_number = number_match.group(0) if number_match else None
        '''
        last_semicolon_text = infCpl_text.split(';')[-1].strip() if ';' in infCpl_text else None

        if last_semicolon_text and len(last_semicolon_text) >= 50:
            regex_pattern = r'\d{44}\s+([A-Z\s.,\'()/-]+)'
            regex_match = re.search(regex_pattern, infCpl_text)
            last_semicolon_text = regex_match.group(1).strip() if regex_match else last_semicolon_text

        if last_semicolon_text and not re.match(r'^[A-Z]{2}', last_semicolon_text):
            pattern = r'\d{44}\s+([A-Z\s.,\'()/-]+)'
            infCpl_match = re.search(pattern, infCpl_text)
            last_semicolon_text = infCpl_match.group(1).strip() if infCpl_match else None
        '''
        vProd_element = root.find('.//nfe:total/nfe:ICMSTot/nfe:vProd', namespaces)
        vProd_text = vProd_element.text if vProd_element is not None else None

        return extracted_number, chNTR_text, xMun_text, vProd_text
    except ET.ParseError:
        print(f"Error parsing XML file: {xml_file}")
        return None, None, None, None, None

def load_new_files_list(csv_file_path):
    """Load the list of new XML files to process from a CSV file."""
    new_files = []
    with open(csv_file_path, mode='r') as file:
        reader = csv.DictReader(file)
        for row in reader:
            if 'file_path' in row:
                file_path = row['file_path'].strip()
                new_files.append(file_path)
    return new_files

def load_existing_data(excel_file_path):
    """Load existing data from the Excel file if it exists."""
    if os.path.exists(excel_file_path):
        return pd.read_excel(excel_file_path)
    else:
        return pd.DataFrame(columns=['file_name', 'chNF', 'chNTR', 'xMun', 'Fornecedor', 'vProd'])

def build_xml_file_mapping(new_files, existing_data):
    """Build a dictionary mapping XML file names to their extracted values, avoiding duplicates."""
    xml_data = {'file_name': [], 'chNF': [], 'chNTR': [], 'xMun': [], 'Fornecedor': [], 'vProd': []}

    for xml_file_path in new_files:
        if not os.path.exists(xml_file_path):
            print(f"File not found: {xml_file_path}")
            continue

        print(f"Processing file: {xml_file_path}")
        file_name_without_ext = os.path.splitext(os.path.basename(xml_file_path))[0]

        # Skip if this file's data is already in existing_data
        if file_name_without_ext in existing_data['file_name'].values:
            print(f"File already processed: {file_name_without_ext}")
            continue

        extracted_number, chNTR_text, xMun_text, last_semicolon_text, vProd_text = extract_data_from_xml(xml_file_path)

        xml_data['file_name'].append(file_name_without_ext)
        xml_data['chNF'].append(extracted_number)
        xml_data['chNTR'].append(chNTR_text)
        xml_data['xMun'].append(xMun_text)
        # xml_data['Fornecedor'].append(last_semicolon_text)
        xml_data['vProd'].append(vProd_text)

    return xml_data

def save_xml_data_to_excel(xml_data, excel_file_path):
    """Save the XML data to an Excel file, combining with existing data and avoiding duplicates."""
    if not xml_data['file_name']:
        print("No new data to save.")
        return

    # Load existing data
    existing_data = load_existing_data(excel_file_path)

    # Convert new data to DataFrame
    new_data = pd.DataFrame(xml_data)

    # Combine the new data with the existing data, avoiding duplicates
    combined_data = pd.concat([existing_data, new_data], ignore_index=True).drop_duplicates(subset=['file_name'])

    # Save the combined DataFrame to the Excel file
    combined_data.to_excel(excel_file_path, index=False)
    print(f"Data saved to: {excel_file_path}")

if __name__ == "__main__":
    # Define paths
    path_to = DIR()
    new_compras = path_to.new_compras
    xl_compras = path_to.xl_compras

    # Load the list of new XML files
    new_files = load_new_files_list(new_compras)

    # Load existing data
    existing_data = load_existing_data(xl_compras)

    # Build the XML file mapping
    xml_data = build_xml_file_mapping(new_files, existing_data)

    # Save the extracted data to an Excel file
    save_xml_data_to_excel(xml_data, xl_compras)
