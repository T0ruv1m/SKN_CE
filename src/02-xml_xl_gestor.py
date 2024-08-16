import os
import xml.etree.ElementTree as ET
import csv
import pandas as pd
from config_tools import DIR

class XMLProcessor:
    """Process XML files to extract <nNF> and <dhEmi> elements."""

    def __init__(self, namespaces=None):
        # Define namespaces (if applicable)
        self.namespaces = namespaces or {'nfe': 'http://www.portalfiscal.inf.br/nfe'}

    def extract_data_from_xml(self, xml_file):
        """Extract the <nNF> and <dhEmi> elements' texts from the XML file."""
        try:
            tree = ET.parse(xml_file)
            root = tree.getroot()

            # Find the <nNF> element
            nnf_element = root.find('.//nfe:nNF', self.namespaces)
            nnf_text = nnf_element.text if nnf_element is not None else None

            # Find the <dhEmi> element and extract the year
            dhEmi_element = root.find('.//nfe:dhEmi', self.namespaces)
            dhEmi_text = dhEmi_element.text[:4] if dhEmi_element is not None and dhEmi_element.text is not None else None

            return nnf_text, dhEmi_text
        except ET.ParseError:
            print(f"Error parsing XML file: {xml_file}")
            return None, None

    def load_new_files_list(self, csv_file_path):
        """Load the list of new XML files to process from a CSV file."""
        new_files = []
        with open(csv_file_path, mode='r') as file:
            reader = csv.DictReader(file)
            for row in reader:
                if 'file_path' in row:
                    file_path = row['file_path'].strip()
                    new_files.append(file_path)
        return new_files

    def build_xml_file_mapping(self, new_files):
        """Build a dictionary mapping XML file names to their extracted values."""
        xml_data = {'chNTR': [], 'nNF': [], 'dhEmi': []}

        for xml_file_path in new_files:
            if not os.path.exists(xml_file_path):
                print(f"File not found: {xml_file_path}")
                continue

            print(f"Processing file: {xml_file_path}")
            file_name_without_ext = os.path.splitext(os.path.basename(xml_file_path))[0]

            # Extract data from the XML file
            nnf_text, dhEmi_text = self.extract_data_from_xml(xml_file_path)

            # Append the extracted data to the dictionary
            xml_data['chNTR'].append(file_name_without_ext)
            xml_data['nNF'].append(nnf_text)
            xml_data['dhEmi'].append(dhEmi_text)

        return xml_data

    def save_xml_data_to_excel(self, xml_data, excel_file_path):
        """Save the extracted XML data to an Excel file."""
        if not xml_data:
            print("No data to save.")
            return

        df = pd.DataFrame(xml_data)
        df.to_excel(excel_file_path, index=False)
        print(f"Data saved to: {excel_file_path}")

if __name__ == "__main__":
    # Define paths
    path_to = DIR()
    new_gestor= path_to.new_gestor 
    xl_gestor = path_to.xl_gestor

    # Instantiate the processor
    processor = XMLProcessor()

    # Load the list of new XML files
    new_files = processor.load_new_files_list(new_gestor)

    # Build the XML file mapping
    xml_data = processor.build_xml_file_mapping(new_files)

    # Save the extracted data to an Excel file
    processor.save_xml_data_to_excel(xml_data, xl_gestor)