import os
import xml.etree.ElementTree as ET
import csv
import pandas as pd
import re
from config_tools import DIR

class XMLProcessor:
    """Process XML files to extract specific elements and save them to an Excel file."""

    def __init__(self, namespaces=None):
        # Define namespaces (if applicable)
        self.namespaces = namespaces or {'nfe': 'http://www.portalfiscal.inf.br/nfe'}

    def extract_data_from_xml(self, xml_file, extraction_type='gestor'):
        """Extract data from XML file based on the specified extraction type."""
        try:
            tree = ET.parse(xml_file)
            root = tree.getroot()

            if extraction_type == 'gestor':
                return self._extract_gestor_data(root)
            elif extraction_type == 'compras':
                return self._extract_compras_data(root)
        except ET.ParseError:
            print(f"Error parsing XML file: {xml_file}")
            return None

    def _extract_gestor_data(self, root):
        """Extract <nNF>, <dhEmi>, and <xFant> elements for Gestor."""
        nnf_element = root.find('.//nfe:nNF', self.namespaces)
        nnf_text = nnf_element.text if nnf_element is not None else None

        dhEmi_element = root.find('.//nfe:dhEmi', self.namespaces)
        dhEmi_text = dhEmi_element.text[:4] if dhEmi_element is not None and dhEmi_element.text is not None else None

        xFant_element = root.find('.//nfe:xFant', self.namespaces) or root.find('.//nfe:xNome', self.namespaces)
        xFant_text = xFant_element.text if xFant_element is not None else None

        return nnf_text, dhEmi_text, xFant_text

    def _extract_compras_data(self, root):
        """Extract specific elements for Compras."""
        chNTR_element = root.find('.//nfe:protNFe/nfe:infProt/nfe:chNFe', self.namespaces)
        chNTR_text = chNTR_element.text if chNTR_element is not None else None

        xMun_element = root.find('.//nfe:dest/nfe:enderDest/nfe:xMun', self.namespaces)
        xMun_text = xMun_element.text if xMun_element is not None else None

        infCpl_element = root.find('.//nfe:infAdic/nfe:infCpl', self.namespaces)
        infCpl_text = infCpl_element.text if infCpl_element is not None else None

        number_match = re.search(r'\b\d{44}\b', infCpl_text)
        extracted_number = number_match.group(0) if number_match else None

        vProd_element = root.find('.//nfe:total/nfe:ICMSTot/nfe:vProd', self.namespaces)
        vProd_text = vProd_element.text if vProd_element is not None else None

        return extracted_number, chNTR_text, xMun_text, vProd_text

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

    def load_existing_data(self, excel_file_path, columns):
        """Load existing data from the Excel file if it exists."""
        if os.path.exists(excel_file_path):
            return pd.read_excel(excel_file_path)
        else:
            return pd.DataFrame(columns=columns)

    def build_xml_file_mapping(self, new_files, existing_data, extraction_type='gestor'):
        """Build a dictionary mapping XML file names to their extracted values, avoiding duplicates."""
        if extraction_type == 'gestor':
            columns = ['chNTR', 'nNF', 'dhEmi', 'xFant']
        else:  # compras
            columns = ['file_name', 'chNF', 'chNTR', 'xMun', 'vProd']

        xml_data = {col: [] for col in columns}

        for xml_file_path in new_files:
            if not os.path.exists(xml_file_path):
                print(f"File not found: {xml_file_path}")
                continue

            print(f"Processing file: {xml_file_path}")
            file_name_without_ext = os.path.splitext(os.path.basename(xml_file_path))[0]

            if extraction_type == 'gestor' and file_name_without_ext in existing_data['chNTR'].values:
                print(f"File already processed: {file_name_without_ext}")
                continue
            elif extraction_type == 'compras' and file_name_without_ext in existing_data['file_name'].values:
                print(f"File already processed: {file_name_without_ext}")
                continue

            extracted_data = self.extract_data_from_xml(xml_file_path, extraction_type=extraction_type)

            if extraction_type == 'gestor':
                nnf_text, dhEmi_text, xFant_text = extracted_data
                xml_data['chNTR'].append(file_name_without_ext)
                xml_data['nNF'].append(nnf_text)
                xml_data['dhEmi'].append(dhEmi_text)
                xml_data['xFant'].append(xFant_text)
            else:  # compras
                extracted_number, chNTR_text, xMun_text, vProd_text = extracted_data
                xml_data['file_name'].append(file_name_without_ext)
                xml_data['chNF'].append(extracted_number)
                xml_data['chNTR'].append(chNTR_text)
                xml_data['xMun'].append(xMun_text)
                xml_data['vProd'].append(vProd_text)

        return xml_data

    def save_xml_data_to_excel(self, xml_data, excel_file_path, columns):
        """Save the XML data to an Excel file, combining with existing data and avoiding duplicates."""
        if not xml_data[next(iter(xml_data))]:  # Check if there's data in the first column
            print("No new data to save.")
            return

        existing_data = self.load_existing_data(excel_file_path, columns)

        new_data = pd.DataFrame(xml_data)

        combined_data = pd.concat([existing_data, new_data], ignore_index=True).drop_duplicates(subset=[columns[0]])

        combined_data.to_excel(excel_file_path, index=False)
        print(f"Data saved to: {excel_file_path}")

class ExcelMerger:
    def __init__(self, file1_path, file2_path, merge_column, output_file):
        """
        Inicializa o ExcelMerger com os caminhos dos arquivos e a coluna de junção.
        
        :param file1_path: Caminho do primeiro arquivo Excel.
        :param file2_path: Caminho do segundo arquivo Excel.
        :param merge_column: Nome da coluna que será utilizada para a junção.
        :param output_file: Caminho do arquivo Excel de saída.
        """
        self.file1_path = file1_path
        self.file2_path = file2_path
        self.merge_column = merge_column
        self.output_file = output_file


    def merge_excel_files(self):
        """
        Realiza a junção dos dois arquivos Excel e salva o resultado em um novo arquivo.
        """
        # Carrega os dados dos arquivos Excel
        df1 = pd.read_excel(self.file1_path)
        df2 = pd.read_excel(self.file2_path)
        print(df1)
        print(df2)

        # Realiza a junção com base na coluna especificada
        merged_df = pd.merge(df1, df2, on=self.merge_column, how='outer')
        merged_df = pd.merge(merged_df, df2, left_on='chNF', right_on='chNTR')  
        
        merged_df = merged_df[['xMun','chNTR_x','nNF_x','chNF','nNF_y','vProd','xFant_y','dhEmi_x']].rename(columns={
            'xMun':'Municipio',
            'xFant_y' : 'FOR',
            'chNTR_x':'chNTR',
            'nNF_x': 'nNTR',
            'nNF_y': 'nNF',
            'vProd': 'Valor',
            'dhEmi_x': 'Ano'
        })
        
        #df2['nNF'] = df2['nNF'].astype(str)
        #merged2 = pd.merge(df1, df2[['chNTR','nNF']], left_on='chNF',right_on='nNF')
        
        # Salva o resultado em um novo arquivo Excel
        merged_df.to_excel(self.output_file, index=False)
        print(f"Arquivos combinados e salvos em {self.output_file}")

if __name__ == "__main__":
    # Define paths
    path_to = DIR()

    # Gestor processing
    processor = XMLProcessor()
    new_gestor = path_to.new_gestor 
    xl_gestor = path_to.xl_gestor

    new_files_gestor = processor.load_new_files_list(new_gestor)
    existing_data_gestor = processor.load_existing_data(xl_gestor, ['chNTR', 'nNF', 'dhEmi', 'xFant'])
    xml_data_gestor = processor.build_xml_file_mapping(new_files_gestor, existing_data_gestor, extraction_type='gestor')
    processor.save_xml_data_to_excel(xml_data_gestor, xl_gestor, ['chNTR', 'nNF', 'dhEmi', 'xFant'])

    # Compras processing
    new_compras = path_to.new_compras
    xl_compras = path_to.xl_compras

    new_files_compras = processor.load_new_files_list(new_compras)
    existing_data_compras = processor.load_existing_data(xl_compras, ['file_name', 'chNF', 'chNTR', 'xMun', 'vProd'])
    xml_data_compras = processor.build_xml_file_mapping(new_files_compras, existing_data_compras, extraction_type='compras')
    processor.save_xml_data_to_excel(xml_data_compras, xl_compras, ['file_name', 'chNF', 'chNTR', 'xMun', 'vProd'])

    # Combiner processing

    file1 = path_to.xl_compras
    file2 = path_to.xl_gestor
    column_to_merge_on = 'chNTR'  # Nome da coluna usada para a junção
    output_file = path_to.xl_combi
    

    merger = ExcelMerger(file1, file2, column_to_merge_on, output_file)
    merger.merge_excel_files()
    