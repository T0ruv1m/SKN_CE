import os
import xml.etree.ElementTree as ET
import csv
from config_tools import DIR
import re
import pandas as pd

def extract_data_from_xml(xml_file):
    """Extrai os textos dos elementos chNFe, xMun, infCpl, chNF, Fornecedor, e vProd do arquivo XML."""
    try:
        tree = ET.parse(xml_file)
        root = tree.getroot()

        # Define namespaces to handle XML namespaces correctly
        namespaces = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}

        # Find the chNFe element and extract its text
        chNFe_element = root.find('.//nfe:protNFe/nfe:infProt/nfe:chNFe', namespaces)
        chNFe_text = chNFe_element.text if chNFe_element is not None else None

        # Find the xMun element under enderDest and extract its text
        xMun_element = root.find('.//nfe:dest/nfe:enderDest/nfe:xMun', namespaces)
        xMun_text = xMun_element.text if xMun_element is not None else None

        # Find the infCpl element and extract its text
        infCpl_element = root.find('.//nfe:infAdic/nfe:infCpl', namespaces)
        infCpl_text = infCpl_element.text if infCpl_element is not None else ''

        # Extract the 44-digit number using regex
        number_match = re.search(r'\b\d{44}\b', infCpl_text)
        extracted_number = number_match.group(0) if number_match else None

        # Extract the text after the last semicolon in infCpl
        last_semicolon_text = infCpl_text.split(';')[-1].strip() if ';' in infCpl_text else None

        if last_semicolon_text and len(last_semicolon_text) >= 50:
                regex_pattern = r'\d{44}\s+([A-Z\s.,\'()/-]+)'
                regex_match = re.search(regex_pattern, infCpl_text)
                last_semicolon_text = regex_match.group(1).strip() if regex_match else last_semicolon_text


        # Check if the last_semicolon_text starts with two uppercase letters
        if last_semicolon_text and not re.match(r'^[A-Z]{2}', last_semicolon_text):
            # Apply the regex pattern to extract uppercase text following the 44-digit number
            pattern = r'\d{44}\s+([A-Z\s.,\'()/-]+)'
            infCpl_match = re.search(pattern, infCpl_text)
            last_semicolon_text = infCpl_match.group(1).strip() if infCpl_match else None
    
        # Find the vProd element within ICMSTot and extract its text
        vProd_element = root.find('.//nfe:total/nfe:ICMSTot/nfe:vProd', namespaces)
        vProd_text = vProd_element.text if vProd_element is not None else None

        return extracted_number, chNFe_text, xMun_text, last_semicolon_text, vProd_text
    except ET.ParseError:
        print(f"Erro ao analisar o arquivo XML: {xml_file}")
        return None, None, None, None, None

def load_new_files_list(csv_file_path):
    """Carrega a lista de novos arquivos XML a serem processados do arquivo CSV."""
    new_files = []
    
    with open(csv_file_path, mode='r') as file:
        reader = csv.DictReader(file)
        for row in reader:
            if 'file_path' in row:
                file_path = row['file_path'].strip()
                new_files.append(file_path)
    
    return new_files

def build_xml_file_mapping(new_files):
    """Constrói um dicionário mapeando nomes de arquivos XML para seus valores extraídos."""
    xml_data = {'file_name': [], 'chNF': [], 'chNFe': [], 'xMun': [], 'Fornecedor': [], 'vProd': []}

    for xml_file_path in new_files:
        # Verifica se o arquivo existe
        if not os.path.exists(xml_file_path):
            print(f"Arquivo não encontrado: {xml_file_path}")
            continue  # Skip if the file doesn't exist

        print(f"Processando arquivo: {xml_file_path}")
        file_name_without_ext = os.path.splitext(os.path.basename(xml_file_path))[0]

        # Extrai os dados do arquivo XML
        extracted_number, chNFe_text, xMun_text, last_semicolon_text, vProd_text = extract_data_from_xml(xml_file_path)

        '''
         # Debugging information to check extracted values
        print(f"Extraído do arquivo {file_name_without_ext}:")
        print(f"chNF: {extracted_number}")
        print(f"chNFe: {chNFe_text}")
        print(f"xMun: {xMun_text}")
        print(f"Fornecedor: {last_semicolon_text}")
        print(f"vProd: {vProd_text}")'''

        # Append the values to the corresponding lists in the dictionary
        xml_data['file_name'].append(file_name_without_ext)
        xml_data['chNF'].append(extracted_number)
        xml_data['chNFe'].append(chNFe_text)
        xml_data['xMun'].append(xMun_text)
        xml_data['Fornecedor'].append(last_semicolon_text)
        xml_data['vProd'].append(vProd_text)

    return xml_data

def save_xml_data_to_excel(xml_data, excel_file_path):
    """Salva o mapeamento de dados XML em um arquivo Excel."""
    if not xml_data:
        print("Nenhum dado extraído para salvar.")
        return

    # Convert the list of dictionaries to a pandas DataFrame
    df = pd.DataFrame(xml_data)
    
    # Save the DataFrame to an Excel file
    df.to_excel(excel_file_path, index=False)

    print(f"Dados XML salvos com sucesso em: {excel_file_path}")

if __name__ == "__main__":
    # Define os caminhos
    path_to = DIR()
    new_compras = path_to.new_compras
    xl_compras = path_to.xl_compras

    # Carrega a lista de novos arquivos XML
    new_files = load_new_files_list(new_compras)

    # Constrói o mapeamento de arquivos XML
    xml_data = build_xml_file_mapping(new_files)

    # Certifique-se de que xml_data é um dicionário antes de salvá-lo
    if isinstance(xml_data, dict):
        # Salva os dados XML extraídos em um arquivo CSV
        save_xml_data_to_excel(xml_data, xl_compras)
    else:
        print("Erro: xml_data não é um dicionário válido")
