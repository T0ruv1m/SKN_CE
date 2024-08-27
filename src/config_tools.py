# config.py
import os
import pandas as pd

class DIR:
    """Classe de configuração para diretórios e variáveis compartilhadas."""

    def __init__(self):
        # Define o diretório raiz
        self.root = os.path.dirname(os.path.abspath(__file__))

        # Diretórios principais
        self.dirs = {
            'dig': '../docs/digitalizados',
            'res': '../docs/registro',
            'reg': '../docs/residuo',
            'mesc': '../docs/mesclados',
            
            'xml_data': '\\\\Mimitop\\02334933000140',  # UNC path with double backslashes
            'gestor_data': 'I:/.shortcut-targets-by-id/1ghlKQQOndN3wMxNW4qM3pTPDbtrYJmTa/GestorDFe/Documentos/CONSORCIO INTERMUNICIPAL DE SAUDE ALTO DAS VERTENTES',
            
            'timestamp': '../root/timestamp.txt',
            'credentials_file': '../root/credentials.json',
            'token_file': '../root/token.json',
            'json_object': '../root/object.json',
            'merged_files_json':'../root/merged_files.json',
            
            'cache_compras':'../root/cache_compras.csv',
            'cache_gestor':'../root/cache_gestor.csv',
            
            'new_compras':'../root/new_compras.csv',
            'new_gestor':'../root/new_gestor.csv',

            'xl_compras':'../data/xlsx_data/xl_compras.xlsx',
            'xl_gestor':'../data/xlsx_data/xl_gestor.xlsx',
            'xl_combi':'../data/xlsx_data/xl_combinada.xlsx'  # Removed redundant key


        }
        
        # Atualiza caminhos baseados no diretório raiz
        self.update_paths()

    def update_paths(self):
        """Atualiza todos os caminhos relativos para caminhos absolutos baseados no diretório raiz."""
        for key, relative_path in self.dirs.items():
            if not os.path.isabs(relative_path):
                setattr(self, key, os.path.join(self.root, relative_path))
            else:
                setattr(self, key, relative_path)

class EmptyFileMaker:
    
    def __init__(self, file_path):
        self.file_path = file_path
        self.extension = os.path.splitext(file_path)[1].lower()
    
    def ensure_directory_exists(self):
        directory = os.path.dirname(self.file_path)
        if not os.path.exists(directory):
            os.makedirs(directory)
            print(f"Novo Diretório Criado: {directory}")
    
    def empty_file(self):
        self.ensure_directory_exists()
        if not os.path.exists(self.file_path):
            if self.extension == '.xlsx':
                self.create_empty_excel()
            elif self.extension == '.txt':
                self.create_empty_txt()
            elif self.extension == '.json':
                self.create_empty_json()
            else:
                print(f"Extensão de arquivo não suportada: {self.extension}")
        else:
            print(f"Este arquivo já existe: {self.file_path}")
    
    def create_empty_excel(self):
        try:
            df = pd.DataFrame()
            df.to_excel(self.file_path, index=False)
            print(f"Arquivo Excel criado: {self.file_path}")
        except Exception as e:
            print(f"Erro criando arquivo Excel: {e}")
    
    def create_empty_txt(self):
        try:
            with open(self.file_path, 'w') as file:
                pass  # Create an empty file
            print(f"Arquivo de texto criado: {self.file_path}")
        except Exception as e:
            print(f"Erro criando arquivo de texto: {e}")
    def create_empty_json(self):
        try:
            with open(self.file_path, 'w') as file:
                file.write('{}')
                print(f"Arquivo JSON vazio criado: {self.file_path}")
        except Exception as e:
            print(f"Error creating JSON file: {e}")
        
class Pandalizer:

    def __init__(self, file_path):
        self.df = pd.read_excel(file_path)

    def Backup_Extracted_Data(self, output_path):  # Added 'self' parameter
        # Implementation for backing up data goes here
        pass
        
    def DF_to_Excel(self, output_path):
        try:
            self.df.to_excel(output_path, index=False)
            print(f"DataFrame salvo em: {output_path}")
        except Exception as e:
            print(f"Erro ao salvar o DataFrame: {e}")
        
    def get_DF(self):
        print(self.df.head(10))
        return self.df
                      
if __name__ == "__main__":
    path_to = DIR()  # Assuming DIR returns an object with `xl_file` attribute
    Create_ExtractedData = EmptyFileMaker(path_to.xl_file)
    Create_Timestamp = EmptyFileMaker(path_to.timestamp)
    
    Create_ExtractedData.empty_file()
    Create_Timestamp.empty_file()
    # Uncomment and implement below as needed
    df_extracted_data = Pandalizer(path_to.xl_file)
    # df_extracted_data.DF_to_Excel('./df_extracted_data2.xlsx')
    # df_extracted_data.get_DF()
