# config.py
import os

class DIR:
    """Classe de configuração para diretórios e variáveis compartilhadas."""
    
    def __init__(self):
        # Diretórios e variáveis de configuração
        
        root = os.path.dirname(os.path.abspath(__file__))
        
        #LOCAL
        
        self.dig = os.path.join(root,'../docs/digitalizados')       # Diretório para input
        self.res = os.path.join(root,'../docs/registro')            # Diretório para resíduo
        self.reg = os.path.join(root,'../docs/residuo')             # Diretório para registro
        self.mesc = os.path.join(root, '../docs/mesclados')         # Diretório para mesclados


        self.xl_file = os.path.join(root,'../data/xlsx_data/extracted_data.xlsx')
        self.gestor_local = 'I:/.shortcut-targets-by-id/1ghlKQQOndN3wMxNW4qM3pTPDbtrYJmTa/GestorDFe/Documentos/CONSORCIO INTERMUNICIPAL DE SAUDE ALTO DAS VERTENTES'
        
        
        self.timestamp = os.path.join(root,'../root/timestamp.txt')
        self.credentials_file = '../root/credentials.json'
        self.token_file = '../root/token.json'
        
        self.xml_files = os.path.join(root, '../data/xml_data')
        

        # CLOUD
        self.drive_cloud = ''
        self.xml_cloud = ''

    def set_input_directory(self, path):
        """Define o diretório de download."""
        self.input_directory = path

    def set_registro_directory(self, path):
        """Define o diretório de upload."""
        self.output_directory = path

    def set_mesclados_directory(self, path):
        """Define o diretório de backup."""
        self.backup_directory = path
