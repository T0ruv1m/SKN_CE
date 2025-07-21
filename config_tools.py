# config.py
import os

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
            'mesc': 'I:\\Drives compartilhados\\Compartilhamento Registro de liquidação',


            #'xml_data': '\\\\Mimitop\\02334933000140',  # UNC path with double backslashes
            #'xml_data': 'I:\\Outros computadores\\MIMI\\Arqs\\02334933000140',
            'xml_data': 'I:\\.shortcut-targets-by-id\\1ghlKQQOndN3wMxNW4qM3pTPDbtrYJmTa\\GestorDFe\\Documentos',
            'gestor_data': 'I:\\.shortcut-targets-by-id\\1ghlKQQOndN3wMxNW4qM3pTPDbtrYJmTa\\GestorDFe\\Documentos',
            #'chNTR_data': '\\\\Mimitop\\PDF',
            #'chNTR_data': 'I:\\Outros computadores\\MIMI\\PDF',
            'chNTR_data': 'I:\\.shortcut-targets-by-id\\1ghlKQQOndN3wMxNW4qM3pTPDbtrYJmTa\\GestorDFe\\Documentos',
            
            'timestamp': './root/timestamp.txt',
            'credentials_file': './root/credentials.json',
            'token_file': './root/token.json',
            'json_object': './root/object.json',
            
            'merged_files_json':'./data/cache_data/merged_files.json',
            'execution_log_csv':'./data/cache_data/execution_log.csv',

            'cache_compras':'./data/cache_data/cache_compras.csv',
            'cache_gestor':'./root/data/cache_data/cache_gestor.csv',
            
            'new_compras':'./data/cache_data/new_compras.csv',
            'new_gestor':'./data/cache_data/new_gestor.csv',

            'xl_compras':'./data/xlsx_data/xl_compras.xlsx',
            'xl_gestor':'./data/xlsx_data/xl_gestor.xlsx',
            'xl_combi':'./data/xlsx_data/xl_combinada.xlsx',  # Removed redundant key
            'xl_consulta': 'I:\\Drives compartilhados\\Compartilhamento Registro de liquidação\\Tabelas\\consulta.xlsx'

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
                   
if __name__ == "__main__":
    path_to = DIR()  # Assuming DIR returns an object with `xl_file` attribute