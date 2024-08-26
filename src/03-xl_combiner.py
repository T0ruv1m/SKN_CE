import pandas as pd
from config_tools import DIR, Pandalizer

# Modificações recentes:
# 1. Criado script para combinar duas tabelas Excel com base em uma coluna correspondente.
# 2. Utilizado o método merge do pandas para realizar o join dos dados.

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
        df2_selected = df2[['chNTR','nNF']]
        
        # Realiza a junção com base na coluna especificada
        merged_df = pd.merge(df1, df2, on=self.merge_column)
        merged_df = pd.merge(merged_df, df2_selected, left_on='chNF', right_on='chNTR')  
        merged_df = merged_df[['xMun','Fornecedor','chNTR_x','nNF_x','chNF','nNF_y','vProd','dhEmi']].rename(columns={
            'xMun':'Municipio',
            'Fornecedor':'FOR',
            'chNTR_x':'chNTR',
            'nNF_x': 'nNTR',
            'nNF_y': 'nNF',
            'vProd': 'Valor',
            'dhEmi': 'Ano'
        })
        
        #df2['nNF'] = df2['nNF'].astype(str)
        #merged2 = pd.merge(df1, df2[['chNTR','nNF']], left_on='chNF',right_on='nNF')
        
        # Salva o resultado em um novo arquivo Excel
        merged_df.to_excel(self.output_file, index=False)
        print(f"Arquivos combinados e salvos em {self.output_file}")

    
# Exemplo de uso
if __name__ == "__main__":
    path_to = DIR()
    file1 = path_to.xl_compras
    file2 = path_to.xl_gestor
    column_to_merge_on = 'chNTR'  # Nome da coluna usada para a junção
    output_file = path_to.xl_combinada
    

    merger = ExcelMerger(file1, file2, column_to_merge_on, output_file)
    merger.merge_excel_files()
    
    