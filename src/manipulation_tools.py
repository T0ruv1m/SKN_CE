import pandas as pd
import os
from config import DIR

class EmptyFileMaker:

    def __init__(self, file_path):
        self.file_path = file_path 
        self.extension = os.path.splitext(file_path)[1].lower()

    def empty_file(self):
        if not os.path.exists(self.file_path):
            if self.extension == '.xlsx':
                self.empty_excel()
            elif self.extension == '.txt':
                self.empty_txt()
        else:
            print(f"Arquivo {self.file_path} existente")

    def empty_excel(self):
        try:
            df = pd.DataFrame()
            df.to_excel(self.file_path, index=False)
            print(f"Arquivo criado: {self.file_path}")
        except Exception as e:
            print(f"Erro ao criar o arquivo Excel: {e}")
    
    def empty_txt(self):
        try:
            open(self.file_path)
        except Exception as e:
            print(f"Erro ao criar o arquivo txt: {e}")

class DF_Excellerator:
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
    path_of = DIR()  # Assuming DIR returns an object with `xl_file` attribute
    Create_Excel = EmptyFileMaker(path_of.xl_file)
    Create_Excel.empty_file()
    Create_txt = EmptyFileMaker(dir.timestamp)
    # Uncomment and implement below as needed
    df_extracted_data = DF_Excellerator(path_of.xl_file)
    # df_extracted_data.DF_to_Excel('./df_extracted_data2.xlsx')
    # df_extracted_data.get_DF()
