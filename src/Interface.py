import os
import tkinter as tk
from tkinter import ttk, scrolledtext
import threading
from config_tools import DIR, EmptyFileMaker
from pdf_merge_routines import find_and_merge_pdfs


class PDFMergerApp:
    
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Merger")

        # Configure the main window
        self.root.geometry('900x400')

        # Create a frame for the progress bar
        self.progress_frame = tk.Frame(self.root)
        self.progress_frame.pack(fill=tk.X, padx=10, pady=5)

        # Create a progress bar
        self.progress = ttk.Progressbar(self.progress_frame, orient=tk.HORIZONTAL, length=580, mode='determinate')
        self.progress.pack(pady=5)

        # Create a frame for the logger
        self.logger_frame = tk.Frame(self.root)
        self.logger_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Create a scrolled text widget for the log
        self.log_text = scrolledtext.ScrolledText(self.logger_frame, wrap=tk.WORD, height=15, width=70, bg='#2A2B2A', fg='white')
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Create a button to start the merging process
        self.start_button = tk.Button(self.root, text="Começar Mesclagem", command=self.start_thread)
        self.start_button.pack(pady=10)

    def log(self, message):
        """Logs a message to the GUI log text area."""
        self.log_text.insert(tk.END, message + '\n')
        self.log_text.yview(tk.END)
        
    def update_progress(self, value, maximum):
        """Updates the progress bar value."""
        self.progress['value'] = value
        self.progress['maximum'] = maximum
        self.root.update_idletasks()

    def start_thread(self):
        """Starts the merging process in a separate thread."""
        threading.Thread(target=self.start_merging).start()

    def start_merging(self):
        """Starts the PDF merging process with error handling."""
        try:
            dir = DIR()

            excel_file = dir.xl_combi
            folder_path = dir.gestor_data
            output_folder = dir.mesc
            column1 = 'chNF'
            column2 = 'chNTR'
            year_column = 'Ano'
            folder_column = 'Municipio'
            suffix_column1 = 'FOR'
            suffix_column2 = 'Valor'
            nNF_column = 'nNF'
            merged_files_json = dir.merged_files_json  # Path to the JSON file to track merged files
            
            os.makedirs(output_folder, exist_ok=True)
        
            total_files = len([f for f in os.listdir(folder_path) if f.endswith(".pdf")])
            self.update_progress(0, total_files)
            
            self.missing_files = set()
            find_and_merge_pdfs(
                excel_file, folder_path, column1, column2, output_folder, 
                year_column, folder_column, suffix_column1, suffix_column2, nNF_column, 
                abbrev_length=3, log_callback=self.log, 
                progress_callback=self.update_progress, total_files=total_files,
                missing_files_set=self.missing_files, merged_files_json=merged_files_json  # Pass the merged files JSON
            )
            self.log(f"Total de arquivos PDF ausentes: {len(self.missing_files)}")

        except Exception as e:
            self.log(f"Erro ao iniciar o processo de mesclagem: {e}")


def main():
    root = tk.Tk()
    app = PDFMergerApp(root)
    root.mainloop()


if __name__ == "__main__": 
    file = EmptyFileMaker('./root/merged_files.json')
    file.empty_file()
    main()

