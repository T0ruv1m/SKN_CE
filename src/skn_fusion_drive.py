import os
import pandas as pd
from pypdf import PdfWriter
import tkinter as tk
from tkinter import ttk, scrolledtext
import threading
from config_tools import DIR

class PDFMergerApp:

    def __init__(self, root):
        self.root = root
        self.root.title("PDF Merger")

        # Configure the main window
        self.root.geometry('600x400')

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
            # Paths and columns configuration
            dir = DIR()
            script_dir = os.path.dirname(os.path.abspath(__file__))

            excel_file = dir.xl_combi
            folder_path = dir.gestor_data
            output_folder = os.path.join(script_dir, '../docs/mesclados')
            column1 = 'chNF'
            column2 = 'chNTR'
            year_column = 'Ano'  # New column for year
            folder_column = 'xMun'
            suffix_column1 = 'Fornecedor'
            suffix_column2 = 'vProd'
            nNF_column = 'nNF'  # Column for integer value prefix

            # Ensure the output folder exists
            os.makedirs(output_folder, exist_ok=True)

            # Get the total number of files in the input folder
            total_files = len([f for f in os.listdir(folder_path) if f.endswith(".pdf")])

            # Initialize the progress bar
            self.update_progress(0, total_files)

            # Start the merging process
            self.missing_files = set()  # Initialize set to track missing files
            find_and_merge_pdfs(
                excel_file, folder_path, column1, column2, output_folder,
                year_column, folder_column, suffix_column1, suffix_column2, nNF_column,  # Pass the new column
                abbrev_length=3, log_callback=self.log,
                progress_callback=self.update_progress, total_files=total_files,
                missing_files_set=self.missing_files  # Pass the missing files set
            )

            # Log the total number of missing files after merging
            self.log(f"Total de arquivos PDF ausentes: {len(self.missing_files)}")

        except Exception as e:
            self.log(f"Erro ao iniciar o processo de mesclagem: {e}")

def merge_pdfs(pdf_list, output_path):
    """Merges PDF files from pdf_list into a single PDF at output_path."""
    try:
        merger = PdfWriter()
        for pdf in pdf_list:
            merger.append(pdf)
        merger.write(output_path)
        merger.close()
    except Exception as e:
        raise RuntimeError(f"Erro ao mesclar PDFs: {e}")

def count_pdfs_in_subfolders(base_folder):
    """Counts the number of PDFs in each subfolder of base_folder."""
    folder_counts = {}
    try:
        for root, dirs, files in os.walk(base_folder):
            for dir in dirs:
                subfolder_path = os.path.join(root, dir)
                num_files = len([f for f in os.listdir(subfolder_path) if f.endswith(".pdf")])
                folder_counts[subfolder_path] = num_files
    except Exception as e:
        raise RuntimeError(f"Erro ao contar PDFs nas subpastas: {e}")
    return folder_counts

def find_and_merge_pdfs(excel_file, folder_path, column1, column2, output_folder, year_column, folder_column, suffix_column1, suffix_column2, nNF_column, abbrev_length=5, log_callback=None, progress_callback=None, total_files=0, missing_files_set=None):
    """Finds, merges, and names PDF files based on Excel data."""
    try:
        # Read the Excel file
        df = pd.read_excel(excel_file)

        # Create a dictionary to map file names to their full paths
        pdf_files = {}

        # Traverse directories and subdirectories to find all PDF files
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                if file.endswith(".pdf"):
                    file_name = os.path.splitext(file)[0]  # Remove the .pdf extension
                    file_path = os.path.join(root, file)
                    pdf_files[file_name] = file_path

        # Prepare progress bar
        processed_files = 0

        # Iterate over each row in the DataFrame
        for index, row in df.iterrows():
            # Construct the expected file name from Excel columns
            expected_name = f"{row[nNF_column]}_{row[column1]}_{row[column2]}_{row[year_column]}_{row[folder_column]}_{row[suffix_column1]}_{row[suffix_column2]}".replace(" ", "").lower()[:abbrev_length]

            # Find matching PDF files
            matching_files = [path for name, path in pdf_files.items() if expected_name in name]

            if matching_files:
                # Merge the PDFs and save the output
                output_file = os.path.join(output_folder, f"{expected_name}.pdf")
                merge_pdfs(matching_files, output_file)
                if log_callback:
                    log_callback(f"Merged PDF saved to: {output_file}")
            else:
                # Log missing files
                missing_files_set.add(expected_name)
                if log_callback:
                    log_callback(f"Missing PDF for: {expected_name}")

            # Update the progress bar
            processed_files += 1
            if progress_callback:
                progress_callback(processed_files, total_files)

    except Exception as e:
        if log_callback:
            log_callback(f"Erro ao mesclar arquivos para a linha {index}: {e}")

# Get the current script's directory
script_dir = os.path.dirname(os.path.abspath(__file__))

def main():
    root = tk.Tk()
    app = PDFMergerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
