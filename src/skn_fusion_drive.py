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

            excel_file = dir.xl_combi
            folder_path = dir.gestor_data
            output_folder = dir.mesc
            column1 = 'chNF'
            column2 = 'chNTR'
            year_column = 'Ano'  # New column for year
            folder_column = 'Municipio'
            suffix_column1 = 'FOR'
            suffix_column2 = 'Valor'
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

        # Iterate over the rows of the DataFrame
        for index, row in df.iterrows():
            try:
                # Skip rows where index value is 2
                if row.get('index') == 2:
                    continue
                
                file1 = str(row[column1])  # Ensure file names are treated as strings
                file2 = str(row[column2])

                # Use only the last 8 characters of file1 and file2 for naming
                file1_last8 = file1[-8:] if len(file1) >= 8 else file1
                file2_last8 = file2[-8:] if len(file2) >= 8 else file2
                
                # Get the year and folder names from specified columns
                year = str(row[year_column])  # Convert year to string for valid folder name
                subfolder_name = str(row[folder_column])  # Convert to string for valid folder name

                # Generate abbreviations for suffixes
                suffix1 = str(row[suffix_column1])[:abbrev_length]  # Abbreviate to first few characters

                # Convert the value in suffix_column2 to a formatted string with two decimal places
                value = row[suffix_column2]
                suffix2 = f"{value:.2f}"  # Format as a string with two decimal places

                # Replace any potential invalid characters in the suffix
                suffix2 = suffix2.replace('.', '-')

                # Get the integer value from nNF_value column or use 'sem-numero' if missing
                nNF_value = row.get(nNF_column)
                if pd.isna(nNF_value) or nNF_value == '':
                    nNF_value = 'sem-numero'
                else:
                    nNF_value = int(nNF_value)  # Ensure it's an integer

                # Construct the full file paths
                file1_path = pdf_files.get(file1)
                file2_path = pdf_files.get(file2)

                # Check if both PDF files exist
                if file1_path and file2_path:
                    # Create year and subfolder if they don't exist
                    year_path = os.path.join(output_folder, year)  # New year folder
                    os.makedirs(year_path, exist_ok=True)

                    subfolder_path = os.path.join(year_path, subfolder_name)
                    os.makedirs(subfolder_path, exist_ok=True)

                    # Define the output file path with the abbreviated suffixes
                    output_filename = f"{suffix1}_{nNF_value}_{file1_last8}_{file2_last8}_{suffix2}.pdf"  # Add nNF_value as prefix
                    output_path = os.path.join(subfolder_path, output_filename)

                    # Merge the PDFs
                    merge_pdfs([file1_path, file2_path], output_path)
                    if log_callback:
                        log_callback(f"Mesclando {file1} e {file2}")
                    
                    processed_files += 2  # Count each file processed

                    # Update progress based on number of processed files
                    if progress_callback:
                        progress_callback(processed_files, total_files)
                else:
                    missing_files = []
                    if not file1_path:
                        missing_files.append(file1)
                    if not file2_path:
                        missing_files.append(file2)
                    
                    if missing_files_set is not None:
                        missing_files_set.update(missing_files)

                    if log_callback:
                        log_callback(f"Arquivos ausentes: {', '.join(missing_files)}")

            except Exception as e:
                if log_callback:
                    log_callback(f"Erro ao mesclar arquivos para a linha {index}: {e}")

        if log_callback:
            log_callback("Mesclagem concluída.")

    except Exception as e:
        if log_callback:
            log_callback(f"Erro ao buscar e mesclar PDFs: {e}")

# Get the current script's directory
script_dir = os.path.dirname(os.path.abspath(__file__))

def main():
    root = tk.Tk()
    app = PDFMergerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
