import os
import pandas as pd
import json
from pypdf import PdfWriter


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


def find_and_merge_pdfs(excel_file, folder_path, column1, column2, output_folder, year_column, folder_column, suffix_column1, suffix_column2, nNF_column, abbrev_length=5, log_callback=None, progress_callback=None, total_files=0, missing_files_set=None, merged_files_json=None):
    """Finds, merges, and names PDF files based on Excel data."""

    try:
        # Load already merged files from JSON if it exists
        if merged_files_json and os.path.exists(merged_files_json):
            with open(merged_files_json, 'r') as json_file:
                merged_files = json.load(json_file)
        else:
            merged_files = []

        df = pd.read_excel(excel_file)
        pdf_files = {}

        for root, dirs, files in os.walk(folder_path):
            for file in files:
                if file.endswith(".pdf"):
                    file_name = os.path.splitext(file)[0]
                    file_path = os.path.join(root, file)
                    pdf_files[file_name] = file_path

        processed_files = 0

        for index, row in df.iterrows():
            try:
                if row.get('index') == 2:
                    continue
                
                file1 = str(row[column1])
                file2 = str(row[column2])
                file1_last8 = file1[-8:] if len(file1) >= 8 else file1
                file2_last8 = file2[-8:] if len(file2) >= 8 else file2
                
                year = str(row[year_column])
                subfolder_name = str(row[folder_column])

                suffix1 = str(row[suffix_column1])[:abbrev_length]
                value = row[suffix_column2]
                suffix2 = f"{value:.2f}".replace('.', '-')

                nNF_value = row.get(nNF_column)
                nNF_value = 'sem-numero' if pd.isna(nNF_value) or nNF_value == '' else int(nNF_value)

                output_filename = f"{suffix1}_{nNF_value}_{file1_last8}_{file2_last8}_{suffix2}.pdf"
                
                if output_filename in merged_files:
                    if log_callback:
                        log_callback(f"Pulando {output_filename} já mesclado.")
                    continue
                
                file1_path = pdf_files.get(file1)
                file2_path = pdf_files.get(file2)

                if file1_path and file2_path:
                    year_path = os.path.join(output_folder, year)
                    os.makedirs(year_path, exist_ok=True)

                    subfolder_path = os.path.join(year_path, subfolder_name)
                    os.makedirs(subfolder_path, exist_ok=True)

                    output_path = os.path.join(subfolder_path, output_filename)
                    merge_pdfs([file1_path, file2_path], output_path)
                    
                    merged_files.append(output_filename)  # Track the merged file
                    if log_callback:
                        log_callback(f"Mesclando {file1} e {file2}")
                    
                    processed_files += 2

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

        # Save the updated list of merged files to JSON
        with open(merged_files_json, 'w') as json_file:
            json.dump(merged_files, json_file)

        if log_callback:
            log_callback("Mesclagem concluída.")

    except Exception as e:
        if log_callback:
            log_callback(f"Erro ao buscar e mesclar PDFs: {e}")
