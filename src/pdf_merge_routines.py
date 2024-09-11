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

def start_merging_routine(dir, log_callback=None, progress_callback=None):
    """Starts the PDF merging process with error handling."""
    try:
        excel_file = dir.xl_combi
        folder_path_gestor = dir.gestor_data
        folder_path_chNTR = dir.chNTR_data  
        output_folder = dir.mesc
        column1 = 'chNF'
        column2 = 'chNTR'
        year_column = 'Ano'
        folder_column = 'Municipio'
        suffix_column1 = 'FOR'
        suffix_column2 = 'Valor'
        nNF_column = 'nNF'
        merged_files_json = dir.merged_files_json
        
        os.makedirs(output_folder, exist_ok=True)
    
        total_files = len([f for f in os.listdir(folder_path_gestor) if f.endswith(".pdf")]) + \
                      len([f for f in os.listdir(folder_path_chNTR) if f.endswith(".pdf")])
        if progress_callback:
            progress_callback(0, total_files)
        
        missing_files = set()
        find_and_merge_pdfs(
            excel_file, folder_path_gestor, folder_path_chNTR, column1, column2, output_folder, 
            year_column, folder_column, suffix_column1, suffix_column2, nNF_column, 
            abbrev_length=3, log_callback=log_callback, 
            progress_callback=progress_callback, total_files=total_files,
            missing_files_set=missing_files, merged_files_json=merged_files_json
        )
        if log_callback:
            log_callback(f"Total de arquivos PDF ausentes: {len(missing_files)}")

    except Exception as e:
        if log_callback:
            log_callback(f"Erro ao iniciar o processo de mesclagem: {e}")

def find_and_merge_pdfs(excel_file, folder_path_gestor, folder_path_chNTR, column1, column2, output_folder, year_column, folder_column, suffix_column1, suffix_column2, nNF_column, abbrev_length=5, log_callback=None, progress_callback=None, total_files=0, missing_files_set=None, merged_files_json=None):
    """Finds, merges, and names PDF files based on Excel data."""
    
    try:
        # Load already merged files from JSON if it exists
        if merged_files_json and os.path.exists(merged_files_json):
            with open(merged_files_json, 'r') as json_file:
                merged_files = json.load(json_file)
        else:
            merged_files = []

        df = pd.read_excel(excel_file)
        pdf_files_gestor = {}
        pdf_files_chNTR = {}
        complementary_files = {}

        # Walk through gestor files
        for root, dirs, files in os.walk(folder_path_gestor):
            for file in files:
                if file.endswith(".pdf"):
                    file_name = os.path.splitext(file)[0]
                    file_path = os.path.join(root, file)
                    
                    if "CCe" in os.path.basename(root):
                        complementary_files[file_name] = file_path
                    else:
                        pdf_files_gestor[file_name] = file_path

        # Walk through chNTR files
        for root, dirs, files in os.walk(folder_path_chNTR):
            for file in files:
                if file.endswith(".pdf"):
                    file_name = os.path.splitext(file)[0]
                    # Remove "-nfe" suffix if present
                    if file_name.endswith("-nfe"):
                        file_name = file_name[:-4]
                    file_path = os.path.join(root, file)
                    pdf_files_chNTR[file_name] = file_path

        processed_files = 0
        successfully_merged_count = 0
        total_rows = len(df)

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

                suffix1 = str(row[suffix_column1])[:abbrev_length].upper()
                value = row[suffix_column2]
                suffix2 = f"{value:.2f}".replace('.', '-')

                nNF_value = row.get(nNF_column)
                nNF_value = 'sem-numero' if pd.isna(nNF_value) or nNF_value == '' else int(nNF_value)

                output_filename = f"{suffix1}_{nNF_value}_{file1_last8}_{file2_last8}_{suffix2}.pdf"
                
                if output_filename in merged_files:
                    print(f"Pulando {output_filename} já mesclado.")
                    continue
                
                file1_path = pdf_files_gestor.get(file1)
                file2_path = pdf_files_chNTR.get(file2)  # This now matches without "-nfe" suffix

                file1_complementary = complementary_files.get(file1)
                
                pdf_list = []
                if file1_path:
                    pdf_list.append(file1_path)
                if file2_path:
                    pdf_list.append(file2_path)
                
                if file1_complementary:
                    pdf_list.append(file1_complementary)

                if pdf_list:
                    year_path = os.path.join(output_folder, year)
                    os.makedirs(year_path, exist_ok=True)

                    subfolder_path = os.path.join(year_path, subfolder_name)
                    os.makedirs(subfolder_path, exist_ok=True)

                    output_path = os.path.join(subfolder_path, output_filename)
                    merge_pdfs(pdf_list, output_path)
                    
                    merged_files.append(output_filename)
                    successfully_merged_count += 1
                    print(f"Mesclando {file1} e {file2} {'com arquivos complementares' if len(pdf_list) > 2 else ''}")
                    
                    processed_files += len(pdf_list)

                    if progress_callback:
                        progress_callback(index + 1, total_rows)
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
            log_callback(f"Mesclagem concluída. Total de arquivos mesclados nesta execução: {successfully_merged_count}")

    except Exception as e:
        if log_callback:
            log_callback(f"Erro ao buscar e mesclar PDFs: {e}")