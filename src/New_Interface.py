import os
import tkinter as tk
from tkinter import ttk, scrolledtext
import threading
from config_tools import DIR
from XML_Handler import XMLProcessor, ExcelMerger
from pdf_merge_routines import start_merging_routine
from xml_cache_controller import XMLFileProcessor, CacheManager

class PDFMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Multi-Routine Processor")

        # Configure the main window
        self.root.geometry('750x400')

        # Create a logger
        self.logger_frame = tk.Frame(self.root)
        self.logger_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        self.log_text = scrolledtext.ScrolledText(self.logger_frame, wrap=tk.WORD, height=15, width=70, bg='#2A2B2A', fg='white')
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Create a frame for buttons
        self.button_frame = tk.Frame(self.root)
        self.button_frame.pack(fill=tk.X, pady=5)

        # Create a toggle button for auto-pipeline
        self.auto_pipeline = tk.BooleanVar()
        self.auto_pipeline_button = tk.Checkbutton(self.button_frame, text="Auto-Pipeline", variable=self.auto_pipeline)
        self.auto_pipeline_button.pack(side=tk.LEFT, padx=5)

        # Create buttons for different routines, side by side
        self.clear_cache_button = tk.Button(self.button_frame, text="Clear Cache", command=self.clear_cache_thread)
        self.clear_cache_button.pack(side=tk.LEFT, padx=5)
        
        self.scan_xml_files_button = tk.Button(self.button_frame, text="Varredura XML", command=self.start_scan_xml_thread)
        self.scan_xml_files_button.pack(side=tk.LEFT, padx=5)

        self.xml_gestor_button = tk.Button(self.button_frame, text="Extrair XML|Gestor", command=self.start_xml_gestor_thread)
        self.xml_gestor_button.pack(side=tk.LEFT, padx=5)

        self.xml_compras_button = tk.Button(self.button_frame, text="Extrair XML|Compras", command=self.start_xml_compras_thread)
        self.xml_compras_button.pack(side=tk.LEFT, padx=5)

        self.excel_merge_button = tk.Button(self.button_frame, text="Mesclar Excel", command=self.start_excel_merge_thread)
        self.excel_merge_button.pack(side=tk.LEFT, padx=5)

        self.merge_button = tk.Button(self.button_frame, text="Fusão PDFs", command=self.start_merge_thread)
        self.merge_button.pack(side=tk.LEFT, padx=5)

        # Create a progress bar (only for PDF merging)
        self.progress_frame = tk.Frame(self.root)
        self.progress_frame.pack(fill=tk.X, padx=10, pady=5)
        self.progress = ttk.Progressbar(self.progress_frame, orient=tk.HORIZONTAL, length=700, mode='determinate')
        self.progress.pack(pady=5)

    def log(self, message):
        """Logs a message to the GUI log text area."""
        self.log_text.insert(tk.END, message + '\n')
        self.log_text.yview(tk.END)

    def update_progress(self, value, maximum):
        """Updates the progress bar value."""
        self.progress['value'] = value
        self.progress['maximum'] = maximum
        self.root.update_idletasks()

    def disable_buttons(self):
        """Disables all buttons and changes their appearance to gray."""
        buttons = [self.merge_button, self.xml_gestor_button, self.xml_compras_button, self.excel_merge_button, self.scan_xml_files_button]
        for button in buttons:
            button.config(state=tk.DISABLED, bg='#A9A9A9')

    def enable_buttons(self):
        """Enables all buttons and restores their original appearance."""
        buttons = [self.merge_button, self.xml_gestor_button, self.xml_compras_button, self.excel_merge_button, self.scan_xml_files_button]
        for button in buttons:
            button.config(state=tk.NORMAL, bg='SystemButtonFace')

    def start_merge_thread(self):
        """Starts the PDF merging process in a separate thread."""
        self.disable_buttons()
        threading.Thread(target=self.run_merging).start()

    def start_xml_gestor_thread(self):
        """Starts the XML processing for Gestor in a separate thread (no progress bar)."""
        self.disable_buttons()
        threading.Thread(target=self.run_xml_gestor_processing).start()

    def start_xml_compras_thread(self):
        """Starts the XML processing for Compras in a separate thread (no progress bar)."""
        self.disable_buttons()
        threading.Thread(target=self.run_xml_compras_processing).start()

    def start_excel_merge_thread(self):
        """Starts the Excel merging process in a separate thread (no progress bar)."""
        self.disable_buttons()
        threading.Thread(target=self.run_excel_merging).start()

    def start_scan_xml_thread(self):
        """Starts the XML scanning process in a separate thread (no progress bar)."""
        self.disable_buttons()
        threading.Thread(target=self.run_scan_xml_files).start()

    def run_merging(self):
        """Runs the PDF merging process."""
        try:
            dir = DIR()
            start_merging_routine(
                dir=dir,
                log_callback=self.log,
                progress_callback=self.update_progress
            )
        except Exception as e:
            self.log(f"Erro ao iniciar o processo de mesclagem: {e}")
        finally:
            self.enable_buttons()
            self.check_auto_pipeline(self.merge_button)

    def run_xml_gestor_processing(self):
        """Runs the XML processing for Gestor."""
        try:
            processor = XMLProcessor()
            path_to = DIR()
            new_gestor = path_to.new_gestor 
            xl_gestor = path_to.xl_gestor

            new_files_gestor = processor.load_new_files_list(new_gestor)
            existing_data_gestor = processor.load_existing_data(xl_gestor, ['chNTR', 'nNF', 'dhEmi', 'xFant'])
            xml_data_gestor = processor.build_xml_file_mapping(new_files_gestor, existing_data_gestor, extraction_type='gestor')
            processor.save_xml_data_to_excel(xml_data_gestor, xl_gestor, ['chNTR', 'nNF', 'dhEmi', 'xFant'])

            self.log("XML Gestor processing completed successfully.")
        except Exception as e:
            self.log(f"Error processing XML Gestor: {e}")
        finally:
            self.enable_buttons()
            self.check_auto_pipeline(self.xml_gestor_button)

    def run_xml_compras_processing(self):
        """Runs the XML processing for Compras."""
        try:
            processor = XMLProcessor()
            path_to = DIR()
            new_compras = path_to.new_compras
            xl_compras = path_to.xl_compras

            new_files_compras = processor.load_new_files_list(new_compras)
            existing_data_compras = processor.load_existing_data(xl_compras, ['file_name', 'chNF', 'chNTR', 'xMun', 'vProd'])
            xml_data_compras = processor.build_xml_file_mapping(new_files_compras, existing_data_compras, extraction_type='compras')
            processor.save_xml_data_to_excel(xml_data_compras, xl_compras, ['file_name', 'chNF', 'chNTR', 'xMun', 'vProd'])

            self.log("XML Compras processing completed successfully.")
        except Exception as e:
            self.log(f"Error processing XML Compras: {e}")
        finally:
            self.enable_buttons()
            self.check_auto_pipeline(self.xml_compras_button)

    def run_excel_merging(self):
        """Runs the Excel merging process."""
        try:
            path_to = DIR()
            file1 = path_to.xl_compras
            file2 = path_to.xl_gestor
            column_to_merge_on = 'chNTR'
            output_file = path_to.xl_combi

            merger = ExcelMerger(file1, file2, column_to_merge_on, output_file)
            merger.merge_excel_files()

            self.log("Excel merging completed successfully.")
        except Exception as e:
            self.log(f"Error merging Excel files: {e}")
        finally:
            self.enable_buttons()
            self.check_auto_pipeline(self.excel_merge_button)

    def run_scan_xml_files(self):
        """Runs the XML scanning process for new or modified files."""
        try:
            path_to = DIR()
            processor = XMLFileProcessor(path_to.xml_data, path_to.cache_compras)
            processor.process_new_files(path_to.new_compras)

            gestor_processor = XMLFileProcessor(path_to.gestor_data, path_to.cache_gestor)
            gestor_processor.process_new_files(path_to.new_gestor)

            self.log("XML scanning completed successfully.")
        except Exception as e:
            self.log(f"Error scanning XML files: {e}")
        finally:
            self.enable_buttons()
            self.check_auto_pipeline(self.scan_xml_files_button)

    def check_auto_pipeline(self, current_button):
        """Check if auto-pipeline is enabled and trigger the next button."""
        if self.auto_pipeline.get():
            button_sequence = [
                self.scan_xml_files_button,
                self.xml_gestor_button,
                self.xml_compras_button,
                self.excel_merge_button,
                self.merge_button
            ]
            current_index = button_sequence.index(current_button)
            if current_index < len(button_sequence) - 1:
                next_button = button_sequence[current_index + 1]

                next_button.invoke()

    def clear_cache_thread(self):
        """Starts the cache clearing process in a separate thread."""
        self.disable_buttons()
        threading.Thread(target=self.run_clear_cache).start()

    def run_clear_cache(self):
        """Runs the cache clearing process."""
        try:
            dir = DIR()
            cache = CacheManager()
            cache.clear_cache_files(dir.cache_gestor)
            cache.clear_cache_files(dir.cache_gestor)
            cache.clear_cache_files(dir.new_compras)
            cache.clear_cache_files(dir.new_gestor)

            self.log("Cache clearing completed successfully.")
        except Exception as e:
            self.log(f"Error clearing cache files: {e}")
        finally:
            self.enable_buttons()

def main():
    root = tk.Tk()
    app = PDFMergerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
