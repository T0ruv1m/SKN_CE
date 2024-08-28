import os
import tkinter as tk
from tkinter import ttk, scrolledtext
import threading
from config_tools import DIR
from pdf_merge_routines import start_merging_routine


class PDFMergerApp:
    
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Merger")

        # Configure the main window
        self.root.geometry('850x400')

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
        threading.Thread(target=self.run_merging).start()

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


def main():
    root = tk.Tk()
    app = PDFMergerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
