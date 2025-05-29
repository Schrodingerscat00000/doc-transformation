# app.py

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import os
from docx_processor import run_document_processing

class DocxUpdaterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Hybrid Document Updater")
        self.root.geometry("650x350")

        main_frame = ttk.Frame(root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # File selection widgets
        self._create_file_selector(main_frame, "English DOCX (with Changes):", 0, self.browse_english)
        self._create_file_selector(main_frame, "Original Chinese DOCX:", 2, self.browse_chinese)
        
        self.eng_path_var = tk.StringVar()
        self.chn_path_var = tk.StringVar()

        self.eng_entry.config(textvariable=self.eng_path_var)
        self.chn_entry.config(textvariable=self.chn_path_var)

        # Process Button
        self.process_button = ttk.Button(main_frame, text="Start Processing", command=self.start_processing)
        self.process_button.grid(row=4, column=0, columnspan=2, pady=20, ipady=5)
        
        # Status Label
        status_frame = ttk.LabelFrame(main_frame, text="Log", padding="10")
        status_frame.grid(row=5, column=0, columnspan=2, sticky="ew", ipady=5)
        main_frame.grid_columnconfigure(0, weight=1)
        
        self.status_label = ttk.Label(status_frame, text="Welcome! Please select your documents.", wraplength=600)
        self.status_label.pack(fill=tk.BOTH, expand=True)

    def _create_file_selector(self, parent, label_text, row_num, command):
        ttk.Label(parent, text=label_text).grid(row=row_num, column=0, sticky="w", pady=(10, 2))
        entry = ttk.Entry(parent, width=70)
        entry.grid(row=row_num + 1, column=0, sticky="ew")
        ttk.Button(parent, text="Browse...", command=command).grid(row=row_num + 1, column=1, padx=5)
        
        if "English" in label_text:
            self.eng_entry = entry
        else:
            self.chn_entry = entry

    def browse_english(self):
        path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
        if path: self.eng_path_var.set(path)

    def browse_chinese(self):
        path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
        if path: self.chn_path_var.set(path)
            
    def update_status(self, message):
        """Thread-safe method to update the status label from the backend."""
        self.root.after(0, lambda: self.status_label.config(text=message))

    def start_processing(self):
        eng_path = self.eng_path_var.get()
        chn_path = self.chn_path_var.get()

        if not (eng_path and chn_path):
            messagebox.showerror("Input Error", "Please select both document files.")
            return

        output_dir = os.path.dirname(chn_path)
        output_filename = os.path.splitext(os.path.basename(chn_path))[0] + "_updated_v2.docx"
        output_path = os.path.join(output_dir, output_filename)
        
        self.process_button.config(state=tk.DISABLED)
        
        # Run the backend logic in a separate thread to keep the GUI responsive
        thread = threading.Thread(
            target=run_document_processing,
            args=(eng_path, chn_path, output_path, self.update_status)
        )
        thread.daemon = True
        thread.start()
        
        self.check_if_done(thread)

    def check_if_done(self, thread):
        """Re-enables the button once the processing thread is complete."""
        if thread.is_alive():
            self.root.after(100, lambda: self.check_if_done(thread))
        else:
            self.process_button.config(state=tk.NORMAL)

if __name__ == "__main__":
    root = tk.Tk()
    app = DocxUpdaterApp(root)
    root.mainloop()