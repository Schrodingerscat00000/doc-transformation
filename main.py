import tkinter as tk
from tkinter import filedialog, messagebox
from docs import extract_tracked_changes, apply_tracked_changes_to_chinese_doc
from alignment import align_changes
from docx import Document
import nltk


def read_docx_text(docx_path):
    doc = Document(docx_path)
    full_text = '\n'.join([p.text for p in doc.paragraphs if p.text.strip()])
    return full_text

class TrackChangesApp:
    def __init__(self, root):
        self.root = root
        root.title("Tracked Changes Bilingual Updater")

        self.eng_doc_path = ''
        self.ch_doc_path = ''

        tk.Label(root, text="English DOCX (with tracked changes)").pack()
        self.eng_label = tk.Label(root, text="No file selected", fg='gray')
        self.eng_label.pack()
        tk.Button(root, text="Select English DOCX", command=self.load_eng_doc).pack()

        tk.Label(root, text="Chinese DOCX (original)").pack()
        self.ch_label = tk.Label(root, text="No file selected", fg='gray')
        self.ch_label.pack()
        tk.Button(root, text="Select Chinese DOCX", command=self.load_ch_doc).pack()

        self.status_label = tk.Label(root, text="")
        self.status_label.pack(pady=10)

        tk.Button(root, text="Process & Save Output", command=self.process_docs).pack(pady=20)

    def load_eng_doc(self):
        path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
        if path:
            self.eng_doc_path = path
            self.eng_label.config(text=path, fg='black')

    def load_ch_doc(self):
        path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
        if path:
            self.ch_doc_path = path
            self.ch_label.config(text=path, fg='black')

    def process_docs(self):
        if not self.eng_doc_path or not self.ch_doc_path:
            messagebox.showerror("Error", "Please select both English and Chinese DOCX files.")
            return

        try:
            self.status_label.config(text="Extracting tracked changes from English DOCX...")
            self.root.update()

            eng_changes = extract_tracked_changes(self.eng_doc_path)

            self.status_label.config(text="Reading full document texts...")
            self.root.update()

            eng_text = read_docx_text(self.eng_doc_path)
            ch_text = read_docx_text(self.ch_doc_path)

            self.status_label.config(text="Aligning English changes with Chinese document...")
            self.root.update()

            mapped_changes = align_changes(eng_changes, eng_text, ch_text)

            self.status_label.config(text="Saving updated Chinese document with tracked changes...")
            self.root.update()

            save_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Documents", "*.docx")])
            if not save_path:
                self.status_label.config(text="Save cancelled.")
                return

            apply_tracked_changes_to_chinese_doc(self.ch_doc_path, mapped_changes, save_path)

            self.status_label.config(text="Process completed successfully!")
            messagebox.showinfo("Success", f"Tracked changes applied and saved to:\n{save_path}")

        except Exception as e:
            self.status_label.config(text="")
            messagebox.showerror("Processing Error", str(e))

if __name__ == '__main__':
    root = tk.Tk()
    app = TrackChangesApp(root)
    root.mainloop()
