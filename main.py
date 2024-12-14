import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import pikepdf
import os
import PyPDF2

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    DRAG_DROP_AVAILABLE = True
except ImportError:
    DRAG_DROP_AVAILABLE = False
    TkinterDnD = tk

class PDFToolbox:
    def __init__(self, master):
        self.master = master
        master.title("pyDfActions")
        master.geometry("600x500")

        self.frame = tk.Frame(master)
        self.frame.pack(padx=20, pady=20, expand=True, fill=tk.BOTH)

        self.create_buttons()
        
        if DRAG_DROP_AVAILABLE:
            master.drop_target_register(DND_FILES)
            master.dnd_bind('<<Drop>>', self.on_drop)

    def create_buttons(self):
        button_configs = [
            ("Merge PDFs", self.merge_pdfs),
            ("Split PDF", self.split_pdf),
            ("Compress PDF", self.compress_pdf),
            ("Password Protect PDF", self.add_password),
            ("Remove PDF Password", self.remove_password),
            ("Extract PDF Text", self.extract_text)
        ]

        for text, command in button_configs:
            btn = tk.Button(self.frame, text=text, command=command, 
                            width=25, height=2)
            btn.pack(pady=10)

    def show_pdf_options(self, file_path):
        def create_option_window():
            option_window = tk.Toplevel(self.master)
            option_window.title("PDF Actions")
            option_window.geometry("300x400")

            def create_action_button(text, command):
                btn = tk.Button(option_window, text=text, command=lambda: [command(file_path), option_window.destroy()], 
                                width=25, height=2)
                btn.pack(pady=10)

            create_action_button("Merge PDFs", self.merge_pdfs)
            create_action_button("Split PDF", self.split_pdf)
            create_action_button("Compress PDF", self.compress_pdf)
            create_action_button("Password Protect PDF", self.add_password)
            create_action_button("Remove PDF Password", self.remove_password)
            create_action_button("Extract PDF Text", self.extract_text)

            btn_cancel = tk.Button(option_window, text="Cancel", command=option_window.destroy, 
                                   width=25, height=2)
            btn_cancel.pack(pady=10)

        create_option_window()

    def on_drop(self, event):
        try:
            files = event.data.split() if hasattr(event, 'data') else []
            cleaned_files = [f.strip('{}') for f in files]
            
            valid_pdfs = [f for f in cleaned_files if f.lower().endswith('.pdf')]
            
            if not valid_pdfs:
                messagebox.showerror("Error", "Please drop PDF files.")
                return
            
            if len(valid_pdfs) > 1:
                merge_choice = messagebox.askyesno(
                    "Multiple PDFs Dropped", 
                    "Do you want to merge these PDFs?"
                )
                if merge_choice:
                    self.merge_pdfs(valid_pdfs)
            else:
                self.show_pdf_options(valid_pdfs[0])
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def merge_pdfs(self, file_paths=None):
        if file_paths is None:
            file_paths = filedialog.askopenfilenames(
                title="Select PDF files to merge",
                filetypes=[("PDF files", "*.pdf")]
            )

        if not file_paths:
            return

        output_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")]
        )

        if not output_path:
            return

        try:
            pdf = pikepdf.Pdf.new()
            
            for path in file_paths:
                src = pikepdf.Pdf.open(path)
                pdf.pages.extend(src.pages)
            
            pdf.save(output_path)
            messagebox.showinfo("Success", f"PDFs merged successfully to {output_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to merge PDFs: {str(e)}")

    def split_pdf(self, file_path=None):
        if file_path is None:
            file_path = filedialog.askopenfilename(
                title="Select PDF to split",
                filetypes=[("PDF files", "*.pdf")]
            )

        if not file_path:
            return

        page_range = simpledialog.askstring(
            "Split PDF", 
            "Enter page range to extract (e.g., '1-3,5,7-9'):"
        )

        if not page_range:
            return

        output_dir = filedialog.askdirectory(
            title="Select directory to save split PDFs"
        )

        if not output_dir:
            return

        try:
            pdf = pikepdf.Pdf.open(file_path)
            
            pages_to_extract = self._parse_page_range(page_range, len(pdf.pages))

            for page_num in pages_to_extract:
                output_pdf = pikepdf.Pdf.new()
                output_pdf.pages.append(pdf.pages[page_num-1])
                
                output_path = os.path.join(output_dir, f"page_{page_num}.pdf")
                output_pdf.save(output_path)

            messagebox.showinfo("Success", f"PDF split successfully. Files saved in {output_dir}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to split PDF: {str(e)}")

    def _parse_page_range(self, range_str, total_pages):
        pages = set()
        for part in range_str.split(','):
            if '-' in part:
                start, end = map(int, part.split('-'))
                pages.update(range(start, end+1))
            else:
                pages.add(int(part))
        
        return sorted(page for page in pages if 1 <= page <= total_pages)

    def compress_pdf(self):
        file_path = filedialog.askopenfilename(
            title="Select PDF to compress",
            filetypes=[("PDF files", "*.pdf")]
        )

        if not file_path:
            return

        if not os.path.exists(file_path):
            messagebox.showerror("Error", f"File {file_path} does not exist.")
            return

        if not file_path.lower().endswith('.pdf'):
            messagebox.showerror("Error", "Input file must be a PDF.")
            return

        directory = os.path.dirname(file_path)
        filename = os.path.basename(file_path)
        
        output_path = filedialog.asksaveasfilename(
            initialfile=f'compressed_{filename}',
            defaultextension=".pdf",
            initialdir=directory,
            filetypes=[("PDF files", "*.pdf")]
        )

        if not output_path:
            return

        try:
            pdf = pikepdf.Pdf.open(file_path)
            
            pdf.save(output_path)
            
            original_size = os.path.getsize(file_path)
            compressed_size = os.path.getsize(output_path)
            
            messagebox.showinfo(
                "Compression Complete", 
                f"PDF compressed successfully.\n"
                f"Original Size: {original_size/1024:.2f} KB\n"
                f"Compressed Size: {compressed_size/1024:.2f} KB\n"
                f"Size Reduction: {(1 - compressed_size/original_size)*100:.2f}%"
            )
        except Exception as e:
            messagebox.showerror("Error", f"Failed to compress PDF: {str(e)}")

    def add_password(self):
        file_path = filedialog.askopenfilename(
            title="Select PDF to password protect",
            filetypes=[("PDF files", "*.pdf")]
        )

        if not file_path:
            return

        password = simpledialog.askstring(
            "Password Protection", 
            "Enter password to protect the PDF:", 
            show='*'
        )

        if not password:
            return

        output_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")]
        )

        if not output_path:
            return

        try:
            pdf = pikepdf.Pdf.open(file_path)
            
            pdf.save(
                output_path, 
                encryption=pikepdf.Encryption(
                    owner=password, 
                    user=password
                )
            )
            
            messagebox.showinfo("Success", f"PDF password protected successfully at {output_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to password protect PDF: {str(e)}")

    def remove_password(self):
        file_path = filedialog.askopenfilename(
            title="Select Password-Protected PDF",
            filetypes=[("PDF files", "*.pdf")]
        )

        if not file_path:
            return

        password = simpledialog.askstring(
            "Remove Password", 
            "Enter the current PDF password:", 
            show='*'
        )

        if not password:
            return

        output_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")]
        )

        if not output_path:
            return

        try:
            pdf = pikepdf.Pdf.open(file_path, password=password)
            
            pdf.save(output_path)
            
            messagebox.showinfo("Success", f"PDF password removed successfully at {output_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to remove PDF password: {str(e)}")

    def extract_text(self, file_path=None):
        if file_path is None:
            file_path = filedialog.askopenfilename(
                title="Select PDF to extract text",
                filetypes=[("PDF files", "*.pdf")]
            )

        if not file_path:
            return

        try:
            with open(file_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                
                text_content = ""
                for page in pdf_reader.pages:
                    text_content += page.extract_text() + "\n\n"

            output_path = filedialog.asksaveasfilename(
                defaultextension=".txt",
                filetypes=[("Text files", "*.txt")],
                initialfile="extracted_text.txt"
            )

            if output_path:
                with open(output_path, 'w', encoding='utf-8') as text_file:
                    text_file.write(text_content)

                messagebox.showinfo(
                    "Text Extraction", 
                    f"Text extracted successfully to {output_path}"
                )

        except Exception as e:
            messagebox.showerror("Error", f"Failed to extract text: {str(e)}")

def main():
    if DRAG_DROP_AVAILABLE:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    
    app = PDFToolbox(root)
    root.mainloop()

if __name__ == "__main__":
    main()
