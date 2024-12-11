import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import pikepdf
import os

class PDFToolbox:
    def __init__(self, master):
        self.master = master
        master.title("pyDfActions")
        master.geometry("500x400")

        self.frame = tk.Frame(master)
        self.frame.pack(padx=20, pady=20, expand=True, fill=tk.BOTH)

        self.create_buttons()

    def create_buttons(self):
        button_configs = [
            ("Merge PDFs", self.merge_pdfs),
            ("Split PDF", self.split_pdf),
            ("Compress PDF", self.compress_pdf)
        ]

        for text, command in button_configs:
            btn = tk.Button(self.frame, text=text, command=command, 
                            width=25, height=2)
            btn.pack(pady=10)

    def merge_pdfs(self):
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

    def split_pdf(self):
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

def main():
    root = tk.Tk()
    app = PDFToolbox(root)
    root.mainloop()

if __name__ == "__main__":
    main()