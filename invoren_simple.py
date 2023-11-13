import tkinter as tk
import pdfplumber
import re
import os
from datetime import datetime
from tkinterdnd2 import DND_FILES, TkinterDnD
import os

def extract_invoice_number(pdf_path):
    invoice_number_pattern = r"10\d{2}"  # Matches numbers from 1000 to 1099
    invoice_number = None

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                search_result = re.findall(invoice_number_pattern, text)
                # Assuming the first match is the invoice number
                if search_result:
                    # Filtering out numbers outside the range 1001 to 1099
                    invoice_number = next((num for num in search_result if 1001 <= int(num) <= 1099), None)
                    if invoice_number:
                        break

    return invoice_number

def rename_pdf_file(file_path, invoice_number):
    directory, _ = os.path.split(file_path)
    current_date = datetime.now().strftime('%d_%m_%Y')
    new_filename = f"Invoice_{invoice_number}_{current_date}.pdf"
    new_path = os.path.join(directory, new_filename)
    os.rename(file_path, new_path)
    return new_filename

def drop(event):
    output_label.config(text="Processing...")
    root.update()

    dropped_files = event.data
    for file in dropped_files.split():
        file_path = file.replace("{", "").replace("}", "")  # Clean up file path
        if file_path.lower().endswith('.pdf'):
            process_pdf(file_path)

    output_label.config(text="Finished processing files.")

def process_pdf(file_path):
    extracted_invoice_number = extract_invoice_number(file_path)
    if extracted_invoice_number:
        rename_pdf_file(file_path, extracted_invoice_number)
        new_file_path = os.path.join(os.path.dirname(file_path), f"Invoice_{extracted_invoice_number}_{datetime.now().strftime('%d_%m_%Y')}.pdf")
        update_output(f"Processed: {file_path}\nInvoice: {extracted_invoice_number}\nNew Filename: {new_file_path}")

def update_output(text):
    output_text.set(text)

root = TkinterDnD.Tk()
root.title("PDF Invoice Processor")
root.geometry("600x400")

label = tk.Label(root, text="Drag and drop your PDF files here")
label.pack(pady=10)

output_text = tk.StringVar()
output_label = tk.Label(root, textvariable=output_text, justify=tk.LEFT, wraplength=500)
output_label.pack(pady=10)

root.drop_target_register(DND_FILES)
root.dnd_bind('<<Drop>>', drop)

root.mainloop()