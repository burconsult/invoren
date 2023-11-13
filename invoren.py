import tkinter as tk
from tkinterdnd2 import DND_FILES, TkinterDnD
import pdfplumber
import re
import os
from datetime import datetime

def extract_invoice_number(pdf_path):
    try:
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
    except Exception as e:
        update_system_error(f"Error in extract_invoice_number: {e}")
        return None

def rename_pdf_file(file_path, invoice_number):
    try:
        directory, _ = os.path.split(file_path)
        current_date = datetime.now().strftime('%d_%m_%Y')
        new_filename = f"Invoice_{invoice_number}_{current_date}.pdf"
        new_path = os.path.join(directory, new_filename)
        
        # Check if file already exists and modify the name if necessary
        counter = 1
        while os.path.exists(new_path):
            new_filename = f"Invoice_{invoice_number}_{current_date}_{counter}.pdf"
            new_path = os.path.join(directory, new_filename)
            counter += 1

        os.rename(file_path, new_path)
        return new_filename
    except Exception as e:
        update_system_error(f"Error in rename_pdf_file: {e}")
        return None

def drop(event):
    dropped_files = event.data.split()
    output_text.set("")  # Clear output at the beginning of a new drop event
    if len(dropped_files) > 5:
        append_to_output("Error: Please drop 5 or fewer files.")
        return

    for file_path in dropped_files:
        file_path = file_path.strip()  # Remove leading/trailing white spaces
        # Correctly handle file paths with spaces and special characters
        if os.path.exists(file_path):
            if file_path.lower().endswith('.pdf'):
                process_pdf(file_path)
            else:
                append_to_output(f"Error: {file_path} is not a PDF file.\n")
        else:
            append_to_output(f"Error: File {file_path} does not exist.\n")

def process_pdf(file_path):
    extracted_invoice_number = extract_invoice_number(file_path)
    if extracted_invoice_number:
        new_file_path = rename_pdf_file(file_path, extracted_invoice_number)
        if new_file_path:
            append_to_output(f"Processed: {file_path}\nInvoice: {extracted_invoice_number}\nNew Filename: {new_file_path}\n\n")
        else:
            append_to_output(f"Error: Could not rename {file_path}.\n\n")
    else:
        append_to_output(f"Error: No relevant invoice info found in {file_path}.\n\n")

def update_output(text):
    output_text.set(text)

def append_to_output(text):
    current_text = output_text.get()
    new_text = current_text + text if current_text else text
    output_text.set(new_text)

def update_system_error(text):
    system_error_text.set(text)

root = TkinterDnD.Tk()
root.title("Invoren - PDF Invoice Renamer")
root.geometry("600x400")

label = tk.Label(root, text="Drag and drop your PDF files here!")
label.pack(pady=10)

output_text = tk.StringVar()
output_label = tk.Label(root, textvariable=output_text, justify=tk.LEFT, wraplength=500)
output_label.pack(pady=10)

system_error_text = tk.StringVar()
system_error_label = tk.Label(root, textvariable=system_error_text, fg="red", justify=tk.LEFT, wraplength=500)
system_error_label.pack(pady=10)

root.drop_target_register(DND_FILES)
root.dnd_bind('<<Drop>>', drop)

root.mainloop()