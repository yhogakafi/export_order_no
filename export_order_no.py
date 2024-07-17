import tkinter as tk
from tkinter import filedialog, messagebox
import pdfplumber
import pandas as pd
import re
import os

# Function to extract specific information from PDF
def extract_info_from_pdf(pdf_path, keyword):
    extracted_info = []
    
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                matches = re.findall(rf"{keyword}\s*:\s*(\S+)", text)
                clean_matches = [match.strip() for match in matches]
                extracted_info.extend(clean_matches)
    
    return extracted_info

# Function to save extracted information to Excel
def save_to_excel(data, output_path):
    df = pd.DataFrame(data, columns=['Extracted Information'])
    df.to_excel(output_path, index=False)

# Function to handle button click for processing PDF
def process_pdf():
    pdf_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if pdf_path:
        keyword = 'No.Pesanan'  # Replace with your keyword
        default_output_name = os.path.splitext(os.path.basename(pdf_path))[0] + '-no pesanan.xlsx'
        output_path = filedialog.asksaveasfilename(initialfile=default_output_name, defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if output_path:
            extracted_info = extract_info_from_pdf(pdf_path, keyword)
            save_to_excel(extracted_info, output_path)
            messagebox.showinfo("Extraction Complete", f"Extracted information saved to:\n{output_path}")

# Create GUI window
root = tk.Tk()
root.title("PDF Information Extractor")

# Create and place widgets
label = tk.Label(root, text="Click below to select PDF file and process:")
label.pack(pady=10)

process_button = tk.Button(root, text="Select PDF File", command=process_pdf)
process_button.pack(pady=10)

# Run the main tkinter event loop
root.mainloop()
