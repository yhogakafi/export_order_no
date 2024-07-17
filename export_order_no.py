import pdfplumber
import pandas as pd
import re

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

# Main script
pdf_path = 'file.pdf'
keyword = 'No. Pesanan'
output_path = 'output.xlsx'

# Extract information
extracted_info = extract_info_from_pdf(pdf_path, keyword)

# Save to Excel
save_to_excel(extracted_info, output_path)

print(f"Extracted information saved to {output_path}")
