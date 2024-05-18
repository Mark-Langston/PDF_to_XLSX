import os
import PyPDF4
import pandas as pd

def get_fillable_fields(pdf_path):
    """Extracts field names from PDF files"""
    fields = []
    with open(pdf_path, 'rb') as file:
        reader = PyPDF4.PdfFileReader(file)
        if reader.isEncrypted:
            try:
                reader.decrypt('')
            except:
                """Throws errors in event of encryption, nice to know if data is useless"""
                print(f"Could not decrypt {pdf_path}. Skipping.")
                return fields
        fields_dict = reader.getFields()
        if fields_dict:
            fields = list(fields_dict.keys())
    return fields

def process_pdfs(folder_path):
    """Converts PDF fields into DataFrames, pandas library lingo for spreadsheet data"""
    data = []
    for filename in os.listdir(folder_path):
        if filename.endswith('.pdf'):
            pdf_path = os.path.join(folder_path, filename)
            fields = get_fillable_fields(pdf_path)
            base_filename = os.path.splitext(filename)[0]
            for field in fields:
                data.append([base_filename, field])
    return pd.DataFrame(data, columns=['File Name', 'Field Name'])

def save_to_excel(df, output_path):
    """Saves the DataFrame to an Excel file"""
    df.to_excel(output_path, index=False)

# Actual execution of the program happens now
folder_path = r'C:\enterDirectoryHere' # Folder path where PDFs are stored
output_path = r'C:\enterDirectoryHere' # Folder path to save excel file

df = process_pdfs(folder_path)
save_to_excel(df, output_path)

print(f"Spreadsheet created successfully and saved to {output_path}")

