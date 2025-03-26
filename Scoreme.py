# Install the necessary libraries
pip install -q tabula-py pandas openpyxl camelot-py pdfplumber pytesseract
pip install -y ghostscript python3-tk poppler-utils

import tabula
import camelot
import pandas as pd
import pdfplumber
import pytesseract
from google.colab import files
import io
import warnings

# Suppress warnings for cleaner output
warnings.filterwarnings('ignore')

def extract_tables_with_pdfplumber(pdf_path):
    """Extract text from PDF with pdfplumber and OCR if necessary."""
    all_tables = []
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()  # Extract text using pdfplumber
                if text:
                    print(f"Text on page {page.page_number}:")
                    print(text)
                else:
                    print(f"No text found on page {page.page_number}. Trying OCR...")
                    img = page.to_image()
                    ocr_result = pytesseract.image_to_string(img.original)  # Apply OCR
                    print(f"OCR Result on page {page.page_number}: {ocr_result.strip()}\n")
                    # You might want to process this OCR result to find tables
    except Exception as e:
        print(f"Error during PDF processing with pdfplumber: {str(e)}")

    return all_tables  # Update this to return structured data if you extract tables from OCR

def extract_tables_reliably(pdf_path):
    """Extract tables using the best available method."""
    all_tables = []

    # Attempt Camelot extraction
    try:
        print("Attempting Camelot extraction...")
        camelot_tables = camelot.read_pdf(pdf_path, flavor='stream', pages='all')
        if camelot_tables:
            print(f"Found {len(camelot_tables)} tables with Camelot.")
            all_tables.extend([t.df for t in camelot_tables])
        else:
            print("No tables found by Camelot.")
    except Exception as e:
        print(f"Camelot error: {str(e)}")

    # Attempt Tabula extraction
    try:
        print("Attempting Tabula extraction...")
        tabula_tables = tabula.read_pdf(pdf_path, pages='all', multiple_tables=True)
        if tabula_tables:
            print(f"Found {len(tabula_tables)} tables with Tabula.")
            all_tables.extend(tabula_tables)
        else:
            print("No tables found by Tabula.")
    except Exception as e:
        print(f"Tabula error: {str(e)}")

    return all_tables

def clean_table(df):
    """Clean and prepare table data."""
    if df.empty:
        print("Received an empty DataFrame for cleaning.")
        return None

    df = df.dropna(how='all').dropna(axis=1, how='all')

    if len(df) > 1:
        for i in range(min(3, len(df))):
            if df.iloc[i].notna().sum() > len(df.columns) / 2:
                df.columns = df.iloc[i].astype(str)
                df = df.iloc[i + 1:]
                print(f"Header found: {df.columns.tolist()}")
                break

    return df.reset_index(drop=True) if not df.empty else None

def create_excel_file(tables, filename):
    """Create and save an Excel file."""
    if not tables:
        print("No tables to save in the Excel file.")
        return False

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for i, table in enumerate(tables):
            sheet_name = f"Table_{i + 1}"[:31]
            table.to_excel(writer, sheet_name=sheet_name, index=False)

    # Save and download file
    with open(filename, 'wb') as f:
        f.write(output.getvalue())
    files.download(filename)
    return True

# Main process
print("Please upload your PDF file:")
uploaded = files.upload()

pdf_file = next(iter(uploaded))

if pdf_file.lower().endswith('.pdf'):
    # Extract tables
    raw_tables = extract_tables_reliably(pdf_file)
    
    if not raw_tables:
        print("No tables found in the PDF.")
        raw_tables = extract_tables_with_pdfplumber(pdf_file)  # Trying pdfplumber as well
    
    if not raw_tables:
        print("Still no tables found after using pdfplumber.")
    else:
        # Log raw table outputs for debugging
        for idx, table in enumerate(raw_tables):
            print(f"Raw Table {idx + 1}:")
            print(table)
            print("\n")
        
        # Clean tables
        cleaned_tables = [clean_table(t) for t in raw_tables]
        cleaned_tables = [t for t in cleaned_tables if t is not None]

        if not cleaned_tables:
            print("No valid tables found after cleaning.")
        else:
            # Create and download Excel file
            output_filename = 'extracted_tables.xlsx'
            if create_excel_file(cleaned_tables, output_filename):
                print(f"Success! Downloaded {output_filename} with {len(cleaned_tables)} tables.")
            else:
                print("Failed to create Excel file.")
else:
    print("Please upload a valid PDF file.")