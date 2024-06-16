import PyPDF2
import pandas as pd
import re
import os

def split_pdf_file(input_pdf_path, output_folder):
    try:
        with open(input_pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            
            for page_num in range(len(reader.pages)):
                writer = PyPDF2.PdfWriter()
                writer.add_page(reader.pages[page_num])
                
                output_pdf_path = f"{output_folder}/Page_{page_num + 1}.pdf"
                
                with open(output_pdf_path, 'wb') as output_file:
                    writer.write(output_file)
    except FileNotFoundError:
        print(f"File '{input_pdf_path}' not found.")
    except PermissionError:
        print(f"Permission denied to access file '{input_pdf_path}'.")
    except PyPDF2.PdfReadError:
        print(f"Error reading file '{input_pdf_path}' as a PDF.")
    except OSError:
        print(f"Error writing to file '{output_pdf_path}'.")


def extract_from_excel(excel_file):
    try:
        df = pd.read_excel(excel_file, engine='openpyxl')
        return df.iloc[:, 0]
    except FileNotFoundError:
        print(f"File '{excel_file}' not found.")
    except ValueError:
        print(f"Error reading file '{excel_file}' as an Excel file.")


def validate_filename(filename):
    return re.sub(r'[<>:"/\\|?*]', '_', filename)


def rename_pdfs(pdf_folder, names):
    for i, name in enumerate(names):
        pdf_path = os.path.join(pdf_folder, f"Page_{i + 1}.pdf")
        new_name = validate_filename(name)
        output_pdf_path = os.path.join(pdf_folder, f"{new_name}.pdf")
        try:
            os.rename(pdf_path, output_pdf_path)
        except FileNotFoundError:
            print(f"File '{pdf_path}' not found.")
        except PermissionError:
            print(f"Permission denied to rename file '{pdf_path}'.")


# Example paths
original_pdf = "/path/to/original/pdf/original.pdf"
export_folder = "/path/to/export/folder"
excel_file = "/path/to/excel/file/names.xlsx"

# Splitting the original PDF
split_pdf_file(original_pdf, export_folder)

# Extracting names from Excel file
names = extract_from_excel(excel_file)

# Renaming the PDF files
rename_pdfs(export_folder, names)