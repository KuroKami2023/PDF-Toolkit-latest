from PyPDF2 import PdfMerger
import os

def merge_pdfs(input_folder, output_pdf_path):
    merger = PdfMerger()
    
    # Iterate through all PDF files in the input folder
    for filename in os.listdir(input_folder):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(input_folder, filename)
            merger.append(pdf_path)
    
    # Write the merged PDF to the output path
    with open(output_pdf_path, 'wb') as output_file:
        merger.write(output_file)

# Example usage:
input_folder = "input_folder"
output_pdf_path = "merged_pdf.pdf"
merge_pdfs(input_folder, output_pdf_path)
