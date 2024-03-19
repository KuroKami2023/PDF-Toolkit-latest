import requests
from PyPDF2 import PdfReader, PdfWriter
import os
import shutil
import pandas as pd
from tenacity import retry, stop_after_attempt, wait_fixed

@retry(stop=stop_after_attempt(3), wait=wait_fixed(1))
def convert_to_excel_with_retry(url, headers, payload, files):
    response = requests.post(url, headers=headers, data=payload, files=files)
    response.raise_for_status()
    return response

def split_pdf(input_pdf_path, output_folder):
    with open(input_pdf_path, 'rb') as file:
        pdf_reader = PdfReader(file)
        for page_number in range(len(pdf_reader.pages)):
            pdf_writer = PdfWriter()
            pdf_writer.add_page(pdf_reader.pages[page_number])
            output_page_path = os.path.join(output_folder, f'page_{page_number + 1}.pdf')
            with open(output_page_path, 'wb') as output_file:
                pdf_writer.write(output_file)
            print(f"Page {page_number + 1} extracted and saved as {output_page_path}")

def download_converted_pages(pdf_links, output_folder):
    for idx, link in enumerate(pdf_links):
        if link is not None:
            response = requests.get(link)
            with open(os.path.join(output_folder, f'converted_page_{idx+1}.xlsx'), 'wb') as output_file:
                output_file.write(response.content)
        else:
            print(f"Skipping download for link at index {idx} because it is None.")

def merge_excels(input_folder, output_excel_path):
    all_data = pd.DataFrame()
    for filename in os.listdir(input_folder):
        if filename.endswith(".xlsx"):
            excel_path = os.path.join(input_folder, filename)
            df = pd.read_excel(excel_path)
            all_data = pd.concat([all_data, df], ignore_index=True)
    all_data.to_excel(output_excel_path, index=False)

def convert_to_excel(input_folder, output_folder):
    url = "https://api.pdfrest.com/excel"
    headers = {
        'Api-Key': '39fbc901-2c3f-40ec-bee0-6096b60d75c6'  
    }

    pdf_links = []

    for filename in os.listdir(input_folder):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(input_folder, filename)
            payload = {
                'output': f'{os.path.splitext(filename)[0]}.xlsx'
            }
            files = {'file': (filename, open(pdf_path, 'rb'), 'application/pdf')}
            
            try:
                response = convert_to_excel_with_retry(url, headers, payload, files)
                if response.status_code == 200:
                    pdf_links.append(response.json().get('download_url'))
                else:
                    print(f"Failed to convert {pdf_path}. Status code: {response.status_code}")
            except Exception as e:
                print(f"An error occurred while converting {pdf_path}: {e}")

    download_converted_pages(pdf_links, output_folder)

# Paths
input_pdf_path = "test2.pdf"
temp_folder = "temp_pages"
output_excel_path = "output87654321.xlsx"

# Split PDF into individual pages
os.makedirs(temp_folder, exist_ok=True)
split_pdf(input_pdf_path, temp_folder)

# Convert each page to Excel
convert_to_excel(temp_folder, temp_folder)

# Merge Excel files into one
merge_excels(temp_folder, output_excel_path)

# Clean up temporary folder
shutil.rmtree(temp_folder)

print("Conversion completed successfully!")
