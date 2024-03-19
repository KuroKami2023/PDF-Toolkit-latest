from PyPDF2 import PdfReader, PdfWriter

def split_pdf(input_pdf_path, output_folder):
    with open(input_pdf_path, 'rb') as file:
        pdf_reader = PdfReader(file)
        for page_number in range(len(pdf_reader.pages)):
            pdf_writer = PdfWriter()
            pdf_writer.add_page(pdf_reader.pages[page_number])
            output_page_path = f"{output_folder}/page_{page_number + 1}.pdf"
            with open(output_page_path, 'wb') as output_file:
                pdf_writer.write(output_file)
            print(f"Page {page_number + 1} extracted and saved as {output_page_path}")

input_pdf_path = "test2.pdf"
output_folder = "input_folder"
split_pdf(input_pdf_path, output_folder)
