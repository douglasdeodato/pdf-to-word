import os
import PyPDF2
from docx import Document

def pdf_to_word(pdf_file, word_file):
    # Check if the PDF file exists
    if not os.path.exists(pdf_file):
        print("PDF file not found.")
        return
    
    # Initialize the PDF reader
    pdf_reader = PyPDF2.PdfFileReader(pdf_file)

    # Initialize the Word document
    doc = Document()

    # Iterate through the PDF pages and extract text
    for page_num in range(pdf_reader.numPages):
        page = pdf_reader.getPage(page_num)
        text = page.extractText()
        doc.add_paragraph(text)

    # Save the Word document
    doc.save(word_file)
    print(f"Conversion complete. Word file saved as {word_file}")

if __name__ == "__main__":
    # Replace 'input_file.pdf' with the name of your PDF file
    input_pdf = "input_file.pdf"

    # Replace 'output_file.docx' with the desired name of the output Word file
    output_word = "output_file.docx"

    pdf_to_word(input_pdf, output_word)
