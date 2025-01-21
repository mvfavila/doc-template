import os
import sys
import csv
import re
import time
from docx import Document
from docx2pdf import convert
from pathlib import Path

# Function to find a single file with a specific extension
def find_single_file(extension):
    files = [f for f in os.listdir('.') if f.endswith(extension)]
    if len(files) == 0:
        print(f"Erro: Nenhum arquivo '{extension}' encontrado no diretório.")
        sys.exit(1)
    if len(files) > 1:
        print(f"Erro: Mais de um arquivo '{extension}' encontrado no diretório. Certifique-se de que haja apenas um.")
        sys.exit(1)
    return files[0]

# Function to validate the first column in the CSV file
def validate_csv(csv_file):
    with open(csv_file, newline='', encoding='utf-8') as csvfile:
        reader = csv.DictReader(csvfile)
        if "Numero_do_Processo" not in reader.fieldnames:
            print("Erro: Uma das colunas do arquivo CSV deve ser nomeada como 'Numero_do_Processo'.")
            sys.exit(1)
        return list(reader)

# Function to sanitize file names
def sanitize_filename(filename):
    return re.sub(r'[^\w\-_\.]', '_', filename)

# Function to generate PDF from DOCX
def generate_pdf(docx_file, pdf_file):
    try:
        convert(docx_file)  # In-place conversion
        pdf_file = Path(docx_file).with_suffix(".pdf").resolve()
        if pdf_file.exists():
            print(f"PDF generated successfully: {pdf_file}")
        else:
            print(f"PDF was not generated for {docx_file}")
    except Exception as e:
        print(f"Error during PDF generation for {docx_file}: {e}")

# Main logic
def main():
    # Find the .docx template file
    template_file = find_single_file('.docx')
    print(f"Arquivo de template encontrado: {template_file}")

    # Find the .csv data file
    csv_file = find_single_file('.csv')
    print(f"Arquivo CSV encontrado: {csv_file}")

    # Validate and read the CSV file
    rows = validate_csv(csv_file)

    # Process each row in the CSV file
    for row in rows:
        process_number = sanitize_filename(row["Numero_do_Processo"])
        docx_output = f"{process_number}_{template_file}"
        pdf_output = f"{process_number}_{os.path.splitext(template_file)[0]}.pdf"

        # Create a copy of the template
        doc = Document(template_file)

        # Replace placeholders with values from the row
        for paragraph in doc.paragraphs:
            for placeholder, value in row.items():
                if f"{{{{{placeholder}}}}}" in paragraph.text:
                    paragraph.text = paragraph.text.replace(f"{{{{{placeholder}}}}}", value)

        # Save the filled .docx file
        doc.save(docx_output)

        # Generate the corresponding PDF file
        generate_pdf(docx_output, pdf_output)
        print(f"Arquivos gerados: {docx_output} e {pdf_output}")

if __name__ == "__main__":
    main()