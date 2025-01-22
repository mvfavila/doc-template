import os
import sys
import csv
import re
from docx import Document
from docx2pdf import convert
from pathlib import Path

# Get the directory of the executable
EXECUTABLE_DIR = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).parent

# Function to find a single file with a specific extension
def find_single_file(extension):
    files = [f for f in EXECUTABLE_DIR.iterdir() if f.suffix == extension]
    if len(files) == 0:
        print(f"Erro: Nenhum arquivo '{extension}' encontrado no diretório {EXECUTABLE_DIR}.")
        sys.exit(1)
    if len(files) > 1:
        print(f"Erro: Mais de um arquivo '{extension}' encontrado no diretório {EXECUTABLE_DIR}. Certifique-se de que haja apenas um.")
        sys.exit(1)
    return files[0]

# Function to validate the first column in the CSV file and check for empty cells
def validate_csv(csv_file):
    csv_path = EXECUTABLE_DIR / csv_file
    with open(csv_path, newline='', encoding='utf-8') as csvfile:
        reader = csv.DictReader(csvfile)
        if "NUMERO_DO_PROCESSO" not in reader.fieldnames:
            print("Erro: Uma das colunas do arquivo CSV deve ser nomeada como 'NUMERO_DO_PROCESSO'.")
            sys.exit(1)
        
        rows = list(reader)
        for row_num, row in enumerate(rows, start=1):
            for column_name, value in row.items():
                if value is None or value.strip() == "":
                    print(f"Erro: A célula na linha {row_num}, coluna '{column_name}' está vazia.")
                    sys.exit(1)
        return rows

# Function to sanitize file names
def sanitize_filename(filename):
    return re.sub(r'[^\w\-_\.]', '_', filename)

# Function to generate PDF from DOCX
def generate_pdf(docx_file, pdf_file):
    try:
        docx_path = EXECUTABLE_DIR / docx_file
        convert(str(docx_path))  # In-place conversion
        pdf_path = docx_path.with_suffix(".pdf").resolve()
        if pdf_path.exists():
            print(f"PDF generated successfully: {pdf_path}")
        else:
            print(f"PDF was not generated for {docx_file}")
    except Exception as e:
        print(f"Error during PDF generation for {docx_file}: {e}")

# Main logic
def main():
    print("Como usar")
    print("1. Ao iniciar, deve existir exatamente um arquivo .docx e um arquivo .csv no mesmo diretório desse executável.")
    print("2. O arquivo .docx deve possuir um placeholder para cada coluna do arquivo .csv. ex.: {{NUMERO_DO_PROCESSO}}")
    print("3. Certifique-se de que existe uma coluna chamada 'NUMERO_DO_PROCESSO' no arquivo .csv.")
    print("4. Os nomes das colunas do arquivo .csv devem ser exatamente as mesmas dos placeholders do arquivo .docx.")
    print()

    print("Processando...")
    # Find the .docx template file
    template_file_path = find_single_file('.docx')
    print(f"Arquivo de template encontrado: {template_file_path.name}")

    # Find the .csv data file
    csv_file_path = find_single_file('.csv')
    print(f"Arquivo CSV encontrado: {csv_file_path.name}")

    # Validate and read the CSV file
    rows = validate_csv(csv_file_path)

    # Process each row in the CSV file
    for row in rows:
        process_number = sanitize_filename(row["NUMERO_DO_PROCESSO"])
        docx_output = EXECUTABLE_DIR / f"{process_number}_{template_file_path.name}"
        pdf_output = EXECUTABLE_DIR / f"{process_number}_{template_file_path.stem}.pdf"

        # Create a copy of the template
        doc = Document(template_file_path)

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