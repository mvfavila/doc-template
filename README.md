This script is a command-line tool that takes a .docx template and a .csv file and replaces placeholders in the .docx with values from the .csv. It also generates a PDF for each .docx output. The script is designed to be used with the Brazilian court system's process documentation.

# Usage

1. Place your .docx template and .csv file in the same folder as this script.
2. Make sure the .csv file has a column named "Numero_do_Processo".
3. Run the script with `python3 main.py`

# How it works

1. The script finds the .docx template and .csv file and validates them.
2. It reads the .csv file and processes each row.
3. For each row, it creates a new .docx file by replacing placeholders in the template with the row values.
4. It generates a corresponding PDF file from the new .docx file.

# Placeholders

Placeholders are any text within double curly braces, e.g. `{{placeholder}}`. The script will replace these placeholders with the values from the .csv file.

# How to run application locally

In the terminal, type:

> python3 -m venv venv
> source venv/bin/activate
> python3 main.py

# How to build release version

> pyinstaller -F main.py
or
> python -m PyInstaller -F main.py
