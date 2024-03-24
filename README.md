# PDF to Text and DOCX Converter

This project provides Python scripts to convert PDF files to plain text (.txt) and Microsoft Word (.docx) formats using Optical Character Recognition (OCR) via the Tesseract OCR engine.

## Features

- Convert PDF files to plain text (.txt) using OCR
- Convert the extracted text to Microsoft Word (.docx) format
- Configurable settings via a `config.yaml` file
- Check dependencies and prerequisites with the `check_dependencies.py` script

## Prerequisites

- Python 3.6 or higher
- Tesseract OCR engine
- Poppler (for pdf2image)

## Installation

1. Install Python from the official website: https://www.python.org

2. Install the required Python packages:

   ```
   pip3 install -r requirements.txt
   ```

3. Install Tesseract OCR and Poppler:

   - macOS (Homebrew):
     ```
     brew install poppler tesseract
     ```
   - Linux (Debian/Ubuntu):
     ```
     sudo apt-get install poppler-utils tesseract-ocr
     ```
   - Windows:
     Download and install the binaries from the official websites:
     - Poppler: https://blog.alivate.com.au/poppler-windows/
     - Tesseract OCR: https://github.com/UB-Mannheim/tesseract/wiki

4. Configure the `config.yaml` file with the appropriate settings (input/output folders, Tesseract path, etc.).

## Usage

1. Run the `check_dependencies.py` script to ensure all dependencies and prerequisites are met:

   ```
   python3 check_dependencies.py
   ```

   This script will check for the required Python packages, external dependencies (Poppler and Tesseract OCR), and the specified paths in the `config.yaml` file. If any issues are found, the script will provide instructions on how to resolve them.

2. Place the PDF files you want to convert in the `input_folder` specified in the `config.yaml` file.

3. Run the `pdf_to_txt.py` script to convert the PDFs to plain text:

   ```
   python3 pdf_to_txt.py
   ```

   The converted text files will be saved in the `output_folder` specified in the `config.yaml` file.

4. (Optional) Run the `txt_to_docx.py` script to convert the extracted text to Microsoft Word (.docx) format:

   ```
   python3 txt_to_docx.py
   ```

   The converted .docx files will be saved in the `docx_output_folder` specified in the `config.yaml` file.

## Configuration

The `config.yaml` file contains the following configurable settings:

- `input_folder`: Input folder containing the PDF files to be converted
- `output_folder`: Output folder for the converted TXT files
- `tesseract_path`: Path to the Tesseract OCR executable
- `docx_output_folder`: Output folder for the converted DOCX files
- `ocr_lang`: OCR language

Please make sure to update the `config.yaml` file with the appropriate paths and settings before running the scripts.