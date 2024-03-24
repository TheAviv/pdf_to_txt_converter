import os
import glob
import logging
from docx import Document
import yaml
from concurrent.futures import ProcessPoolExecutor
from tqdm import tqdm

def setup_logging(logging_level):
    logging.basicConfig(level=logging_level, format='%(asctime)s - %(levelname)s - %(message)s')

def convert_text_to_docx(text_file, docx_output_folder, output_prefix, output_suffix):
    try:
        # Extract the file name without extension
        file_name = os.path.basename(text_file).replace('.txt', '')

        # Initialize a new Document for .docx output
        doc = Document()

        # Read the text from the input file
        with open(text_file, 'r', encoding='utf-8') as f:
            text = f.read()

        # Add the text to the .docx document
        doc.add_paragraph(text)

        # Define the output .docx file name based on the input file name
        docx_output_file = os.path.join(docx_output_folder, f'{output_prefix}{file_name}{output_suffix}.docx')

        # Save the .docx file
        doc.save(docx_output_file)
        logging.info(f"Converted '{text_file}' to '{docx_output_file}'")
    except Exception as e:
        logging.error(f"Error converting text file '{text_file}': {str(e)}")

def main():
    with open('config.yaml', 'r') as f:
        config = yaml.safe_load(f)

    text_input_folder = config['output_folder']  # Folder containing the text files
    docx_output_folder = config['docx_output_folder']  # Folder to save the .docx files
    output_prefix = config['output_prefix']
    output_suffix = config['output_suffix']
    logging_level = config['logging_level']

    setup_logging(logging_level)

    # Check if the input text folder exists
    if not os.path.exists(text_input_folder):
        logging.error("The 'Converted_TXT' folder does not exist.")
        logging.error("Please run the 'pdf_to_txt.py' script first to generate the text files.")
        exit()

    # Find all text files in the input folder
    text_files = glob.glob(os.path.join(text_input_folder, '*.txt'))

    # Check if there are any text files in the input folder
    if not text_files:
        logging.error("No text files found in the 'Converted_TXT' folder.")
        logging.error("Please run the 'pdf_to_txt.py' script first to generate the text files.")
        exit()

    # Ensure the output folder exists
    os.makedirs(docx_output_folder, exist_ok=True)

    total_files = len(text_files)
    logging.info(f"Found {total_files} text files in '{text_input_folder}'")

    # Use ProcessPoolExecutor for parallel processing
    with ProcessPoolExecutor() as executor:
        futures = [executor.submit(convert_text_to_docx, text_file, docx_output_folder, output_prefix, output_suffix) for text_file in text_files]
        for _ in tqdm(executor.map(lambda x: x, futures), total=total_files, desc="Converting TXT to DOCX", unit="file"):
            pass

    logging.info("All text files converted to .docx.")

if __name__ == "__main__":
    main()