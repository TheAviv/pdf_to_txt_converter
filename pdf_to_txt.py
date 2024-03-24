import os
import glob
import logging
from pdf2image import convert_from_path
import pytesseract
import yaml
from concurrent.futures import ProcessPoolExecutor, as_completed
import cv2
import numpy as np
import time

def setup_logging(logging_level):
    logging.basicConfig(level=logging_level, format='%(asctime)s - %(levelname)s - %(message)s')

def preprocess_image(image):
    logging.info("Preprocessing image...")
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]
    denoised = cv2.fastNlMeansDenoisingColored(image, None, 10, 10, 7, 21)
    logging.info("Image preprocessed.")
    return denoised

def detect_orientation(image, oem, psm):
    logging.info("Detecting image orientation...")
    orientations = [0, 90, 180, 270]
    max_chars = 0
    best_orientation = 0

    for orientation in orientations:
        rotated = cv2.rotate(image, orientation)
        text = pytesseract.image_to_string(rotated, config=f'--oem {oem} --psm {psm}')
        char_count = len(text)
        if char_count > max_chars:
            max_chars = char_count
            best_orientation = orientation
    logging.info(f"Detected orientation: {best_orientation} degrees.")
    return best_orientation

def convert_pdf_to_text(pdf_file, tesseract_path, output_folder, ocr_lang, ocr_engine_mode, page_segmentation_mode, output_prefix, output_suffix):
    pytesseract.pytesseract.tesseract_cmd = tesseract_path

    pdf_file_name = os.path.basename(pdf_file).replace('.pdf', '')
    logging.info(f"Starting conversion for '{pdf_file_name}'")
    try:
        pages = convert_from_path(pdf_file, 300)
        text = ""
        
        for page_num, page in enumerate(pages, start=1):
            logging.info(f"Processing page {page_num}/{len(pages)} of '{pdf_file_name}'...")
            page_image = cv2.cvtColor(np.array(page), cv2.COLOR_RGB2BGR)
            preprocessed_image = preprocess_image(page_image)
            orientation = detect_orientation(preprocessed_image, ocr_engine_mode, page_segmentation_mode)
            rotated_image = cv2.rotate(preprocessed_image, orientation)
            logging.info("Performing OCR...")
            page_text = pytesseract.image_to_string(rotated_image, lang=ocr_lang, config=f'--oem {ocr_engine_mode} --psm {page_segmentation_mode}')
            text += page_text + "\n"
            logging.info(f"Page {page_num}/{len(pages)} processed.")

        output_file = os.path.join(output_folder, f'{output_prefix}{pdf_file_name}{output_suffix}.txt')
        logging.info(f"Writing output for '{pdf_file_name}'...")
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(text)
        logging.info(f"Text extracted and saved for '{pdf_file_name}'.")

    except Exception as e:
        logging.error(f"Error processing '{pdf_file_name}': {e}")

def main():
    with open('config.yaml', 'r') as f:
        config = yaml.safe_load(f)

    setup_logging(config['logging_level'])

    input_folder = config['input_folder']
    output_folder = config['output_folder']
    tesseract_path = config['tesseract_path']
    ocr_lang = config['ocr_lang']
    ocr_engine_mode = config['ocr_engine_mode']
    page_segmentation_mode = config['page_segmentation_mode']
    output_prefix = config['output_prefix']
    output_suffix = config['output_suffix']

    if not os.path.exists(input_folder) or not os.listdir(input_folder):
        logging.error("Input folder does not exist or is empty.")
        return

    os.makedirs(output_folder, exist_ok=True)
    pdf_files = glob.glob(os.path.join(input_folder, '*.pdf'))
    logging.info(f"Found {len(pdf_files)} PDF files in '{input_folder}'. Processing them now...")

    for pdf_file in pdf_files:
        convert_pdf_to_text(pdf_file, tesseract_path, output_folder, ocr_lang, ocr_engine_mode, page_segmentation_mode, output_prefix, output_suffix)

    logging.info("All PDFs processed.")

if __name__ == "__main__":
    main()
