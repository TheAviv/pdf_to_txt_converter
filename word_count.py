import os
import csv
import yaml
from docx import Document  # Importing the Document class for handling DOCX files

# Function to count words in a single file
def count_words_in_file(filepath):
    if filepath.endswith('.txt'):
        with open(filepath, 'r', encoding='utf-8') as file:
            contents = file.read()
            word_count = len(contents.split())
    elif filepath.endswith('.docx'):
        doc = Document(filepath)
        word_count = sum(len(para.text.split()) for para in doc.paragraphs)
    else:
        word_count = 0
    return word_count

# Function to process all TXT and DOCX files in a directory and save results to CSV
def process_directory(config_path):
    # Load configuration
    with open(config_path, 'r') as config_file:
        config = yaml.safe_load(config_file)
    
    input_folder = config['input_folder_for_word_count']
    output_path = config['output_path_for_word_count']
    
    # Ensure the output directory exists
    output_dir = os.path.dirname(output_path)
    if not os.path.exists(output_dir) and output_dir != '':
        os.makedirs(output_dir)
    
    results = []
    # Iterate over all files in the given directory
    for filename in os.listdir(input_folder):
        if filename.endswith(".txt") or filename.endswith(".docx"):
            filepath = os.path.join(input_folder, filename)
            word_count = count_words_in_file(filepath)
            results.append((filename, word_count))
    
    # Sort results by filenames in alphabetical order before writing to CSV
    sorted_results = sorted(results, key=lambda x: x[0])
    
    # Export sorted results to CSV
    with open(output_path, 'w', newline='', encoding='utf-8') as csv_file:
        writer = csv.writer(csv_file)
        writer.writerow(['Filename', 'Word Count'])  # Writing header
        writer.writerows(sorted_results)
    print(f"Word counts have been saved to {output_path}")

# Path to the configuration file
config_file_path = 'config.yaml'  # Adjust the path as necessary

# Running the function
process_directory(config_file_path)
