import subprocess
import os
from importlib.metadata import version, PackageNotFoundError
import yaml
from packaging.version import parse

def check_python_package(package_name, required_version):
    try:
        installed_version = version(package_name)
        if parse(installed_version) >= parse(required_version):
            print(f"[✓] Python package '{package_name}' (version {installed_version}) is installed.")
            return True
        else:
            print(f"[!] Python package '{package_name}' is installed, but the version ({installed_version}) is lower than the required version ({required_version}).")
            return False
    except PackageNotFoundError:
        print(f"[!] Python package '{package_name}' is NOT installed. Please install it using pip.")
        return False

def check_external_dependency(command, success_keyword, help_option=False):
    command_list = [command, '--help'] if help_option else [command, '--version']
    try:
        result = subprocess.run(command_list, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, env=os.environ)
        if success_keyword in result.stdout or result.returncode == 0:
            print(f"[✓] External dependency '{command}' is installed.")
            return True
        else:
            print(f"[!] Error or unexpected output running '{command}'.")
            return False
    except FileNotFoundError:
        print(f"[!] External dependency '{command}' is NOT installed. Please install it using your system's package manager.")
        return False

def check_tesseract_path(tesseract_path):
    if os.path.isfile(tesseract_path):
        print(f"[✓] Tesseract OCR executable found at '{tesseract_path}'.")
        return True
    else:
        print(f"[!] Tesseract OCR executable not found at '{tesseract_path}'. Please check the installation or update the path.")
        return False

def check_input_folder(input_folder):
    if os.path.isdir(input_folder) and os.listdir(input_folder):
        print(f"[✓] Input folder '{input_folder}' exists and contains files.")
        return True
    else:
        print(f"[!] Input folder '{input_folder}' does not exist or is empty.")
        return False

def provide_installation_instructions():
    instructions = """
Installation Instructions:
- Python packages can be installed using pip. Run the following command:
  pip install -r requirements.txt

- Poppler (for pdf2image) and Tesseract OCR (for pytesseract) installation varies by OS:
  - macOS (Homebrew):
    brew install poppler tesseract
  - Linux (Debian/Ubuntu):
    sudo apt-get install poppler-utils tesseract-ocr
  - Windows:
    Download and install the binaries from the official websites:
    - Poppler: https://blog.alivate.com.au/poppler-windows/
    - Tesseract OCR: https://github.com/UB-Mannheim/tesseract/wiki
"""
    return instructions

def validate_config(config):
    required_keys = ['input_folder', 'output_folder', 'tesseract_path', 'docx_output_folder', 'ocr_lang']
    for key in required_keys:
        if key not in config:
            print(f"[!] Missing configuration key: {key}")
            return False
    return True

def main():
    with open('config.yaml', 'r') as f:
        config = yaml.safe_load(f)

    if not validate_config(config):
        print("[!] Invalid configuration. Please check the config.yaml file.")
        exit()

    input_folder = config['input_folder']
    tesseract_path = config['tesseract_path']

    results = []

    print("\nChecking Python packages:")
    with open('requirements.txt', 'r') as f:
        for line in f:
            package_info = line.strip().split('>=')
            package_name = package_info[0]
            required_version = package_info[1] if len(package_info) > 1 else None
            success = check_python_package(package_name, required_version)
            results.append(success)

    print("\nChecking external dependencies:")
    external_dependencies = [
        ('pdfinfo', 'Usage', True),
        ('tesseract', 'tesseract', False)
    ]
    for command, keyword, help_option in external_dependencies:
        success = check_external_dependency(command, keyword, help_option)
        results.append(success)

    print("\nChecking Tesseract OCR executable path:")
    success = check_tesseract_path(tesseract_path)
    results.append(success)

    print("\nChecking input folder:")
    success = check_input_folder(input_folder)
    results.append(success)

    print()
    if all(results):
        print("All dependencies and prerequisites are met. You're ready to run the conversion scripts!")
    else:
        print("Some dependencies or prerequisites are missing or not correctly configured.")
        print("Please refer to the installation instructions below:")
        print(provide_installation_instructions())

    print()

if __name__ == "__main__":
    main()