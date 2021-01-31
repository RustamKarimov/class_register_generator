from pathlib import Path

BASE_DIR = Path(__file__).parent.parent

DATA_FILES_LOCATION = BASE_DIR / 'data'

EXCEL_TEMPLATE = DATA_FILES_LOCATION / 'register_template.xlsx'
CLASS_LIST = DATA_FILES_LOCATION / 'class_list.xlsx'
OUTPUT_FOLDER = BASE_DIR / 'output'
OUTPUT_FOLDER_FOR_EXCEL_FILES = OUTPUT_FOLDER / 'excel'
OUTPUT_FOLDER_FOR_PDF_FILES = OUTPUT_FOLDER / 'pdf'
