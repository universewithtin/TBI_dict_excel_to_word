from tkinter import Tk, filedialog
from docx import Document
import re

def compare_pages(line):
    pattern = r'\.\s+(\d+)\s+б\.\)'
    match = re.search(pattern, line)
    prev_page = 0
    if match:
        page_number = int(match.group(1))
        if prev_page > page_number:
            print(f"ERR:: {page_number}")
        prev_page = page_number

def parse_lines(file_path):
    doc = Document(file_path)
    full_text = ''
    for para in doc.paragraphs:
        full_text += para.text + '\n'

    lines = full_text.split('\n')

    for line in lines:
        compare_pages(line.strip())

def open_file_dialog():
    root = Tk()
    root.withdraw()
    
    file_path = filedialog.askopenfilename(title="WORD DOC", filetypes=(("Word files", "*.docx"), ("All files", "*.*")))
    
    if file_path:
        parse_lines(file_path)


open_file_dialog()