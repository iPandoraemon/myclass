import os

from docx import Document
from docxcompose.composer import Composer

road_docx = r'C:\output_fafa\cs\cs_docx'
road_all = r'C:\output_fafa\cs\cs_docx'
original_docx_path = road_docx
new_docx_path = f'{road_all}/activity_name.docx'


all_file_path = []
for file_name in os.listdir(original_docx_path):
    all_file_path.append(f'{original_docx_path}/{file_name}')

first_document = Document(all_file_path[0])
first_document.add_page_break()
middle_new_docx = Composer(first_document)

for index, word in enumerate(all_file_path[1:]):
    word_document = Document(word)
    if index != len(all_file_path) - 2:
        word_document.add_page_break()
        middle_new_docx.append(word_document)

middle_new_docx.save(new_docx_path)