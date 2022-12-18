import os

from docx2pdf import convert
from pdf2docx import Converter

import file_handler


def pdf_to_docx_process(file_name: str) -> None:
    cv = Converter(rf'{os.getcwd()}\{file_name}')
    cv.convert(rf'{os.getcwd()}\{file_name.split(".")[0]}.docx')
    print(f'Файл "{file_name}" успешно преобразован!')


def convert_pdf_to_docx() -> None:
    pdf_files = file_handler.find_file_with_ext('pdf')

    print('Список файлов с раширением pdf в данном каталоге: \n')

    for i in range(0, len(pdf_files)):
        print(f'{i + 1}. {pdf_files[i]}')

    chosen_file_id = int(
        input('Введите номер файла для преобразования (чтобы преобразовать все файлы из данного каталога введите 0: '))

    if chosen_file_id not in range(0, len(pdf_files) + 1):
        print('Неверный номер')
        return

    if chosen_file_id == 0:
        for j in range(0, len(pdf_files)):
            pdf_to_docx_process(pdf_files[j])
    else:
        pdf_to_docx_process(pdf_files[chosen_file_id-1])


def docx_to_pdf_process(file_name: str) -> None:
    convert(rf'{os.getcwd()}\{file_name}', rf'{os.getcwd()}\{file_name.split(".")[0]}.pdf')
    print(f'Файл "{file_name}" успешно преобразован!')


def convert_docx_to_pdf() -> None:
    docx_files = file_handler.find_file_with_ext('docx')

    print('Список файлов с раширением docx в данном каталоге: \n')

    for i in range(0, len(docx_files)):
        print(f'{i + 1}. {docx_files[i]}')

    chosen_file_id = int(
        input('Введите номер файла для преобразования (чтобы преобразовать все файлы из данного каталога введите 0: '))

    if chosen_file_id not in range(0, len(docx_files) + 1):
        print('Неверный номер')
        return

    if chosen_file_id == 0:
        for j in range(0, len(docx_files)):
            docx_to_pdf_process(docx_files[j])
    else:
        docx_to_pdf_process(docx_files[chosen_file_id-1])
