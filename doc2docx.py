
from glob import glob
import re
import os
import win32com.client as win32
from win32com.client import constants
from funktions import get_log
import shutil
import pkg_resources
from path_to_project import get_path_to_project

# paths to .doc files
paths = glob(r'docs/*.doc', recursive=True)


def save_as_docx(path_file):
    try:
        # Opening MS Word
        word = win32.gencache.EnsureDispatch('Word.Application')
    except AttributeError:
        # При ошибке наиболее вероятна необходимость очистки кэша приложения
        f_loc = glob(r'C:\Users\*\AppData\Local\Temp\gen_py', recursive=True)
        # for f in f_loc:
            # os.unlink(f)
            # os.rmdir(f)
        shutil.rmtree(f_loc)
        word = win32.gencache.EnsureDispatch('Word.Application')
    file_name = path_file.split('\\')[1]
    doc = word.Documents.Open(f"{get_path_to_project()}docs/{file_name}")
    doc.Activate()

    # Rename path with .docx
    new_file_abs = os.path.abspath(path_file)
    new_file_abs = re.sub(r'\.\w+$', '.docx', new_file_abs)

    # Save and Close
    word.ActiveDocument.SaveAs(
        new_file_abs, FileFormat=constants.wdFormatXMLDocument
    )
    doc.Close(False)
    print(f'Файл конвертирован: {os.path.dirname(new_file_abs)}')
    get_log(f'Файл конвертирован: {os.path.dirname(new_file_abs)}')


# Нужно найти полный адрес
for path in paths:
    save_as_docx(path)


