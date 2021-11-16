
import datetime
import docx
from glob import glob
from pathlib import Path

paths = glob(r'docs/*.docx', recursive=True)


def get_metadata():
    for fl in paths:
        doc = docx.Document(fl)
        properties = doc.core_properties
        print(fl)
        print('Автор документа:', properties.author)
        print('Автор последней правки:', properties.last_modified_by)
        print('Дата создания документа:', properties.created)
        print('Дата последней правки:', properties.modified)
        print('Дата последней печати:', properties.last_printed)
        print('Количество сохранений:', properties.revision)


def get_log(string):
    with open('log.txt', 'a', encoding='utf-8') as f:
        res = f'{datetime.datetime.now()}, {string}\n'
        f.write(res)


def rename_files():
    files = glob(f'{paths}*', recursive=True)
    for i, f in enumerate(files):
        p = Path(f)
        get_log(f'file: {f} {i}')
        p.rename(Path(p.parent, f"{i}{p.suffix}"))

