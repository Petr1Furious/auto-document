import threading
import time
import multiprocessing

import docx
import easywebdav
from easywebdav import OperationFailed
from openpyxl import load_workbook
from requests.exceptions import ConnectionError
import yaml
import os
import string
from urllib.parse import unquote


def make_defaults():
    existing = []
    if os.path.exists('config.yml'):
        parsed = yaml.safe_load(open('config.yml'))
        if parsed is not None:
            existing = list(parsed)

    desc = ['Имя пользователя', 'Пароль для приложения', 'Путь к таблице', 'Название листа для студентов',
            'Название листа для руководителей', 'Путь для документов',
            'Начальный ряд для обработки в листе для студентов',
            'Конечный ряд для обработки в листе для студентов или 0, если нужно идти до конца таблицы',
            'Начальный ряд для обработки в листе для руководителей',
            'Конечный ряд для обработки в листе для руководителей или 0, если нужно идти до конца таблицы',
            'Перезаписать существующие документы']
    names = ['username', 'password', 'table-path', 'leaders-sheet-name', 'students-sheet-name', 'documents-path',
             'leaders-first-row', 'leaders-last-row', 'students-first-row', 'students-last-row',
             'overwrite-existing-documents']
    values = ['', '', '/table.xlsx', 'ВЫГРУЗКАЗаявки от руководителей', 'ВЫГРУЗКА Инициативные заявки', '/', 3, 0,
              2, 0, False]

    f = open('config.yml', 'a')

    for i in range(len(desc)):
        if names[i] not in existing:
            if desc[i] != '':
                f.write('# ' + desc[i] + '\n')
            f.write(names[i] + ': ')
            if isinstance(values[i], str):
                f.write('\'' + values[i] + '\'')
            else:
                f.write(str(values[i]))
            f.write('\n')


def load_config():
    return yaml.safe_load(open('config.yml'))


def get_person_name(s):
    words = s.split(' ')
    return words[0] + ' ' + [word[0].upper() + '.' for word in words[1:]]


def get_project_name(s):
    exceptions = ['a', 'an', 'the', 'about', 'above', 'across', 'after', 'against', 'along', 'among', 'around',
                  'at', 'before', 'behind', 'between', 'beyond', 'but', 'by', 'concerning', 'despite', 'down',
                  'during', 'except', 'following', 'for', 'from', 'in', 'including', 'into', 'like', 'near',
                  'of', 'off', 'on', 'onto', 'out', 'over', 'past', 'plus', 'since', 'throughout', 'to', 'towards',
                  'under', 'until', 'up', 'upon', 'with', 'within', 'without']
    words = s.split(' ')
    res_words = []
    for word in words:
        if word.lower() not in exceptions:
            res_words.append(string.capwords(word))
        else:
            res_words.append(word.lower())
    return ' '.join(res_words)


def join(*to_join, sep=' '):
    return sep.join([mystr(s) for s in to_join if mystr(s) != ''])


def mystr(cell_text):
    return '' if cell_text is None else str(cell_text)


def get_file_name(name, surname, patronymic, project_name):
    return 'Описание проекта_' + mystr(name) + ' ' + mystr(surname)[0] + '.' + \
           (' ' + mystr(patronymic)[0] + '.' if patronymic is not None and patronymic != '-' else '') + \
           '_' + mystr(project_name) + '.docx'


def make_document_leader(row):
    if row[3] is None or row[4] is None:
        return

    doc = docx.Document('template.docx')

    rows = doc.tables[0].rows
    rows[0].cells[1].text = mystr(row[25])
    rows[1].cells[1].text = get_project_name(mystr(row[26]))
    rows[2].cells[1].text = mystr(row[27])
    rows[3].cells[1].text = join(row[28], row[29])
    rows[5].cells[1].text = \
        '1' if mystr(row[28]) == 'Индивидуальный' or mystr(row[29]) == 'Индивидуальный' else mystr(row[30])
    rows[6].cells[1].text = mystr(row[3]) + ' ' + mystr(row[4]) + ' ' + ('' if row[5] == '-' else mystr(row[5]))
    rows[7].cells[1].text = join(row[10], row[11], row[17], row[18], row[19])
    rows[8].cells[1].text = mystr(row[3]) + ' ' + mystr(row[4]) + ' ' + ('' if row[5] == '-' else mystr(row[5]))
    rows[9].cells[1].text = mystr(row[32])
    rows[10].cells[1].text = mystr(row[33])
    rows[11].cells[1].text = mystr(row[34])
    rows[12].cells[1].text = mystr(row[35])
    rows[13].cells[1].text = mystr(row[41])
    rows[15].cells[1].text = mystr(row[31])
    rows[16].cells[1].text = 'На email ' + mystr(row[6]) + '\n\n1 сентября 2022 по 15 октября 2022'
    rows[17].cells[1].text = join(row[36], row[37], row[39], sep='\n')

    doc.save(get_file_name(row[3], row[4], row[5], row[25]))


def make_documents(leaders_sheet, func, first_row, last_row):
    values_table = [[c.value for c in r] for r in leaders_sheet.rows]

    for i in range(first_row, len(values_table) if last_row == -1 else last_row + 1):
        func(values_table[i])


def upload(webdav, local_path, remote_path):
    webdav.upload(local_path, remote_path)


def try_upload():
    try:
        config = load_config()

        # password = 'hlhijtvxlwjuyrpl'

        try:
            webdav = easywebdav.connect('webdav.yandex.ru', username=config['username'],
                                        password=config['password'], protocol='https')
            webdav.ls()
        except ConnectionError as e:
            print('Отсутствует подключение к интернету.')
            return False
        except OperationFailed as e:
            print('Неверный логин или пароль.')
            return False

        try:
            webdav.download(config['table-path'], 'table.xlsx')
        except OperationFailed as e:
            print('По заданному пути таблица не найдена.')
            return False

        workbook = load_workbook('table.xlsx')
        leaders_sheet = workbook[config['leaders-sheet-name']]
        students_sheet = workbook[config['students-sheet-name']]

        make_documents(leaders_sheet, make_document_leader,
                       config['leaders-first-row'] - 1, config['leaders-last-row'] - 1)

        documents_list = []
        try:
            documents_list = webdav.ls(config['documents-path'])
        except OperationFailed as e:
            print('Папка для документов не найдена.')
            return False

        documents_names_list = [os.path.splitext(os.path.basename(unquote(x.name)))[0] for x in documents_list if
                                os.path.splitext(os.path.basename(unquote(x.name)))[1] == '.docx']

        threads = []
        for file in [f for f in os.listdir() if os.path.isfile(f)]:
            filename, file_extension = os.path.splitext(file)
            if file_extension == '.docx' and not filename == 'template':
                if config['overwrite-existing-documents'] or documents_names_list.count(filename) == 0:
                    threads.append(multiprocessing.Process(
                        target=upload, args=(webdav, str(file), config['documents-path'] + str(file))))
                    # webdav.upload(str(file), config['documents-path'] + str(file))

        for thread in threads:
            thread.start()

        for thread in threads:
            thread.join()

        os.remove('table.xlsx')
    except OperationFailed as e:
        print('Ошибка во время загрузки: ' + str(e))
        return False
    except ConnectionError as e:
        print('Ошибка подключения: ' + str(e))
        return False
    # except Exception as e:
    #     print('Ошибка: ' + str(e))
    #     return False

    return True


if __name__ == '__main__':
    make_defaults()

    if try_upload():
        print('Документы успешно созданы и загружены.')
    else:
        print('Документы не были созданы или загружены')

    for file in [f for f in os.listdir() if os.path.isfile(f)]:
        filename, file_extension = os.path.splitext(file)
        if (file_extension == ".xlsx" or file_extension == ".docx") and filename != "template":
            os.remove(file)
