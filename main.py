from openpyxl import Workbook
from datetime import datetime
import os
import chardet
import codecs
import re

FOLDER_PATH = ''  # путь на директорию
LINKS_NOT_FORMAT = []  # файлы не utf-8
LINKS_FORMAT = []  # файлы utf-8
DATA = {}  # словарь, где k наименование товара, v список qr-кодов
REPEAT_CODES = []  # повторяющиеся в DATA qr-коды
FILE_NAME = ''  # имя файла: арт, наименование товара
TOTAL_QR_CODES = 0  # общее количество собранных qr-кодов в директории
DATE_OF_PROCESSING = None  # дата обработки данных


def file_search(path: str) -> list:  # сбор txt файлов в папке (без проверки формата)
    directory = os.fsencode(path)
    txt_files = []
    for file in os.listdir(directory):
        filename = os.fsdecode(file)
        if filename.endswith(".txt"):
            txt_files.append(os.path.join(path, filename))
    return txt_files


def change_encoding_to_utf8(txt_file):
    global LINKS_FORMAT
    with codecs.open(txt_file, 'r', 'cp1251') as file:  # Читаем содержимое файла в исходной кодировке CP1251
        content = file.read()

    with codecs.open(txt_file, 'w', 'utf-8') as file:  # Конвертируем содержимое в UTF-8 и записываем обратно в файл
        file.write(content)

    print(f'Файл "{os.path.basename(txt_file)}" был перекодирован в UTF-8.')
    LINKS_FORMAT.append(txt_file)


def utf_search(path: str) -> None:  # проверка формата отобранных txt файлов
    global LINKS_NOT_FORMAT
    global LINKS_FORMAT
    for txt_file in file_search(path):
        with open(txt_file, 'rb') as file:
            result = chardet.detect(file.read())
            encoding = result['encoding']
            if encoding.lower() in ('utf-8', 'utf-8-sig'):
                LINKS_FORMAT.append(txt_file)
            elif encoding.lower() in ('cp1251', 'windows-1251'):
                report = f' Изменение формата: Файл "{os.path.basename(txt_file)}" имел кодировку {encoding}.'
                LINKS_NOT_FORMAT.append(report)
                print(report)
                change_encoding_to_utf8(txt_file)
            else:
                change_encoding_to_utf8(txt_file)
                report = f'Ошибка: Файл "{os.path.basename(txt_file)}" имеет кодировку {encoding}. Обработка отклонена.'
                LINKS_NOT_FORMAT.append(report)
                print(report)


def input_path():
    global FOLDER_PATH
    print('Укажите путь к папке для обработки:')
    FOLDER_PATH = input()
    print('Запуск обработки.\n')
    utf_search(FOLDER_PATH)


def add_product(item, qr) -> None:  # Заполнение словаря, где k - арт., название; v - список qr
    global REPEAT_CODES
    global DATA
    value_exists = any(qr in values for values in DATA.values())  # Проверка наличия qr в словаре
    if value_exists:
        print(f'Код {qr} повторяется.')
        REPEAT_CODES.append((item, qr))
        qr = 'ОШИБКА! ' + qr + ' - повтор кода'
    if item in DATA:
        DATA[item].append(qr)
    else:
        DATA[item] = [qr]


def file_processing():  # Обработка файла txt
    global FILE_NAME
    for link in LINKS_FORMAT:
        with open(link, 'r', encoding='utf-8') as file:

            character, loc_char = '', ''  # характеристика (арт., наименование, цвет, размер) и накопитель
            lines_to_add = 2  # сколько следующих строк добавить после удовлетворения первого условия
            codes, total_codes = [], 0

            for line in file.readlines():
                line = line.strip()

                if line.isdigit() and len(line) == 5:  # первое условие — артикул
                    loc_char = line
                    lines_to_add = 2  # сбрасываем счётчик при добавлении строки
                elif lines_to_add > 0 and line.strip():  # сверяем, сколько строк добавить и что она не пустая
                    loc_char = ', '.join([loc_char, line])
                    lines_to_add -= 1

                    if lines_to_add == 0:  # если накопитель заполнен, то счетчик = 0
                        character = loc_char

                # обработка кода
                if '(01)04' in line and '(21)' in line and len(line) == 35:
                    total_codes += 1
                    line = line.replace('(01)', '01').replace('(21)', '21')
                    add_product(character, line)
    FILE_NAME = ", ".join(character.split(", ")[:2])


def replace_special_chars() -> None:
    global FILE_NAME
    FILE_NAME = re.sub(r'[:\\?#]', '_', FILE_NAME)


def record_report():  # запись отчета в txt
    global TOTAL_QR_CODES
    global DATE_OF_PROCESSING
    TOTAL_QR_CODES = sum(len(value) for value in DATA.values())  # Общее количество собранных qr
    DATE_OF_PROCESSING = (datetime.now().strftime('%d.%m.%y / %H:%M:%S'))  # дата и время создания файла

    with open(os.path.join(FOLDER_PATH, FILE_NAME + f', {TOTAL_QR_CODES} - отчет, выгрузка кодов.txt'), 'w',
              encoding='utf-8') as new_file:
        print(f'{"-" * 10} ОТЧЕТ от {DATE_OF_PROCESSING} {"-" * 10}\n', file=new_file)
        print('Ошибка открытия:', file=new_file)
        print(LINKS_NOT_FORMAT if len(LINKS_NOT_FORMAT) > 0 else 'Все файлы обработаны как UTF-8', file=new_file)
        print(file=new_file)
        print('Задублированные коды:', file=new_file)
        print(REPEAT_CODES if len(REPEAT_CODES) > 0 else 'Дубликаты не найдены', file=new_file)
        print(file=new_file)
        print('Обработаны файлы:', file=new_file)
        for key, value in DATA.items():
            print(f'{key} - {len(value)} шт.', file=new_file)
        print(f'Итог: {TOTAL_QR_CODES} кодов обработано', file=new_file)

        print(file=new_file)
        print(f'{"-" * 10} ВЫГРУЗКА ВСЕХ КОДОВ {"-" * 10}\n', file=new_file)
        for key, value in DATA.items():
            print(f'{key}', file=new_file)
            for v in value:
                print(f'{v}', file=new_file)


def record_excel():  # запись в excel в 1 столбец
    wb = Workbook()
    ws = wb.active  # захватываем активный лист
    num_cell = 1
    for k, v in DATA.items():
        for code in v:
            cell = 'A' + str(num_cell)
            ws[cell] = code
            num_cell += 1
    wb.save(os.path.join(FOLDER_PATH, FILE_NAME + f', {TOTAL_QR_CODES} - для ДТ.xlsx'))  # имя файла xlsx


def output_of_results():  # вывод итоговой информации
    print('-' * 10)
    for key, value in DATA.items():
        print(f'{key} - {len(value)} шт.')
    print('-' * 10)
    print(f'Итог: {TOTAL_QR_CODES} кодов обработано')
    print(f'Дата и время: {DATE_OF_PROCESSING}')


def main():
    input_path()
    file_processing()
    replace_special_chars()
    record_report()
    record_excel()
    output_of_results()


if __name__ == "__main__":
    main()
