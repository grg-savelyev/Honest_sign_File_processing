from openpyxl import Workbook
from datetime import datetime
import time
import os
import chardet


#  Создание списка ссылок на файлы txt в указанной директории (без проверки формата)
def file_search(directory_in_str: str) -> list:
    directory = os.fsencode(directory_in_str)
    txt_files = []
    for directory_file in os.listdir(directory):
        filename = os.fsdecode(directory_file)
        if filename.endswith(".txt"):
            txt_files.append(os.path.join(directory_in_str, filename))
    return txt_files


# Создание списка ссылок на файлы txt формата utf-8
def utf_search(path: str) -> list:
    links = []
    for txt_file in file_search(path):
        with open(txt_file, 'rb') as directory_file:
            result = chardet.detect(directory_file.read())
            encoding = result['encoding']
            if encoding.lower() in ('utf-8', 'utf-8-sig'):
                links.append(txt_file)
                # извлечения имени файла из ссылки
                # print(f'"{os.path.basename(txt_file)}" внесен в список на обработку.')
            else:
                print(f'Ошибка: Файл "{os.path.basename(txt_file)}" имеет кодировку {encoding}. Обработка отклонена.')
    return links


#  Заполнение словаря, где k - арт., название; v - список qr
def add_product(item, qr) -> None:
    value_exists = any(qr in values for values in product_dict.values())  # Проверка наличия qr в словаре
    if value_exists:
        print(f'Код {qr} повторяется.')
        qr = 'ОШИБКА! ' + qr + ' - повтор кода'
    if item in product_dict:
        product_dict[item].append(qr)
    else:
        product_dict[item] = [qr]


# CТАРТ
print('Укажите путь к папке для обработки:')
folder_path = input()  # указываем путь на каталог
product_dict = {}  # словарь, где k название файла, v список кодов
start_time = time.perf_counter()  # контроль времени, старт обработки
good_links: list = utf_search(folder_path)  # ссылки на файлы соответствующие utf-8
print('Запуск обработки. \n')

# Обработка файла txt
for link in good_links:
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

# дата и время создания файла
dt = (datetime.now().strftime('%d.%m.%y / %H:%M:%S'))

name_file = ", ".join(character.split(", ")[:2])  # Имя файла: арт, продукт
sum_total_qr = sum(len(value) for value in product_dict.values())  # Общее количество собранных qr

# запись финального словаря в txt
with open(f'{name_file}, {sum_total_qr}', 'w', encoding='utf-8') as new_file:
    print(f'{dt}\n', file=new_file)
    for key, value in product_dict.items():
        print(f'{key}', file=new_file)
        for v in value:
            print(f'{v}', file=new_file)

# запись в excel в 1 столбец
wb = Workbook()
ws = wb.active  # захватываем активный лист
num_cell = 1
for k, v in product_dict.items():
    for code in v:
        cell = 'A' + str(num_cell)
        ws[cell] = code
        num_cell += 1
wb.save(f'{name_file}, {sum_total_qr}.xlsx')  # имя файла xlsx

# вывод итоговой информации
print('-' * 10)
for k, v in product_dict.items():
    print(f'{k} - {len(v)} шт.')
print('-' * 10)
print(f'Итог: {sum_total_qr} кодов обработано')
print(f'Дата и время: {dt}')

# вывод времени обработки файла
end_time = time.perf_counter()
elapsed_time = end_time - start_time
print(f'Скорость обработки: {elapsed_time}')
