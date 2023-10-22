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
                print(f'"{os.path.basename(txt_file)}" внесен в список на обработку.')
            else:
                print(f'Ошибка: Файл "{os.path.basename(txt_file)}" имеет кодировку {encoding}. Обработка отклонена.')
    return links


#  Заполнение словаря, где k - арт., название; v - список qr
def add_product(item, qr) -> None:
    value_exists = any(qr in values for values in product_dict.values())  # Проверка наличия qr в словаре
    if value_exists:
        # print(f'Код {qr} повторяется.')
        qr = qr + ' - повтор кода'
    if item in product_dict:
        product_dict[item].append(qr)
    else:
        product_dict[item] = [qr]


# CТАРТ ПРОГРАММЫ
print('Укажите путь к папке для обработки:')
folder_path = input()  # указываем путь на каталог
product_dict = {}  # словарь, где k название файла, v список кодов
start_time = time.perf_counter()  # контроль времени, старт обработки
good_links: list = utf_search(folder_path)  # ссылки на файлы соответствующие utf-8
print()
print('Запуск обработки.')

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
    print(f'"{os.path.basename(link)}" обработан.')

# дата и время создания файла
dt = (datetime.now().strftime('%d.%m.%y / %H:%M:%S'))

# ДОРАБОТАТЬ ОТЧЕТ!!!
# если найдено более 1 модели, то запись в отчет.
# if len([k for k in product_dict.keys()]) > 1:
#     print('ВАЖНО! В исходнике найдено более 1 модели! Смотри отчет.txt')
#     with open(f'Отчет.txt', 'w', encoding='utf-8') as new_file:
#         print(f'{dt}\n', file=new_file)
#         for k, v in product_dict.items():
#             print(f'    {k} - {len(v)} шт.', file=new_file)
#             print('\n'.join(v), file=new_file)

sum_total_qr = sum(len(value) for value in product_dict.values())  # Общее количество собранных qr

"""
запись в excel в 1 столбец и не зависит от кол-ва моделей.
См. отчет.
Если в исходнике нескольких моделей, название файла по последнему найденному и сумме QR кодов.
"""
wb = Workbook()
ws = wb.active  # захватываем активный лист
num_cell = 1
for k, v in product_dict.items():
    for code in v:
        cell = 'A' + str(num_cell)
        ws[cell] = code
        num_cell += 1
wb.save(f'{", ".join(character.split(", ")[:2])}, {sum_total_qr}.xlsx')  # имя файла xlsx

# вывод итоговой информации о содержании файла
print()
print('Итог:')
for k, v in product_dict.items():
    print(f'{k} - {len(v)} шт.')
print(f'Файл обработан {dt}')

# вывод времени обработки файла
end_time = time.perf_counter()
elapsed_time = end_time - start_time
print(f'Время работы программы = {elapsed_time}')

# pprint(product_dict)
