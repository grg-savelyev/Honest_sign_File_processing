from openpyxl import Workbook
from datetime import datetime

product_dict = {}


def add_product(item, qr):
    if item in product_dict:
        product_dict[item].append(qr)
    else:
        product_dict[item] = [qr]


# обработка исходника txt.
with open('Этикетка обувь (тест).txt', 'r', encoding='utf-8') as shoes:
    file = shoes.readlines()

    file_info, lines_to_add = '', 2  # сколько следующих строк добавить после удовлетворения условию
    codes, total_codes = [], 0

    for line in file:
        line = line.strip()
        # формирование названия файла: арт, наименование, цвет, размер, кол-во кодов
        if line.isdigit() and len(line) == 5 and line not in file_info:
            file_info = line
            lines_to_add = 2  # сбрасываем счётчик при добавлении строки

        elif lines_to_add > 0 and line.strip():  # сверяем, что строка не пустая
            file_info = ', '.join([file_info, line])
            lines_to_add -= 1

        # обработка кода
        if '(01)04' in line and '(21)' in line and len(line) == 35:
            total_codes += 1
            line = line.replace('(01)', '01').replace('(21)', '21')
            add_product(file_info, line)

# дата и время создания файла
dt = (datetime.now().strftime('%d.%m.%y / %H:%M:%S'))

# если найдено более 1 модели, то запись в отчет.
if len([k for k in product_dict.keys()]) > 1:
    print('ВАЖНО! В исходнике найдено более 1 модели! Смотри отчет.txt')
    with open(f'Отчет.txt', 'w', encoding='utf-8') as new_file:
        print(f'{dt}\n', file=new_file)
        for k, v in product_dict.items():
            print(f'    {k} - {len(v)} шт.', file=new_file)
            print('\n'.join(v), file=new_file)

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
wb.save(f'{file_info}, {total_codes}.xlsx')

# вывод итоговой информации о содержании файла
for k, v in product_dict.items():
    print(f'{k} - {len(v)} шт.')
print(f'Файл обработан {dt}')
