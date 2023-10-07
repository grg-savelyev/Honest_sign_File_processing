from openpyxl import Workbook
from datetime import datetime

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
        if '(01)0' in line:
            total_codes += 1
            line = line.replace('(01)', '01').replace('(21)', '21')
            codes.append(line)

# запись в файл в txt
with open(f"{file_info}, {total_codes}.txt", 'w', encoding='utf-8') as new_file:
    for code in codes:
        print(code, file=new_file)  # запись в txt

# запись в excel
wb = Workbook()
ws = wb.active  # захватываем активный лист
num_cell = 1
for code in codes:
    cell = 'A' + str(num_cell)
    ws[cell] = code
    num_cell += 1
wb.save(f'{file_info}, {total_codes}.xlsx')

# дата и время создания файла
dt = (datetime.now().strftime('%d.%m.%y / %H:%M:%S'))

# вывод итоговой информации о содержании файла
print(f"Файл обработан. {dt}\n"
      f"{file_info} - {total_codes} штук.")
