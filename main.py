""" название файла: арт, наименование, цвет, размер """


def file_name(file):
    file_information = []
    lines_to_add = 2  # Сколько следующих строк добавить после удовлетворения условию

    for line in file:
        line = line.strip()

        if len(line) == 5 and line not in file_information:
            # it_num = line
            file_information.append(line)
            lines_to_add = 2  # Сбрасываем счетчик при добавлении строки

        elif lines_to_add > 0 and line.strip():  # Проверяем, что строка не пустая
            file_information.append(line)
            lines_to_add -= 1

    return file_information


with open('shoes.txt', 'r', encoding='utf-8') as shoes:
    original_file = shoes.readlines()
    file_information = ', '.join(file_name(original_file))

    with open(f"{file_information}.txt", 'w', encoding='utf-8') as new_file:
        total_codes = 0
        for line in original_file:
            if ('(01)0') in line:
                total_codes += 1
                line = line.lstrip().rstrip('\n')
                line = line.replace('(01)', '01').replace('(21)', '21')
                print(line, file=new_file)

        print()
        print(f"Файл обработан.\n"
              f"{file_information} - {total_codes} штук.")
