# Код для python 3.6 (C:\Python36-32\python readExcel.py)

import json
import niswitch
import pylightxl
import time


# чтение xlsx страниц groups, matrixs и возврат строки в формате JSON
# file(строка) - путь к файлу
# результат данной функции планируется обрабатывать в get_group() для выбора с какой строкой работать
# возвращает СТРОКУ JSON {groups:[{комутация1},{комутация2},{комутация3},...], matrixs:[имя матицы1, имя матрицы2,...]}
# {комутация} - это {"name": имя коммутации, "code":[[точка],[точка],...]}
# например {"groups": [{"name": "тестовая", "code": [["c43", "r0", "эталоны"],..., "matrixs": ["эталоны", "Сторона_А"...
def read_xlsx(file):
    # 1) получение групп коммутаций
    arr_groups = []
    workbook = pylightxl.readxl(file)
    nrows = len(workbook.ws('groups').col(1))  # количество строк

    # определяю номер столбца с кодом коммутации
    col_code = workbook.ws('groups').row(1).index('Код коммутации') + 1
    row = 2  # начинать с первой строки, так как нулевая строка это шапка
    while (row <= nrows):
        arr_groups.append({"name":workbook.ws('groups').index(row=row, col=2),
                           "code": json.loads(workbook.ws('groups').index(row=row, col=col_code))})
        row += 1


    # 2) получение списка матриц (имена устройств)
    arr_matrixs = []
    number_rows = len(workbook.ws('matrixs').col(1))  # количество строк
    row = 2  # начинать с первой строки, так как нулевая строка это шапка
    while (row <= number_rows):
        if workbook.ws('matrixs').index(row=row, col=1):
            arr_matrixs.append(workbook.ws('matrixs').index(row=row, col=1))
        row += 1
    return json.dumps({"groups": arr_groups, "matrixs": arr_matrixs} , ensure_ascii=False)


# из массива с вариантами коммутации [{"name": "тестовая", "code": [["c43", "r0", "эталоны"],...]
# возвращает "code" для группы с именем name
def get_group(json_data, name):
    #json_data = json.loads(str_data)
    for group in json_data:
        if group["name"] == name:
           #x = json.dumps(group["code"], ensure_ascii=False)
           return group["code"]
    return []

# отправляет в матрицу команду на комутацию
def connect_matrix(c, r, device):
    with niswitch.Session(device) as session:
        session.connect(c, r)


# функция получает реле которые нужно скомутировать (str_json_data), сканирует матрицы перечисленные
# в массиве arr_matrixs, а затем
# 1) размыкает реле которые замкнуты но при этом не входят в str_json_data
# 2) проверяет произошло ли размыкание реле (предосторожность)
# 3) замыкает оставшиеся реле которые входят в str_json_data но при этом не замкнуты

# (если session.get_relay_position вернуло 11 - реле закрыто, 10 - открыто )
def connect_groups(str_json_data, name, delay):
    try:
        obj_json_data = json.loads(str_json_data)
        code = get_group(obj_json_data['groups'], name)
        if len(code) == 0:
            return 'Группы с таким именем не существует или группа не имеет коммутаций!'
        # 1) размыкаю реле котоорые не должны быть скомутированы
        for matrix in obj_json_data['matrixs']:
            arr_closed = [];  # то что осталось закрытым после размыкания ненужных реле
            arr_opened = [];  # реле которые долны быть разомкнуты после окончания работы функции в формате сторки [{"rele":"kcard1ab", "matrix": "Device1" ...]
            with niswitch.Session(matrix) as session:
                number_col = session.num_of_columns
                number_row = session.num_of_rows
                row = 0
                while row < number_row:
                    # перебор шин
                    if session.get_relay_position('kcard1ab' + str(row)).value == 11:
                        flag_point_need = False
                        for point in code:
                            if point[2] != matrix:
                                continue
                            if point[0] == ('ab' + str(row)):
                                flag_point_need = True
                                break
                        if not flag_point_need:
                            session.disconnect('card1r' + str(row), 'ab' + str(row))
                            arr_opened.append('kcard1ab' + str(row))
                        arr_closed.append(['ab' + str(row), 'r' + str(row), matrix])
                    col = 0

                    # перебор колонок
                    while col < number_col:
                        if session.get_relay_position('kcard1r' + str(row) + 'c' + str(col)).value == 11:
                            flag_point_need = False
                            for point in code:
                                if point[2] != matrix:
                                    continue
                                if point[0] == ('c' + str(col)) and point[1] == ('r' + str(row)):
                                    flag_point_need = True
                                    break
                            if not flag_point_need:
                                session.disconnect('card1r' + str(row), 'c' + str(col))
                                arr_opened.append('kcard1ab' + str(row))
                            arr_closed.append(['c' + str(col), 'r' + str(row), matrix])
                        col += 1
                    row += 1

                # 2) проверка результата шага 1 (произошло ли размыкание ненужных реле)
                for point in arr_opened:
                    if point[2] != matrix:
                        continue
                    if session.get_relay_position(point).value == 11:
                        return 'Реле ' + point + ' не удалось разомкнуть!'

                # 3) замыкаю требуемые реле которые не были замкнуты ранее
                for point in code:
                    if point[2] != matrix:
                        continue
                    flag_point_close = False # True если реле уже закрыто
                    for closed_point in arr_closed:
                        if point[0] == closed_point[0] and point[1] == closed_point[1]:
                            flag_point_close = True
                            break
                    if not flag_point_close:
                        if 'ab' in point[0]:  # для реле шинразрешение на коммутацию для реде шин
                            session.channels[point[0]].analog_bus_sharing_enable = True
                        session.connect('card1' + point[1], point[0])

                # 4) проверка замкнулись ли реле которые необходимо замкнуть
                for point in code:
                    if point[2] != matrix:
                        continue
                    if 'ab' in point[0]:  # для реле шин
                        if session.get_relay_position('kcard1' + point[0]).value != 11:
                            return 'Не удалось замкнуть kcard1' + point[0] + '!'
                    else:  # для обычных реле
                        if session.get_relay_position('kcard1' + point[1] + point[0]).value != 11:
                            return 'Не удалось замкнуть kcard1' + point[1] + point[0] + '!'
        time.sleep(delay)
        return 'ok'
    except Exception as ex:
        return ex
