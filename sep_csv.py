
import pandas as pd
import re
# import mysql.connector

# Читаем csv файл
df = pd.read_csv(r'C:\Users\aaznu\pars_doc\base_mto_su.csv')

# Убираем строки с пустыми значениями без их разбора
df = df.dropna()

# Вводим названия столбцов для новой таблицы
new_col = ['name', 'equipment', 'list_of_licensed_software', 'building', 'floor', 'room_num', 'description', 'address']

# Дублируем строки в основную базу для полноты с отсечением шумов
full_res = [[name, eq, li] for name, room, eq, li in zip(df['name'], df['room'], df['equipment'],
                                                         df['list_of_licensed_software']) if len(room.split('\n')) == 4]


def repack_el(res_list):
    count = len(res_list)
    result = [res_list[0], res_list[1]]
    spl_count = 0
    if len(res_list[3]) < 5:
        result.append(res_list[2] + res_list[3])
        spl_count += 1
    else:
        result.append(res_list[2])
    if count - len(result) == 3 and '4430' in res_list[count - 1]:
        result.append(res_list[3]+res_list[4])
        spl_count += 1
        result.append(res_list[5])
    else:
        if count - len(result) == 3:
            result.append(res_list[3])
            result.append(res_list[4]+','+res_list[5])
            spl_count += 1
        else:
            describes = []
            if count - len(result) > 3 and '4430' in res_list[count - 1]:
                for el in res_list[spl_count * 2 + 3:count - 2]:
                    describes.append(el)
                result.append('/'.join(describes).lower())
                result.append(res_list[count - 1])
            else:
                for el in res_list[spl_count * 2 + 3:count - 2]:
                    describes.append(el)
                result.append('/'.join(describes).lower())
                result.append(res_list[count - 2]+'/'+res_list[count - 1])
    return result


def split_room():
    split_res = [el.split('\n') for el in df['room'] if len(el.split('\n')) == 4]
    add_res = []
    for el in split_res:
        # Стандартизируем записи
        # Убираем лишние пробелы.
        s = ' '.join(el)
        el = re.sub(r'\s+', ' ', s)
        el = re.sub(' ,', ',', el)
        el = re.sub(r'(^\s)', '', el, 1)
        # Разделяем строку символом ; убирая шумы в виде лишних знаков препинания и слов.
        if 'корпус' and 'этаж' and ('помещение' or 'кабинет') in el.lower():
            el = re.sub(r'(\Sорпус\s)|(№)|(\Sдрес\S\s)|(\Sтаж\s)|(\S\Sедиацентр\S\S)', '', el)
            el = re.sub(r'(?:\Sомещение\s)|(?:^ +)', '', el)
            el = re.sub(r'( +)|(,\s+)', ';', el, 3)
            el = re.sub(r'(\()|(\))', ';', el)
            el = re.sub(r'(;\s;)|(;.\s;)|(;;)|(;,,\s;)|(;этаж;)', ';', el)
            el = re.sub(r'(;$)|(;\s$)|(;;$)|(;\Sедиацентр)', '', el)
        else:
            el = re.sub(r'(\Sорпус\s)|(№)|(\Sдрес\S\s)|(\Sтаж\s)|(\S\Sедиацентр\S\S)', '', el)
            el = re.sub(r'(^\s)', '', el, 1)
            el = re.sub(r'( +)|(,\s+)', ';None;None;', el, 1)
            el = re.sub(r'(\()|(\))', ';', el)
            el = re.sub(r'(;$)|(;\s$)|(;;$)', '', el)

        # Дальше работаем с записью как со списком сохраняя длинну в 5 элементов
        el = el.split(';')
        if len(el) == 5:
            add_res.append(el)
        else:
            add_res.append(repack_el(el))
    return add_res


"""
add_res = []
for ls in split_res:
    place = []
    for st in ls:
        pattern = r'(\(*)(\)*)'        # Убираем скобки
        res = re.sub(pattern, '', st)
        place.append(res)
    add_res.append(place)
    # print(place)
"""


def get_data():
    add_res = split_room()
    data_list = []
    for i, el in enumerate(full_res):
        try:
            el.extend(add_res[i])
            data_list.append(el)
        except AttributeError:
            continue
    return data_list


def create_xlsx():
    data = pd.DataFrame(data=get_data(), columns=new_col)
    data.to_excel('mto_v2.xlsx', encoding='utf=8')


