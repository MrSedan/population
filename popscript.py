#!/usr/bin/env python3
__author__ = 'MrSedan'
import openpyxl, requests, re, os, xmltodict, datetime
from openpyxl.styles import Alignment

while True:
    try:
        name = input("Имя файла: ").replace(".xlsx","")
        # Открытие исходной таблицы
        wb = openpyxl.load_workbook(filename=f'./{name}.xlsx')
        sheet = wb[wb.sheetnames[0]]

        # Удаление измененной ранее таблицы(если она есть)
        if os.path.exists(f"./{name}Edited.xlsx"):
            os.remove(f"./{name}Edited.xlsx")
        print("Таблица успешно загружена!")
        break
    except PermissionError:
        print("Ошибка! Возможно итоговая табица уже где-то открыта.")

# Создание новой таблицы
wg = openpyxl.Workbook()
wg.create_sheet('Лист1')
wg.remove(wg['Sheet'])
sh = wg['Лист1']
for i,s in enumerate(["Города", "Даты", "Население"]):
    sh.cell(column=i+1,row=1).value = s
wg.save(f"./{name}Edited.xlsx")

# Определение доп. данных
max_row = len(sheet["A"])
a = "A3"
b = f"A26"
col = 1
row = 2

# Открытие пустой таблицы
wb2 = openpyxl.load_workbook(filename=f'./{name}Edited.xlsx')
sheet1 = wb2[wb2.sheetnames[0]]

for i in sheet[f"{a}:{b}"]:
    if not i[0].value: continue  # Пропуск пустых строк
    val = i[0].value
    date = sheet.cell(column=3, row=i[0].row).value
    data = requests.get(  # Запрос к Wikipedia
        f"https://ru.wikipedia.org/w/api.php?format=xml&action=query&list=search&srwhat=text&srsearch={val.split(' ',1)[1]}")
    doc = xmltodict.parse(data.text)  # Парсинг ответа
    try:
        sear = [i['@title'] for i in doc['api']['query']['search']['p'] if  # Выборка возможных страниц
                i['@title'].startswith(val.split(' ',1)[1]) and "(штат)" not in i['@title'] and 'район' not in i['@title']]
    except:
        continue
    ser = []
    f = []
    for j in sear:
        try:
            data = requests.get(f"https://ru.wikipedia.org/wiki/{j}")  # Получение кода страницы
            text = re.split(r'<th class="plainlist" style="width:40%;">Население</th>', data.text)                              #
            text[1] = text[1].replace("&#160;", "").replace(" ", "")                                                            #
            sert = re.search(r'(</span>|\"nowrap\">)+(\d{1,3}(?:\S*\d{3})*)(&#160;челов|<sup)', text[1])                        #
            ser.append(sert.group(0).replace("</span>", "").replace("&#160;", "").replace("<sup", "").replace("челов",          # Поиск числа и очистка от лишнего
                                                                                                        "").replace(            #
                "\"nowrap\">", "").split('<')[0])
            f.append(j)
        except:
            continue
            pass
    for k,t in enumerate(ser):
        sheet1.cell(column=col + 2, row=row).value = int(t)
        sheet1.cell(column=col + 2, row=row).number_format = "# ### ##0"
        sheet1.cell(column=col, row=row).value = val
        sheet1.cell(column=col+4, row=row).value = f[k]
        sheet1.cell(column=col + 1, row=row).value = date
        sheet1.cell(column=col + 1, row=row).alignment = Alignment(horizontal='right')
        row += 1
    # Запись новых данных в таблицу

sheet1['A1'].alignment = Alignment(horizontal='center')
sheet1['B1'].alignment = Alignment(horizontal='center')
sheet1['C1'].alignment = Alignment(horizontal='center')

# Сохранение полученной таблицы
wb2.save(f"./{name}Edited.xlsx")
print("Готово!")