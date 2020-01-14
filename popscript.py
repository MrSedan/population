#!/usr/bin/env python3
import openpyxl, wikipedia, requests, re
from bs4 import BeautifulSoup
wb = openpyxl.load_workbook(filename='./test.xlsx')
sheet = wb[wb.sheetnames[0]]
a = input()
b = input()
for i in sheet[a:b]:
    val = i.value
    sear = wikipedia.search(val)
    [print(sear.index(i)+1,i) for i in sear]
    n = int(input())-1
    wik = wikipedia.page(sear[n])
    data = requests.get(f"https://ru.wikipedia.com/wiki/{val}")
    text = re.split(r'<th class="plainlist" style="width:40%;">Население</th>',data.text)
    ser = re.search(r"</span>(\d{1,3}(?:\S*\d{3})*)<sup", text[1])
    ser = ser.group(0).replace("</span>","").replace("&#160;","").replace("<sup","")
    print(ser)
