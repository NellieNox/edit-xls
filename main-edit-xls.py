import xlrd
import os
import datetime
#import tkinter

book = xlrd.open_workbook(r'C:\\Users\\Nellie\\dev\\edit-xls\\req.xls')    # открываем "книгу"
sheet = book.sheet_by_index(2)    # берём лист с индексом 2 - т.е. третий лист
mas_data = []       # список для хранения данных
row_num = 26        # с этой строки (27-ой) начинается список фамилий
f = open('C:\\Users\\Nellie\\dev\\edit-xls\\text.txt', 'w')

while (sheet.cell_value(row_num, 2) != ''):     # перебираем данные, пока не наткнёмся на пустую ячейку. Второй параметр = 2, т.к. нужен столбец С
    birthday = datetime.datetime(*xlrd.xldate_as_tuple(sheet.cell_value(row_num, 3), book.datemode))    # получаем дату рождения и преобразуем из float в datetime
    full_name = sheet.cell_value(row_num, 2).split()
    f.write(full_name[0] + ',' + birthday.strftime("%d.%m.%Y") + '\n')
    row_num = row_num + 1       # переходим на другую строку

f.close()
