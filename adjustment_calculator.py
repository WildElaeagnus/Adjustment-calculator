# рассчет корректировки из эксель файла
# import numpy as np
import pandas as pd
from file_browser import *
import re
from collections import namedtuple

# DEBUG = True
DEBUG = False
xl_columns = namedtuple('xl_columns', 'data tu corr')
col = xl_columns(3, 8, 10)


if(DEBUG):
    # only worked after $pip install openpyxl
    df = pd.read_excel('C:/Users/akorz/Desktop/Adjustment-calculator/test_files/10. Свод начислений ТЭ2100-00812 с учетом кор-ки от 31.07.21.xlsx', index_col=0)  

if(not DEBUG):
    fb = file_browser_()
    fb.file_browser_()
    df = pd.read_excel(fb.filename, index_col=0)  

adj = df.iloc[:, [col.data, col.tu, col.corr]]
# print("df" + str(df))
# print("adj" + str(adj))

pog = adj.iloc[:, [2]]
# print(pog)

# pog.dropna(inplace = True)
# print(pog[:, [0]].str.find("Корректировочный СФ", 0))
# print(pog)
pos_list = []
OFFSET = 1
pos = 0 
count = 0
for i in pog.iterrows():
    pos += 1
    for j in i:
        if (str(j).find("Корректировочный СФ") > -1 ):
            count += 1
            pos_list.append(pos)
print("Значения, которые надо скорректировать: " + str(count))
pos_list = [x - OFFSET for x in pos_list]
print("Позиции в файле: " + str(pos_list))
index_data = df.iloc[pos_list, [col.data, col.tu, col.corr]]
# print(index_data)
# найти есть ли на позицию в 8 столбце выше или ниже ячейки с таким же номером теплоустановки
# поиск номера идет по позиции _ в строке
# найти в даных которое надо изменить все номера ЭУ
power_plant_num = []
for i in index_data.iterrows():
    # найти номер установки
    power_plant_num.append(int(str(re.findall(r'\d+'+'_', str(i)))[2:-3]))
print("power_plant_num")
print(power_plant_num)
pos_list_xl = [x + 2 for x in pos_list]
print("pos_list_xl")
print(pos_list_xl)
# найти в стоблбе 8 в каждой ячейке номер ЭУ
j = 1
for i in df.iloc[:, [col.data]].to_numpy():
    # взять индексы из массива и сравнить с данными из столба 10
    j+=1
    if (j in pos_list_xl):
        # print(df.iloc[j-3, [3]].to_numpy())
        # print(i) # это значения которые надо сложить с основной ячейкой
        # print(df.iloc[j-1, [3]].to_numpy())

        # print(df.iloc[j-3, [8]].to_numpy())
        # print(df.iloc[j-2, [8]].to_numpy())
        # print(df.iloc[j-1, [8]].to_numpy()) # это значения которые надо сложить с основной ячейкой

        j_3 = str(df.iloc[j-3, [col.tu]].to_numpy())
        j_2 = str(df.iloc[j-2, [col.tu]].to_numpy())
        j_1 = str(df.iloc[j-1, [col.tu]].to_numpy())

        # print(re.findall(r'\d+'+'_', j_3))
        # print(re.findall(r'\d+'+'_', j_2))

        # if номер ТУ j-3 == номер ТУ j-2
        if(str(re.findall(r'\d+'+'_', j_3)) == str(re.findall(r'\d+'+'_', j_2 ))):
            # df.at[row, col] = j_3 val + i val
            df.iloc[j-3, col.data] = df.iloc[j-3, col.data] + i
            df.iloc[j-2, col.corr] = 'solved ' + str(df.iloc[j-2, col.corr])
            print(df.iloc[j-3, col.data])

        # if номер ТУ j-1 == номер ТУ j-2
        if(str(re.findall(r'\d+'+'_', j_1)) == str(re.findall(r'\d+'+'_', j_2 ))):
            # df.at[row, col] = j_1 val + i val
            df.iloc[j-1, col.data] = df.iloc[j-1, col.data] + i
            df.iloc[j-2, col.corr] = 'solved ' + str(df.iloc[j-2, col.corr])
            print(df.iloc[j-1, col.data])
        
        # if nothin mathces
        if((str(re.findall(r'\d+'+'_', j_1)) == str(re.findall(r'\d+'+'_', j_2 ))) and (str(re.findall(r'\d+'+'_', j_3)) == str(re.findall(r'\d+'+'_', j_2 )))):
            print("неудалось найти ячейку")

# df.to_excel("output.xlsx", header=False, index=False)
# df.to_csv("out_csv.csv", encoding="Windows 1251", header=False, index=False) 
with pd.ExcelWriter("output.xlsx") as writer:
    df.to_excel(writer, header=False, index=False, )