# рассчет корректировки из эксель файла

import pandas as pd
import re
from collections import namedtuple
from file_browser import *

# DEBUG = True
DEBUG = False

xl_columns = namedtuple('xl_columns', 'data tu corr')
col = xl_columns(2, 8, 10)


if(DEBUG):
    # only worked after $pip install openpyxl
    df = pd.read_excel('C:/Users/akorz/Desktop/Adjustment-calculator/test_files/10. Свод начислений ТЭ2100-00812 с учетом кор-ки от 31.07.21.xlsx', index_col=0)  

if(not DEBUG):
    fb = file_browser_()
    fb.file_browser_()
    df = pd.read_excel(fb.filename, index_col=0)  

adj = df.iloc[:, [col.data, col.tu, col.corr]]

pog = adj.iloc[:, [2]]

OFFSET = 1

    
def find_in_df(string_to_find):
    pos_list_ = []
    pos = 0 
    count = 0
    for i in pog.iterrows():
        pos += 1
        for j in i:
            if (str(j).find(string_to_find) > -1 ):
                count += 1
                pos_list_.append(pos)
    return pos_list_

# найти номера строк, данные из которых надо пересчитать
corr_str = "Корректировочный СФ"
pos_list_corr = find_in_df(corr_str)

# найти первую ячейку в ряде заголовков столбцов
pos_lbs = find_in_df("Расчетный период")

# refactor later
# region 
pos_list_corr = [x - OFFSET for x in pos_list_corr]
print("Позиции в файле: " + str(pos_list_corr))
index_data = df.iloc[pos_list_corr, [col.data, col.tu, col.corr]]
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
pos_list_xl = [x + 2 for x in pos_list_corr]
print("pos_list_xl")
print(pos_list_xl)
# найти в стоблбе 8 в каждой ячейке номер ЭУ
j = 1
for i in df.iloc[:, [col.data]].to_numpy():
    # взять индексы из массива и сравнить с данными из столба 10
    j+=1
    if (j in pos_list_xl):
        j_3 = str(df.iloc[j-3, [col.tu]].to_numpy())
        j_2 = str(df.iloc[j-2, [col.tu]].to_numpy())
        j_1 = str(df.iloc[j-1, [col.tu]].to_numpy())

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
# endregion

dft = df.reset_index()
# убрать ненужные строки в начале
dft = dft.drop(range(pos_lbs[0]-1))
# поставить имена столбцов
dft = dft.rename(columns=dft.iloc[0])
# если столбец имеет НаН то надо его удалить
dft = dft.loc[:, dft.columns.notnull()]
dft = dft.reset_index()
dft = dft.drop(['index'], axis=1)

list_to_drop = (dft.index[dft['Вид СФ'] == 'solved Корректировочный СФ'].tolist()) 
dft = dft.drop(list_to_drop, axis=0)
# to_drop = ['Цена', 'Дата СФ']
to_drop = []

dft = dft.drop(to_drop, axis=1)
print(dft)



with pd.ExcelWriter("output.xlsx") as writer:
    dft.to_excel(writer, header=False, index=False, )

if not DEBUG: input()
