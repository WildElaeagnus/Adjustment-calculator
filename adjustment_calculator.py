# рассчет корректировки из эксель файла

from numpy import NaN
import pandas as pd
import re
from collections import namedtuple

from file_browser import *

# DEBUG = True
DEBUG = False

xl_columns = namedtuple('xl_columns', 'data tu corr')
col = xl_columns(2, 8, 10)

cols_names = namedtuple('cols_names', '''
                            rp 
                            col 
                            sum 
                            numc 
                            num 
                            tu 
                            dateCF 
                            typeCF 
                            price''')
col_names = cols_names('Расчетный период', 
                            'Количество', 
                            'Сумма', 
                            'Номенклатура.Код', 
                            'Номенклатура', 
                            'Теплоустановка', 
                            'Дата СФ', 
                            'Вид СФ', 
                            'Цена')
if(DEBUG):
    # only worked after $pip install openpyxl
    filepath = 'C:/Users/akorz/Desktop/Python_code/Adjustment-calculator/test_files/10. Свод начислений ТЭ2100-00812 с учетом кор-ки от 31.07.21.xlsx'
    # filepath = 'C:/Users/akorz/Desktop/Python_code/Adjustment-calculator/test_files/10. Свод начислений 717108ОДН.xlsx'
    
    df = pd.read_excel(filepath, index_col=0,)  

if(not DEBUG):
    fb = file_browser_()
    fb.file_browser_()
    df = pd.read_excel(fb.filename, index_col=0, )  

adj = df.iloc[:, [col.data, col.tu, col.corr]]

pog = adj.iloc[:, [2]]
    
def find_in_df(string_to_find):
    pos_list_ = []
    count = 0
    for pos, i in enumerate(pog.iterrows(), start=1):
        for j in i:
            if (str(j).find(string_to_find) > -1 ):
                count += 1
                pos_list_.append(pos)
    return pos_list_

# найти номера строк, данные из которых надо пересчитать
corr_str = "Корректировочный СФ"
corr_str_ = "Исправление СФ"
corr_list = [corr_str, corr_str_]
# pos_list_corr = find_in_df(corr_str)

# найти первую ячейку в ряде заголовков столбцов
pos_lbs = find_in_df("Расчетный период")

df = df.reset_index()
# убрать ненужные строки в начале
df = df.drop(range(pos_lbs[0]-1))
# поставить имена столбцов
df = df.rename(columns=df.iloc[0])
# если столбец имеет НаН то надо его удалить
df = df.loc[:, df.columns.notnull()]
df = df.reset_index()
df = df.drop(['index'], axis=1)

dfp = df.sort_values([col_names.rp, col_names.tu, col_names.numc])

dfpi = dfp.reset_index()
# убираем пустые значения из столба количество
dfpi.dropna(subset = [col_names.col], inplace=True)
dfpi = dfpi[[col_names.rp, col_names.tu, col_names.col, col_names.numc, col_names.typeCF]]
dfpi2 = dfpi
with pd.ExcelWriter("before.xlsx") as writer:
    dfpi2.to_excel(writer, header=True, index=False, sheet_name='before' )

# извлекает номер ТУ из ячейки
def re_str(cell):
    return (str(re.findall(r'\d+'+'_', str(cell)))[2:-3])
# список со строками в которых лежат исправления
i_pos = []
for i, row in dfpi.iterrows():
    if (row[col_names.typeCF] == corr_str or row[col_names.typeCF] == corr_str_):
        i_pos.append(i)
# список индексов с рядами, которые надо удалить в конце работы программы
drop_list = []

print(dfpi)
print(i_pos)
for i in i_pos:
    # расчетный период совпадает выше
    if (dfpi.at[i, col_names.rp] == dfpi.at[i-1, col_names.rp]):
        # если номер ТУ совпадает с номером ячейки ВЫШЕ
        if re_str(dfpi.at[i, col_names.tu]) == re_str(dfpi.at[i-1, col_names.tu]):
            # их номера номенклатура код совпадают
            if (dfpi.at[i, col_names.numc] == dfpi.at[i-1, col_names.numc]):
                # если вид СФ верхней ячейки пусто то надо туда добавить колличество из нижней
                # а текущую удалить
                if dfpi.isnull().at[i-1, col_names.typeCF]: 
                    dfpi.at[i-1, col_names.col] += dfpi.at[i, col_names.col]
                    drop_list.append(i)
                    # изменить ячейку, чтоб при сравнении с нижними она не учитывалась
                    dfpi.at[i, col_names.typeCF] = 'solved'
    # расчетный период совпадает ниже
    if (dfpi.at[i, col_names.rp] == dfpi.at[i+1, col_names.rp]):
        # если ниже номр ТУ совпадает
        if (re_str(dfpi.at[i, col_names.tu]) == re_str(dfpi.at[i+1, col_names.tu])):
            # их номера номенклатура код совпадают
            if (dfpi.at[i, col_names.numc] == dfpi.at[i+1, col_names.numc]):
                # если вид СФ нижней ячейки пусто то надо туда добавить колличество из нижней
                # а текущую удалить
                if dfpi.isnull().at[i+1, col_names.typeCF]: 
                    dfpi.at[i+1, col_names.col] += dfpi.at[i, col_names.col]
                    drop_list.append(i)
                    dfpi.at[i, col_names.typeCF] = 'solved'
                # если ячейка ниже = Корректировочный, то в нее добавляем данные из этой ячейки
                # и ставим значение next line
                if dfpi.at[i+1, col_names.typeCF] == corr_str :
                    dfpi.at[i+1, col_names.col] += dfpi.at[i, col_names.col]
                    drop_list.append(i)
                    dfpi.at[i, col_names.typeCF] = 'next line'
                # # если ячейка ВЫШЕ имеет значение next line, а ниже ЕСТЬ корректировочный 
                # # то надо ... в теории то же самое что и без next line
                # if (dfpi.at[i+1, col_names.typeCF] == corr_str) and (dfpi.at[i-1, col_names.typeCF] == 'next line'):
                    
            # соответсвенно если ячейка ВЫШЕ имеет значение next line, а ниже значения корректировочный для 
            # этой ТУ уже нет или номенклатура код другая, то надо:
            #   очистить Вид СФ
            #   НЕ добавлять в лист на удаление drop_list
            #   записать сумму верхней и текущей в столбце колличество
    if (dfpi.at[i-1, col_names.typeCF] == 'next line') :
        if (dfpi.at[i, col_names.typeCF] != 'next line'):
            dfpi.at[i, col_names.typeCF] = NaN


# пометить все неудачные рассчеты как невыполненные
dfpi.loc[dfpi[col_names.typeCF] == corr_str, [col_names.typeCF]] = 'unable to solve'
dfpi.loc[dfpi[col_names.typeCF] == corr_str_, [col_names.typeCF]] = 'unable to solve'

# в конце удаляем строки ненужные +НаН в колонке количество
with pd.ExcelWriter("after.xlsx") as writer:
    dfpi.to_excel(writer, header=True, index=False, sheet_name='after' )
print(dfpi)
    
dfpi = dfpi.drop(drop_list, axis=0)
dfpi = dfpi[dfpi["Расчетный период"] != "Расчетный период"]

# убрать пустые значения из столба теплоустановок
dfpi.dropna(subset=[col_names.tu], inplace=True)

dfpi.loc["Total", col_names.col] = dfpi[col_names.col].sum()
dfpi.at["Total",col_names.rp] = "Total"
with pd.ExcelWriter("output.xlsx") as writer:
    dfpi.to_excel(writer, header=True, index=False, )

if not DEBUG: input()
