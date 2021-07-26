# рассчет корректировки из эксель файла

import pandas as pd
from file_browser import *

DEBUG = True
# DEBUG = False



if(DEBUG):
    df = pd.read_excel('C:/Users/akorz/Desktop/Adjustment-calculator/test_files/10. Свод начислений ТЭ2100-00812 с учетом кор-ки от 31.07.21.xlsx', index_col=0)  

if(not DEBUG):
    fb = file_browser_()
    fb.file_browser_()
    df = pd.read_excel(fb.filename, index_col=0)  


print(df)