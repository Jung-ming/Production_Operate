from Cut01_剪下去_DIP import *
import pandas as pd


def 文件處理(文件路徑):
    data_DIP = pd.read_excel(f'{文件路徑}', header=1, sheet_name='DIP')
    data_DIP = 剪下去_for_DIP(data_DIP)
    data_DIP.to_excel('結果_DIP.xlsx', index=False)
