""" 接收Cut01與02程式碼，用來處理讀進來的文件 """
from Cut01_剪下去_DIP import *
from Cut02_剪下去_SMT import *
import pandas as pd


def 文件處理(文件路徑, 目標日期):
    data_DIP = pd.read_excel(f'{文件路徑}', header=1, sheet_name='DIP')
    data_SMT = pd.read_excel(f'{文件路徑}', header=1, sheet_name='SMT')
    data_DIP = 剪下去_for_DIP(data_DIP, 目標日期)
    data_SMT = 剪下去_for_SMT(data_SMT, 目標日期)
    # data_DIP.to_excel('結果_DIP.xlsx', index=False)
    # data_SMT.to_excel('結果_SMT.xlsx', index=False)

    # {資料夾路徑}/{輸出檔名}
    writer = pd.ExcelWriter(f'//file-server/生管部/五股廠/Jimmy/剪下結果.xlsx', engine='xlsxwriter')

    data_DIP.to_excel(writer, index=False, sheet_name='DIP')
    data_SMT.to_excel(writer, index=False, sheet_name='SMT')

    writer.save()
