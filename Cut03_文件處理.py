""" 接收Cut01與02程式碼，用來處理讀進來的文件 """
from Cut01_剪下去_DIP import *
from Cut02_剪下去_SMT import *
from Paste01_貼上去_DIP import *
from Paste02_貼上去_SMT import *
import pandas as pd
import os
import platform
import datetime


def 獲取桌面路徑():
    # 獲取使用者電腦系統 windows Linux等
    系統 = platform.system()
    目錄 = os.path.expanduser("~")

    if 系統 == "Windows":
        return os.path.join(目錄, "Desktop").replace('\\', '/')
    elif 系統 == "Darwin":  # macOS
        return os.path.join(目錄, "Desktop").replace('\\', '/')
    elif 系統 == "Linux":
        return os.path.join(目錄, "Desktop").replace('\\', '/')
    else:
        # 默認返回用戶主目錄
        return 目錄.replace('\\', '/')


def 文件處理_剪下去(文件路徑, 目標日期):
    data_DIP = pd.read_excel(f'{文件路徑}', header=1, sheet_name='DIP')
    data_SMT = pd.read_excel(f'{文件路徑}', header=1, sheet_name='SMT')
    data_DIP = 剪下去_for_DIP(data_DIP, 目標日期)
    data_SMT = 剪下去_for_SMT(data_SMT, 目標日期)

    桌面路徑 = 獲取桌面路徑()
    當天日期_文字格式 = datetime.datetime.now().strftime('%y%m%d%H%M')
    輸出檔名 = '剪下結果' + 當天日期_文字格式 + '.xlsx'
    writer = pd.ExcelWriter(f'{桌面路徑}/{輸出檔名}', engine='xlsxwriter')

    data_DIP.to_excel(writer, index=False, sheet_name='DIP')
    data_SMT.to_excel(writer, index=False, sheet_name='SMT')

    writer.close()


def 文件處理_貼上去(文件路徑, 起始日期, 結束日期):
    data_DIP = pd.read_excel(f'{文件路徑}', header=1, sheet_name='DIP')
    data_SMT = pd.read_excel(f'{文件路徑}', header=1, sheet_name='SMT')
    data_DIP = 貼上去_for_DIP(data_DIP, 起始日期, 結束日期)
    data_SMT = 貼上去_for_SMT(data_SMT, 起始日期, 結束日期)

    桌面路徑 = 獲取桌面路徑()
    當天日期_文字格式 = datetime.datetime.now().strftime('%y%m%d%H%M')
    輸出檔名 = '貼上結果' + 當天日期_文字格式 + '.xlsx'
    writer = pd.ExcelWriter(f'{桌面路徑}/{輸出檔名}', engine='xlsxwriter')

    data_DIP.to_excel(writer, index=False, sheet_name='DIP')
    data_SMT.to_excel(writer, index=False, sheet_name='SMT')

    writer.close()
