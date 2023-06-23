""" 接收Cut01與02程式碼，用來處理讀進來的文件 """
from Cut01_剪下去_DIP import *
from Cut02_剪下去_SMT import *
from Paste01_貼上去_DIP import *
from Paste02_貼上去_SMT import *
from Production01_排程自動化 import *
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


def 排程時間寫入(data):
    工作表名稱 = ['DIP', 'SMT']
    隱藏行 = ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K',
           'L', 'M', 'O', 'P', 'Q', 'R', 'S',
           'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB',
           'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ',
           'AM', 'AN', 'AO', 'AV', 'AW', 'AX', 'AY', 'AZ',
           'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH',
           'BI', 'BJ', 'BK', 'BL', 'BM', 'BO']

    桌面路徑 = 獲取桌面路徑()
    當天日期_文字格式 = datetime.datetime.now().strftime('%y%m%d%H%M')
    輸出檔名 = '排程結果' + 當天日期_文字格式 + '.xlsx'
    print(f'{桌面路徑}/{輸出檔名}')
    writer = pd.ExcelWriter(f'{桌面路徑}/{輸出檔名}', engine='xlsxwriter')

    # 格式設定
    正常格式 = writer.book.add_format({'font_size': 11, 'text_wrap': True})
    寫入格式 = writer.book.add_format({'font_size': 11, 'text_wrap': True, 'font_color': 'red'})

    for 足標, 工作表 in enumerate(data):
        工作表.to_excel(writer, sheet_name=工作表名稱[足標], index=False)
        worksheet = writer.sheets[工作表名稱[足標]]

        for 隱藏 in 隱藏行:
            worksheet.set_column(f'{隱藏}:{隱藏}', None, None, {'hidden': True})
        # AP 41 DIP首件
        # AQ 42 Output
        # 43 TSET
        # 44 成品
        # 37 結束時間
        # 36 開始時間
        worksheet.set_column(41, 41, 18, 正常格式)
        worksheet.set_column(42, 42, 18, 正常格式)
        worksheet.set_column(43, 43, 18, 正常格式)
        worksheet.set_column(44, 44, 18, 正常格式)
        worksheet.set_column(37, 37, 20)
        worksheet.set_column(36, 36, 20)

        for index, row in 工作表.iterrows():
            if not pd.isna(row['填寫註記_1']):
                worksheet.write(index + 1, 41, 工作表.iloc[index, 41], 寫入格式)
            if not pd.isna(row['填寫註記_2']):
                worksheet.write(index + 1, 42, 工作表.iloc[index, 42], 寫入格式)
            if not pd.isna(row['填寫註記_3']):
                worksheet.write(index + 1, 43, 工作表.iloc[index, 43], 寫入格式)
            if not pd.isna(row['填寫註記_4']):
                worksheet.write(index + 1, 44, 工作表.iloc[index, 44], 寫入格式)

    writer.close()


def 文件處理_排程自動化(data, 起始日期, 結束日期):
    目標日期區間 = 取得日期區間(起始日期, 結束日期)
    # 讀取檔案
    data_DIP = pd.read_excel(data, header=1, sheet_name='DIP')
    data_SMT = pd.read_excel(data, header=1, sheet_name='SMT')

    # 新增判斷用欄位
    data_DIP = data_DIP.assign(判斷用時間_給測試用=None, 填寫註記_1=None, 填寫註記_2=None, 填寫註記_3=None, 填寫註記_4=None)
    data_SMT = data_SMT.assign(判斷用時間_給測試用=None, 填寫註記_1=None, 填寫註記_2=None, 填寫註記_3=None, 填寫註記_4=None)

    # 根據條件自動產生Output
    data_DIP = 自動排程時間_DIP(data_DIP, 目標日期區間)
    data_SMT = 自動排程時間_SMT(data_SMT, 目標日期區間)

    # 將排好時間的資料寫入Excel
    排程時間寫入([data_DIP, data_SMT])
