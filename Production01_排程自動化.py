import pandas as pd
import pandas.io.formats.excel
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import filedialog


def 取得日期區間(起始日期, 結束日期):
    # print(f'日期輸入範例 : {datetime.today().strftime("%Y/%m/%d")} 建議複製使用，確保格式正確')
    # start_date_str = input("請輸入起始日期 (YYYY/MM/DD): ")
    # end_date_str = input("請輸入結束日期 (YYYY/MM/DD): ")
    起始日期 = 起始日期.split(' ')[0]
    結束日期 = 結束日期.split(' ')[0]

    起始日期 = datetime.strptime(起始日期, "%Y/%m/%d")
    結束日期 = datetime.strptime(結束日期, "%Y/%m/%d")

    日期區間 = []
    當前日期 = 起始日期

    while 當前日期 <= 結束日期:
        日期區間.append(當前日期)
        當前日期 += timedelta(days=1)

    print(日期區間)
    return 日期區間


def 星期日判斷(date, mode=None):
    if date.weekday() == 6:  # 0 表示星期一，6 表示星期日
        if mode == '減法':
            # 成品如果是星期日就再減一天，不是就不用加
            date -= timedelta(days=1)
        else:
            # 如果是星期日就再加一天，不是就不用加
            date += timedelta(days=1)
    else:
        pass
    return date


def 分鐘判斷(date):
    # 去除不是整點的部分
    分鐘部分 = date.strftime('%M')
    if 分鐘部分 != '00' and 分鐘部分 != '30':
        date += timedelta(hours=1)
        date = date.replace(minute=00)
    else:
        pass

    return date


def 首件時間計算(row, date):
    # 基礎時間設定 固定+2小時 急件+1.5小時 特急+1小時
    if pd.isna(row['工單備註']):
        date = (date + timedelta(hours=2))
    elif '急件' in row['工單備註']:
        date = date + timedelta(hours=1, minutes=30)
    elif '特急' in row['工單備註']:
        date = date + timedelta(hours=1)
    else:
        date = (date + timedelta(hours=2))

    # 時間變數設定，用於下列判斷
    晚上七點半 = datetime.strptime('19:30', '%H:%M').time()
    凌晨零點 = datetime.strptime('00:00', '%H:%M').time()
    早上八點半 = datetime.strptime('8:30', '%H:%M').time()

    # 判斷結束時間是否超過下午7:30
    if date.time() > 晚上七點半:
        # 如果符合條件，將時間加一天並重置為早上9:30
        首件時間 = (date + timedelta(days=1)).replace(hour=9, minute=30)

    # 判斷結束時間是否為當天凌晨(介於00:00~8:30)，若是則重置為當天10:30
    elif 凌晨零點 <= date.time() <= 早上八點半:
        首件時間 = date.replace(hour=10, minute=30)

    else:  # 若不是則代表結束時間加上2小時後介於(8:30~19:30)之間，可直接填寫
        首件時間 = date

    # 同時進行分鐘和星期日判斷
    首件時間 = 星期日判斷(分鐘判斷(首件時間))
    首件時間_文字格式 = 首件時間.strftime('%#m/%#d %H:%M') + '前首件*5'
    return 首件時間_文字格式


def 押首件時間(data, index, row, 日期):
    if not pd.isna(row['結束時間']) and pd.isna(row['DIP首件產出時間/數量']):
        # 結束時間是計算基礎，此部分將結束時間轉換成可計算的時間格式
        結束時間 = row['結束時間'].to_pydatetime()

        # 首先針對404作條件判斷，因為404狀況不太一樣，再針對不是404的客戶判斷，條件為S+D
        # 如果是404且MO為3開頭，則必為RD工單，製程為S、D、S+D
        # 這裡判斷不等於404是必須的，因為404 S+D 測試和成品不一定打X
        # 且首件時間必為X(沒有首件)
        if (row['SOURCE'] == '四零四' and row['MO'][0] == '3') or (
                row['SOURCE'] != '四零四' and row['製程'] in
                ['S+D', 'S+D+P', 'D', 'S+D+P(P)', 'D+P(P)', 'S', 'S+P', 'S+P(P)']):
            data.at[index, 'DIP首件產出時間/數量'] = 'X'

        else:
            首件時間_文字格式 = 首件時間計算(row, 結束時間)
            data.at[index, 'DIP首件產出時間/數量'] = 首件時間_文字格式

        data.at[index, '填寫註記_1'] = 'Y'
    return data


def Output時間計算(row, 日期, 開始時間):
    # 判斷基礎為當天12:30
    當天十二點半 = 日期.replace(hour=12, minute=30)

    # 開始時間是否早於當天12點半
    if 開始時間 < 當天十二點半:
        # 開始時間一律都押下午3點半
        開始時間 = 開始時間.replace(hour=15, minute=30)
        if pd.isna(row['工單備註']):
            Output時間 = (開始時間 + timedelta(days=2))
        elif '急件' in row['工單備註']:
            Output時間 = (開始時間 + timedelta(days=1))
        elif '特急' in row['工單備註']:
            Output時間 = 開始時間
        else:
            Output時間 = (開始時間 + timedelta(days=2))
    else:  # 不符則代表晚於當天12:30
        開始時間 = 開始時間.replace(hour=15, minute=30)
        if pd.isna(row['工單備註']):
            Output時間 = (開始時間 + timedelta(days=3))
        elif '急件' in row['工單備註']:
            Output時間 = (開始時間 + timedelta(days=2))
        elif '特急' in row['工單備註']:
            Output時間 = (開始時間 + timedelta(days=1))
        else:
            Output時間 = (開始時間 + timedelta(days=3))

    # 判斷是否為星期日，並轉換成文字
    Output時間 = 星期日判斷(Output時間).strftime('%#m/%#d %H:%M')
    Output時間_文字格式 = Output時間 + '*' + str(int(row['工令量']))

    return Output時間_文字格式


def 押Output時間(data, index, row, 日期):
    # 這部分是用首件時間是否有排來控制是否要排Output時間，
    if pd.isna(row['OUTPUT']) and not pd.isna(row['開始時間']) and not pd.isna(row['DIP首件產出時間/數量']):
        開始時間 = row['開始時間'].to_pydatetime()
        Output時間_文字格式 = Output時間計算(row, 日期, 開始時間)
        data.at[index, 'OUTPUT'] = Output時間_文字格式
        data.at[index, '填寫註記_2'] = 'Y'

    return data


def 成品時間計算(row):
    交期 = row['預定完工日']
    成品時間 = 交期.replace(hour=16, minute=30)
    成品時間 -= timedelta(days=1)
    成品時間 = 星期日判斷(成品時間, mode='減法')

    成品時間_文字 = 成品時間.strftime('%#m/%#d %H:%M')
    成品時間_內文 = 成品時間_文字 + '*' + str(int(row['工令量']))

    return 成品時間, 成品時間_內文


def 押成品時間(data, index, row):
    if not pd.isna(row['DIP首件產出時間/數量']) and not pd.isna(row['OUTPUT']) \
            and pd.isna(row['成品']):
        母工單後段 = int(row['母工單單號'].split('-')[1])
        工單後段 = int(row['工號'].split('-')[1])
        工單差距 = 工單後段 - 母工單後段
        # 因為成品就算打'x'測時仍然有可能要押時間，所以這裡不像先前首件打X的就不用算，
        # 而是所有項目都計算出一個給測試用的判斷時間，到了測試的部分，再另外判斷是否使用該判斷時間
        成品時間, 成品時間_文字格式 = 成品時間計算(row)
        # 判斷順序 - 是否等於 404 3開頭的RD工單
        if (row['SOURCE'] == '四零四' and row['MO'][0] == '3') or (
                row['SOURCE'] != '四零四' and row['製程'] in
                ['S+D', 'S+D+P', 'D', 'S+D+P(P)', 'D+P(P)']):
            data.at[index, '成品'] = 'X'
            data.at[index, '判斷用時間_給測試用'] = 'X'
        # 經過前面的條件判斷，後面的項目必然不會是類似S+D的工單
        # 而這部分判斷為，但工單差距為1，且不是S+D這一類，代表不會有成品時間(不用包裝)
        # 即測試出貨客戶
        elif 工單差距 == 1:
            if row['SOURCE'] == '四零四':
                if '先請款暫放研騰待配對' in row['出貨日(Text)']:
                    data.at[index, '成品'] = 成品時間_文字格式
                    data.at[index, '判斷用時間_給測試用'] = 成品時間
                else:
                    data.at[index, '成品'] = 'X'
                    data.at[index, '判斷用時間_給測試用'] = 成品時間
            else:
                data.at[index, '成品'] = 'X'
                data.at[index, '判斷用時間_給測試用'] = 成品時間
        else:
            data.at[index, '成品'] = 成品時間_文字格式
            data.at[index, '判斷用時間_給測試用'] = 成品時間

        data.at[index, '填寫註記_4'] = 'Y'

    return data


def 測試時間計算(row):
    TEST時間 = row['判斷用時間_給測試用']
    if row['成品'] == 'X':
        TEST時間 = TEST時間.replace(hour=16, minute=30)
        TEST時間 = TEST時間.strftime('%#m/%#d  %H:%M')
        TEST時間_文字格式 = TEST時間 + '*' + str(int(row['工令量']))
    else:
        TEST時間 = (TEST時間 - timedelta(days=1))
        TEST時間 = 星期日判斷(TEST時間, mode='減法')
        TEST時間 = TEST時間.strftime('%#m/%#d')
        TEST時間_文字格式 = TEST時間 + '*' + str(int(row['工令量']))

    return TEST時間_文字格式


def 押測試時間(data, index, row):
    if pd.isna(row['TEST']) and not pd.isna(row['判斷用時間_給測試用']):
        # if row['SOURCE'] in 測試出貨客戶:
        #     TEST時間_文字格式 = 測試時間計算(row, mode='測試出貨')
        #     data.at[index, 'TEST'] = TEST時間_文字格式
        if row['判斷用時間_給測試用'] == 'X':
            data.at[index, 'TEST'] = 'X'
        else:
            TEST時間_文字格式 = 測試時間計算(row)
            data.at[index, 'TEST'] = TEST時間_文字格式
        data.at[index, '填寫註記_3'] = 'Y'

    return data


def 自動排程時間_DIP(data, 目標日期區間):
    for 日期 in 目標日期區間:
        # 此部分為抓出結束時間，並依據各種條件判斷，填入首件時間
        for index, row in data.iterrows():
            if row['線別'] in ['2201', '2202'] and \
                    row['結束時間'].date().strftime("%#m/%#d") == 日期.date().strftime("%#m/%#d"):
                # 這邊的寫法是有問題的，因為這兩個函數所用的row都來自於最初的data
                # 而押Output時間所用的row應該是押首件時間這個data的row
                # 因此真正的問題並不是出在data有沒有正確地給到押Output時間這個函式
                # 而是在於row並不是來自於最新的data
                # data = 押Output時間(data, index, row, 日期)
                data = 押首件時間(data, index, row, 日期)

        # 此部分為抓出開始時間，並依據各種條件判斷，填入Output時間
        for index, row in data.iterrows():
            if row['線別'] in ['2201', '2202'] and \
                    row['結束時間'].date().strftime("%#m/%#d") == 日期.date().strftime("%#m/%#d"):
                data = 押Output時間(data, index, row, 日期)

        # 此部分為成品時間
        for index, row in data.iterrows():
            if row['線別'] in ['2201', '2202'] and \
                    row['結束時間'].date().strftime("%#m/%#d") == 日期.date().strftime("%#m/%#d"):
                data = 押成品時間(data, index, row)

        # 此部分為TEST時間
        for index, row in data.iterrows():
            if row['線別'] == '2201' or row['線別'] == '2202' and \
                    row['結束時間'].date().strftime("%#m/%#d") == 日期.date().strftime("%#m/%#d"):
                data = 押測試時間(data, index, row)

    return data


def 自動排程時間_SMT(data, 目標日期區間):
    for 日期 in 目標日期區間:
        # 此部分為抓出結束時間，並依據各種條件判斷，填入首件時間
        for index, row in data.iterrows():
            if (row['線別'] == '2101' or row['線別'] == '2102') and row['製程'] == 'S' \
                    and row['結束時間'].date().strftime("%#m/%#d") == 日期.date().strftime("%#m/%#d"):
                data = 押首件時間(data, index, row, 日期)

        # 此部分為抓出開始時間，並依據各種條件判斷，填入Output時間
        for index, row in data.iterrows():
            if (row['線別'] == '2101' or row['線別'] == '2102') and row['製程'] == 'S':
                # 這部分是用首件時間是否有排來控制是否要排Output時間，
                data = 押Output時間(data, index, row, 日期)

            # 此部分為成品時間
        for index, row in data.iterrows():
            if (row['線別'] == '2101' or row['線別'] == '2102') and row['製程'] == 'S':
                data = 押成品時間(data, index, row)

            # 此部分為TEST時間
        for index, row in data.iterrows():
            if (row['線別'] == '2101' or row['線別'] == '2102') and row['製程'] == 'S':
                data = 押測試時間(data, index, row)

    return data


