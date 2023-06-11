from datetime import datetime
import numpy as np
import pandas as pd


def 抓取待處理足標(data, 目標線別, 起始日期, 結束日期):
    # return會導致函數立即結束!!
    # 讀取FFFF線別，並將指定日期區間的資料按類型區分
    FFFF線別 = data.query('線別 == "FFFF"')
    操作足標_四零四 = []
    操作足標_重工 = []
    操作足標_H2O = []
    操作足標_營邦 = []
    操作足標_其他 = []
    for 足標, 欄位 in FFFF線別.iterrows():
        if 起始日期 <= 欄位['預定開工日'].to_pydatetime() <= 結束日期 and '暫不生產' not in str(欄位['工單備註']):
            if '重工' in 欄位['名稱規格']:
                操作足標_重工.append(足標)
            elif 欄位['SOURCE'] == 'H2O':
                操作足標_H2O.append(足標)
            elif 欄位['SOURCE'] == '四零四':
                操作足標_四零四.append(足標)
            elif 欄位['SOURCE'] == '營邦':
                操作足標_營邦.append(足標)
            else:
                操作足標_其他.append(足標)
    if 目標線別 == '2109':
        return 操作足標_重工
    elif 目標線別 == '2103':
        return 操作足標_H2O
    elif 目標線別 == '2102':
        return 操作足標_四零四
    elif 目標線別 == '2105':
        return 操作足標_營邦
    else:
        return 操作足標_其他


def 插入指定線別(data, columns, 目標線別, 起始日期, 結束日期):
    操作足標 = 抓取待處理足標(data, 目標線別, 起始日期, 結束日期)
    線別 = data.query(f'線別 == "{目標線別}"')
    目標線別最後足標 = 線別.index[-1]
    工序最末項 = 線別['工序'].iloc[-1]

    for 足標, 欄位 in data.query(f'index == {操作足標}').iterrows():
        欄位['線別'] = 目標線別
        欄位['工序'] = 工序最末項 + 1
        欄位['線別+工序'] = int(欄位['線別'] + str(int(欄位['工序'])).zfill(3))
        data = pd.DataFrame(np.insert(data.values, 目標線別最後足標 + 1, values=欄位, axis=0))
        目標線別最後足標 += 1
        工序最末項 += 1

    data.columns = columns
    操作足標 = [x + len(操作足標) for x in 操作足標]
    data.drop(操作足標, inplace=True)
    data = data.reset_index(drop=True)

    return data


def 貼上去_for_SMT(data, 起始日期, 結束日期):
    起始日期 = datetime.strptime(起始日期, '%Y/%m/%d %H:%M')
    結束日期 = datetime.strptime(結束日期, '%Y/%m/%d %H:%M')
    columns = data.columns

    data = 插入指定線別(data, columns, '2102', 起始日期, 結束日期)
    data = 插入指定線別(data, columns, '2103', 起始日期, 結束日期)
    data = 插入指定線別(data, columns, '2104', 起始日期, 結束日期)
    data = 插入指定線別(data, columns, '2105', 起始日期, 結束日期)
    data = 插入指定線別(data, columns, '2109', 起始日期, 結束日期)

    return data


# data = pd.read_excel('Production schedule 20230511.xls', header=1, sheet_name='SMT')
# 起始日期 = datetime(2023, 6, 2)
# 結束日期 = datetime(2023, 6, 2)
#
# data = 貼上去_for_SMT(data, 起始日期, 結束日期)
#
# data.to_excel('貼上結果.xlsx', index=False)
