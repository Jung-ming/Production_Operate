from datetime import datetime
import numpy as np
import pandas as pd


def 抓取待處理足標(data, 目標線別, 起始日期, 結束日期):
    # return會導致函數立即結束!!
    # 讀取FFFF線別，並將指定日期區間的資料按類型區分
    FFFF線別 = data.query('線別 == "FFFF"')
    操作足標_四零四 = []
    操作足標_重工 = []
    操作足標_coating = []
    操作足標_其他 = []
    for 足標, 欄位 in FFFF線別.iterrows():
        if 起始日期 <= 欄位['預定開工日'].to_pydatetime() <= 結束日期 and '暫不生產' not in str(欄位['工單備註']):
            if '重工' in 欄位['名稱規格']:
                操作足標_重工.append(足標)
            elif 欄位['備註資訊'] == 'coating':
                操作足標_coating.append(足標)
            elif 欄位['SOURCE'] == '四零四':
                操作足標_四零四.append(足標)
            else:
                操作足標_其他.append(足標)
    if 目標線別 == '2209':
        return 操作足標_重工
    elif 目標線別 == '2203':
        return 操作足標_coating
    elif 目標線別 == '2201':
        return 操作足標_四零四
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


def 貼上去_for_DIP(data, 起始日期, 結束日期):
    起始日期 = datetime.strptime(起始日期, '%Y/%m/%d %H:%M')
    結束日期 = datetime.strptime(結束日期, '%Y/%m/%d %H:%M')
    columns = data.columns

    data = 插入指定線別(data, columns, '2201', 起始日期, 結束日期)
    data = 插入指定線別(data, columns, '2202', 起始日期, 結束日期)
    data = 插入指定線別(data, columns, '2203', 起始日期, 結束日期)
    data = 插入指定線別(data, columns, '2209', 起始日期, 結束日期)

    return data
