from datetime import datetime
import numpy as np
import pandas as pd


def 抓取待處理足標(data, 目標線別, 起始日期, 結束日期, 線別客戶名單, 抓取四零四=True, 抓取其他客戶=True):
    # return會導致函數立即結束!!
    # 讀取FFFF線別，並將指定日期區間的資料按類型區分
    FFFF線別 = data.query('線別 == "FFFF"')
    操作足標_2101 = []
    操作足標_2102 = []
    操作足標_2103 = []
    操作足標_2104 = []
    操作足標_2105 = []
    操作足標_嘉義 = []
    操作足標_重工 = []
    客戶2101名單 = 線別客戶名單[0]
    客戶2102名單 = 線別客戶名單[1]
    客戶2103名單 = 線別客戶名單[2]
    客戶2105名單 = 線別客戶名單[3]

    for 足標, 欄位 in FFFF線別.iterrows():
        if 起始日期 <= 欄位['預定開工日'].to_pydatetime() <= 結束日期 and '暫不生產' not in str(欄位['工單備註']):
            if 抓取四零四 and 抓取其他客戶:
                if 欄位['廠區'] == 'CY':
                    操作足標_嘉義.append(足標)
                elif '重工' in 欄位['名稱規格']:
                    操作足標_重工.append(足標)
                elif 欄位['SOURCE'] in 客戶2101名單:
                    操作足標_2101.append(足標)
                elif 欄位['SOURCE'] in 客戶2102名單:
                    操作足標_2102.append(足標)
                elif 欄位['SOURCE'] in 客戶2103名單:
                    操作足標_2103.append(足標)
                elif 欄位['SOURCE'] in 客戶2105名單:
                    操作足標_2105.append(足標)
                else:
                    操作足標_2104.append(足標)
            elif 抓取四零四 and 欄位['SOURCE'] == '四零四':
                if 欄位['廠區'] == 'CY':
                    操作足標_嘉義.append(足標)
                elif '重工' in 欄位['名稱規格']:
                    操作足標_重工.append(足標)
                elif 欄位['SOURCE'] in 客戶2101名單:
                    操作足標_2101.append(足標)
                elif 欄位['SOURCE'] in 客戶2102名單:
                    操作足標_2102.append(足標)
                elif 欄位['SOURCE'] in 客戶2103名單:
                    操作足標_2103.append(足標)
                else:
                    操作足標_2104.append(足標)
            elif 抓取其他客戶 and 欄位['SOURCE'] != '四零四':
                if 欄位['廠區'] == 'CY':
                    操作足標_嘉義.append(足標)
                elif '重工' in 欄位['名稱規格']:
                    操作足標_重工.append(足標)
                elif 欄位['SOURCE'] in 客戶2101名單:
                    操作足標_2101.append(足標)
                elif 欄位['SOURCE'] in 客戶2102名單:
                    操作足標_2102.append(足標)
                elif 欄位['SOURCE'] in 客戶2103名單:
                    操作足標_2103.append(足標)
                else:
                    操作足標_2104.append(足標)

    if 目標線別 == '2109':
        return 操作足標_重工
    elif 目標線別 == '2101':
        return 操作足標_2101
    elif 目標線別 == '2102':
        return 操作足標_2102
    elif 目標線別 == '2103':
        return 操作足標_2103
    elif 目標線別 == '2105':
        return 操作足標_2105
    elif 目標線別 == '2106':
        return 操作足標_嘉義
    else:
        return 操作足標_2104


def 插入指定線別(data, columns, 目標線別, 起始日期, 結束日期, 線別客戶名單, 抓取四零四, 抓取其他客戶):
    操作足標 = 抓取待處理足標(data, 目標線別, 起始日期, 結束日期, 線別客戶名單, 抓取四零四, 抓取其他客戶)
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


def 生管排序修正(data):
    基礎值 = 1
    for 足標, 欄位 in data.iterrows():
        if not pd.isnull(欄位['線別']) and 欄位['線別'] != 'FFFF':
            if 2101 <= int(欄位['線別']) <= 2109:
                data.at[足標, '生管\n排序'] = 足標 + 基礎值
        else:
            基礎值 -= 1

    return data


def 貼上去_for_SMT(data, 起始日期, 結束日期, 線別客戶名單, 抓取四零四, 抓取其他客戶):
    起始日期 = datetime.strptime(起始日期, '%Y/%m/%d %H:%M')
    結束日期 = datetime.strptime(結束日期, '%Y/%m/%d %H:%M')
    columns = data.columns

    data = 插入指定線別(data, columns, '2101', 起始日期, 結束日期, 線別客戶名單, 抓取四零四, 抓取其他客戶)
    data = 插入指定線別(data, columns, '2102', 起始日期, 結束日期, 線別客戶名單, 抓取四零四, 抓取其他客戶)
    data = 插入指定線別(data, columns, '2103', 起始日期, 結束日期, 線別客戶名單, 抓取四零四, 抓取其他客戶)
    data = 插入指定線別(data, columns, '2104', 起始日期, 結束日期, 線別客戶名單, 抓取四零四, 抓取其他客戶)
    data = 插入指定線別(data, columns, '2105', 起始日期, 結束日期, 線別客戶名單, 抓取四零四, 抓取其他客戶)
    data = 插入指定線別(data, columns, '2106', 起始日期, 結束日期, 線別客戶名單, 抓取四零四, 抓取其他客戶)
    data = 插入指定線別(data, columns, '2109', 起始日期, 結束日期, 線別客戶名單, 抓取四零四, 抓取其他客戶)
    data = 生管排序修正(data)

    return data
