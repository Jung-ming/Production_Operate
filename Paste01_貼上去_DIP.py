from datetime import datetime
import numpy as np
import pandas as pd


def 抓取待處理足標(data, 目標線別, 起始日期, 結束日期, 線別客戶名單, 抓取四零四=True, 抓取其他客戶=True):
    # return會導致函數立即結束!!
    # 讀取FFFF線別，並將指定日期區間的資料按類型區分
    FFFF線別 = data.query('線別 == "FFFF"')
    操作足標_2201 = []
    操作足標_2202 = []
    操作足標_重工 = []
    操作足標_嘉義 = []
    操作足標_coating = []

    # 根據不同條件抓取對應的足標
    for 足標, 欄位 in FFFF線別.iterrows():
        if 起始日期 <= 欄位['預定開工日'].to_pydatetime() <= 結束日期 \
                and '暫不生產' not in str(欄位['齊料日(Text)'])\
                and 'Forecast' not in str(欄位['備註資訊']):
            if 抓取四零四 and 抓取其他客戶:
                if 欄位['廠區'] == 'CY':
                    操作足標_嘉義.append(足標)
                elif '重工' in 欄位['名稱規格']:
                    操作足標_重工.append(足標)
                elif 欄位['備註資訊'] == 'coating':
                    操作足標_coating.append(足標)
                elif 欄位['SOURCE'] in 線別客戶名單:
                    操作足標_2201.append(足標)
                else:
                    操作足標_2202.append(足標)
            elif 抓取四零四 and 欄位['SOURCE'] == '四零四':
                if 欄位['廠區'] == 'CY':
                    操作足標_嘉義.append(足標)
                elif '重工' in 欄位['名稱規格']:
                    操作足標_重工.append(足標)
                elif 欄位['備註資訊'] == 'coating':
                    操作足標_coating.append(足標)
                elif 欄位['SOURCE'] in 線別客戶名單:
                    操作足標_2201.append(足標)
                else:
                    操作足標_2202.append(足標)
            elif 抓取其他客戶 and 欄位['SOURCE'] != '四零四':
                if 欄位['廠區'] == 'CY':
                    操作足標_嘉義.append(足標)
                elif '重工' in 欄位['名稱規格']:
                    操作足標_重工.append(足標)
                elif 欄位['備註資訊'] == 'coating':
                    操作足標_coating.append(足標)
                elif 欄位['SOURCE'] in 線別客戶名單:
                    操作足標_2201.append(足標)
                else:
                    操作足標_2202.append(足標)

    # 用想操作的線別決定要返回哪個足標
    if 目標線別 == '2209':
        return 操作足標_重工
    elif 目標線別 == '2203':
        return 操作足標_coating
    elif 目標線別 == '2201':
        return 操作足標_2201
    elif 目標線別 == '2205':
        return 操作足標_嘉義
    else:
        return 操作足標_2202


def 插入指定線別(data, columns, 目標線別, 起始日期, 結束日期, 客戶名單, 抓取四零四, 抓取其他客戶):
    # 根據使用者輸入的目標線別，抓取對應的足標
    操作足標 = 抓取待處理足標(data, 目標線別, 起始日期, 結束日期, 客戶名單, 抓取四零四, 抓取其他客戶)
    # 獲取目標線別的資料，包括足標和工序等，方便後續填寫資料時使用
    線別 = data.query(f'線別 == "{目標線別}"')
    目標線別最後足標 = 線別.index[-1]
    工序最末項 = 線別['工序'].iloc[-1]

    # 將操作足標添加到目標線別
    for 足標, 欄位 in data.query(f'index == {操作足標}').iterrows():
        # 設定填入時的資料，包括線別、工序等，
        # 其他就按照原本工單的資訊，不需修改
        欄位['線別'] = 目標線別
        欄位['工序'] = 工序最末項 + 1
        欄位['線別+工序'] = int(欄位['線別'] + str(int(欄位['工序'])).zfill(3))

        # 設定完後將資料插入到目標線別的最後一個，所以位置是最後一個足標再加1
        data = pd.DataFrame(np.insert(data.values, 目標線別最後足標 + 1, values=欄位, axis=0))

        # 插入後因為會多一項，所以最後足標和工序都要+1
        目標線別最後足標 += 1
        工序最末項 += 1

    # 因為插入後會導致欄位名丟失，所以重新賦予欄位名
    data.columns = columns

    # 操作足標插入到其他線別後，要進行刪除，但插入會導致足標無法對應到原本的資料
    # 因次需要再加上足標的總數量，這個總數量就是總共插入多少，也反映足標向下變動的位置
    # 例如原本足標是1，上方插入10筆資料，那原本足標1所對應的資料就是足標11所對應的資料
    操作足標 = [x + len(操作足標) for x in 操作足標]
    data.drop(操作足標, inplace=True)

    # 刪除資料後會導致足標不連貫，故重置足標順序
    data = data.reset_index(drop=True)
    return data


def 生管排序修正(data):
    # 將工單貼上後，重新排序所有線別的生管排序
    基礎值 = 1
    for 足標, 欄位 in data.iterrows():
        if not pd.isnull(欄位['線別']) and 欄位['線別'] != 'FFFF':
            if 2201 <= int(欄位['線別']) <= 2209:
                # 此算式為觀察所得，足標-(線別的尾數+1) = 生管排序
                data.at[足標 , '生管\n排序'] = 足標 + 基礎值
                # print(足標, 欄位['母工單單號'], 欄位['工序'], 欄位['生管\n排序'])
        else:
            基礎值 -= 1

    return data


def 貼上去_for_DIP(data, 起始日期, 結束日期, 客戶名單, 抓取四零四, 抓取其他客戶):
    起始日期 = datetime.strptime(起始日期, '%Y/%m/%d %H:%M')
    結束日期 = datetime.strptime(結束日期, '%Y/%m/%d %H:%M')
    columns = data.columns

    data = 插入指定線別(data, columns, '2201', 起始日期, 結束日期, 客戶名單, 抓取四零四, 抓取其他客戶)
    data = 插入指定線別(data, columns, '2202', 起始日期, 結束日期, 客戶名單, 抓取四零四, 抓取其他客戶)
    data = 插入指定線別(data, columns, '2203', 起始日期, 結束日期, 客戶名單, 抓取四零四, 抓取其他客戶)
    data = 插入指定線別(data, columns, '2205', 起始日期, 結束日期, 客戶名單, 抓取四零四, 抓取其他客戶)
    data = 插入指定線別(data, columns, '2209', 起始日期, 結束日期, 客戶名單, 抓取四零四, 抓取其他客戶)
    data = 生管排序修正(data)

    return data


if __name__ == '__main__':
    data_DIP = pd.read_excel('貼上去測試用 20230511.xls', header=1, sheet_name='DIP')
    print(data_DIP.columns)
    # 抓取四零四 = True
    # 抓取其他客戶 = True
    # 起始日期 = '2023/06/01 00:00'
    # 結束日期 = '2023/06/10 00:00'
    # 線別2201名單 = ['四零四', '益網']
    # data = 貼上去_for_DIP(data_DIP, 起始日期, 結束日期, 線別2201名單, 抓取四零四, 抓取其他客戶)
    # data.to_excel('輸出結果.xlsx', index=False)
