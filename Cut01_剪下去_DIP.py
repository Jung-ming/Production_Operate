from datetime import datetime
import numpy as np
import pandas as pd
from Cut04_選擇日期 import *


def 更新線別和工序(row, 新線別, 工序最末項, data, 線別_9201_最後足標):
    row['線別'] = 新線別
    row['生管\n排序'] = 'OK'
    row['工序'] = 工序最末項 + 1
    row['線別+工序'] = int(row['線別'] + str(int(row['工序'])).zfill(3))
    data = pd.DataFrame(np.insert(data.values, 線別_9201_最後足標 + 1, values=row, axis=0))

    return data


def 剪下去_for_DIP(data):
    目標日期 = 選擇日期()
    print('選中時間', datetime.strptime(目標日期, '%Y/%m/%d %H:%M').replace(hour=8, minute=30))
    判斷日期 = datetime.strptime(目標日期, '%Y/%m/%d %H:%M').replace(hour=8, minute=30)

    # 給後面重新賦予列名用
    columns = data.columns

    線別_9201 = data.query('線別 == "9201"')
    # [-1]可以直接選擇最後一個足標和取得工序最末項
    # 取得這兩者的用意在於，從9201線別的最後一項開始插入，並且工序也是從最末項開始遞增
    線別_9201_最後足標 = 線別_9201.index[-1]

    工序最末項 = 線別_9201['工序'].iloc[-1]
    待刪除足標_2201 = []

    # 將符合條件的項目從2201減到9201
    for 足標, 欄位 in data.iterrows():
        if 欄位['線別'] == '2201' and 欄位['結束時間'].to_pydatetime() <= 判斷日期:
            待刪除足標_2201.append(足標)
            data = 更新線別和工序(欄位, '9201', 工序最末項, data, 線別_9201_最後足標)
            工序最末項 += 1
            線別_9201_最後足標 += 1

    # 剪下去後，將不用的項目刪掉
    # 放到這裡做是因為，如果前面就先刪會導致足標亂掉，讓最後的結果出問題
    data.drop(待刪除足標_2201, inplace=True)

    # 由於insert會導致列名丟失，所以重新賦予列名，後面才能繼續操作
    data.columns = columns

    # 刪掉後要再把2201的工序和線別+工序更新位正確的數字
    for 足標, 欄位 in data.iterrows():
        if 欄位['線別'] == '2201':
            # 記住欄位是不能直接賦值的，前面是因為有insert，可以把改動的欄位寫入，這裡則需要另外用其他方式寫入
            # 這裡的邏輯是在於 待刪除足標_2201 的總數 就是被剪下去的數量
            # 所以原來工序只要減去總數，就代表是正確的工序 比如原來工序為10，前面9個被剪掉，那工序10-9，就是新的正確工序(1)
            工序 = int(欄位['工序'] - len(待刪除足標_2201))
            data.at[足標, '工序'] = 工序
            # 這裡將其轉換成 int() 寫入格式會自動靠右，比較好看
            # str(工序).zfill(3) 當字串不滿3個，從左側自動填0
            data.at[足標, '線別+工序'] = int(欄位['線別'] + str(工序).zfill(3))

    # drop會把舊索引列去掉，這樣才能保持列數相同，否則會多一列
    data = data.reset_index(drop=True)

    線別_9202 = data.query('線別 == "9202"')
    線別_9202_最後足標 = 線別_9202.index[-1]
    工序最末項 = 線別_9202['工序'].iloc[-1]
    待刪除足標_2202 = []

    for 足標, 欄位 in data.iterrows():
        if 欄位['線別'] == '2202' and 欄位['結束時間'].to_pydatetime() <= 判斷日期:
            # print(欄位['母工單單號'], 足標, 欄位['結束時間'], 判斷日期)
            待刪除足標_2202.append(足標)
            data = 更新線別和工序(欄位, '9202', 工序最末項, data, 線別_9202_最後足標)
            工序最末項 += 1
            線別_9202_最後足標 += 1

    data.columns = columns

    data.drop(待刪除足標_2202, inplace=True)

    for 足標, 欄位 in data.iterrows():
        if 欄位['線別'] == '2202':
            工序 = int(欄位['工序'] - len(待刪除足標_2202))
            data.at[足標, '工序'] = 工序
            data.at[足標, '線別+工序'] = int(欄位['線別'] + str(工序).zfill(3))

    return data
