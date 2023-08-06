import numpy as np
import pandas as pd
from Cut04_選擇日期 import *


def 目標項目線別下移(data, 初始線別, 目標線別, 判斷日期):
    data_目標線別 = data.query(f'線別 == "{目標線別}"')
    # [-1]可以直接選擇最後一個足標和取得工序最末項
    # 取得這兩者的用意在於，從9201(目標線別)的最後一項開始插入，並且工序也是從最末項開始遞增

    # 設定當目標縣別為空時，替代的目標線別
    if 目標線別 == '9201':
        操作用目標線別 = 'FFFF'
    else:
        操作用目標線別 = str(int(目標線別) - 1)

    # 判斷目標線別是否為空，是否需要採用替代的目標線別
    if data_目標線別.empty:
        if 操作用目標線別 == 'FFFF':
            data_目標線別 = data.query(f'線別 == "{操作用目標線別}"')
            目標線別最後足標 = data_目標線別.index[1] - 2
        else:
            data_目標線別 = data.query(f'線別 == "{操作用目標線別}"')
            目標線別最後足標 = data_目標線別.index[-1]
        工序最末項 = 0
    else:
        目標線別最後足標 = data_目標線別.index[-1]
        工序最末項 = data_目標線別['工序'].iloc[-1]

    # 標記移下去的項目的足標，後面要根據這些足標把資料刪掉
    待刪除足標 = []

    # 將符合條件的項目從2201(初始線別)剪到9201(目標線別)
    for 足標, 欄位 in data.iterrows():
        if 欄位['線別'] == 初始線別 and 欄位['開始時間'].to_pydatetime() < 判斷日期:
            待刪除足標.append(足標)
            data = 更新線別和工序_9開頭(欄位, 目標線別, 工序最末項, data, 目標線別最後足標)
            工序最末項 += 1
            目標線別最後足標 += 1

    return data, 待刪除足標


def 更新線別和工序_9開頭(row, 新線別, 工序最末項, data, 目標線別最後足標):
    # 這裡的 (目標線別最後足標 + 1) 不會直接改變變數，而插入的線別本來就是最後足標再加1
    # 代表插在最後足標下面，故邏輯沒錯
    row['線別'] = 新線別
    row['生管\n排序'] = 'OK'
    row['工序'] = 工序最末項 + 1
    row['線別+工序'] = int(row['線別'] + str(int(row['工序'])).zfill(3))
    data = pd.DataFrame(np.insert(data.values, 目標線別最後足標 + 1, values=row, axis=0))

    return data


def 更新線別和工序_2開頭(data, 線別, 已刪除足標):
    for 足標, 欄位 in data.iterrows():
        if 欄位['線別'] == 線別:
            # 記住欄位是不能直接賦值的，前面是因為有insert，可以把改動的欄位寫入，這裡則需要另外用其他方式寫入
            # 這裡的邏輯是在於 待刪除足標_2201 的總數 就是被剪下去的數量
            # 所以原來工序只要減去總數，就代表是正確的工序 比如原來工序為10，前面9個被剪掉，那工序10-9，就是新的正確工序(1)
            工序 = int(欄位['工序'] - len(已刪除足標))
            data.at[足標, '工序'] = 工序
            # 這裡將其轉換成 int() 寫入格式會自動靠右，比較好看
            # str(工序).zfill(3) 當字串不滿3個，從左側自動填0
            data.at[足標, '線別+工序'] = int(欄位['線別'] + str(工序).zfill(3))


def 全部剪下流程(data, 初始線別, 目標線別, columns, 判斷日期):
    data, 待刪除足標 = 目標項目線別下移(data, 初始線別=初始線別, 目標線別=目標線別, 判斷日期=判斷日期)

    # 剪下去後，將不用的項目刪掉
    # 放到這裡做是因為，如果前面就先刪會導致足標亂掉，讓最後的結果出問題

    data.drop(待刪除足標, inplace=True)
    # 由於insert會導致列名丟失，所以重新賦予列名，後面才能繼續操作
    data.columns = columns

    更新線別和工序_2開頭(data, 初始線別, 待刪除足標)

    # drop會把舊索引列去掉，這樣才能保持列數相同，否則會多一列
    data = data.reset_index(drop=True)

    return data


def 剪下去_for_DIP(data, 目標日期):
    DIP判斷日期 = datetime.strptime(目標日期, '%Y/%m/%d %H:%M').replace(hour=8, minute=30)
    DIP判斷日期_2205 = datetime.strptime(目標日期, '%Y/%m/%d %H:%M').replace(hour=8, minute=00)

    # 給後面重新賦予列名用
    columns = data.columns

    data = 全部剪下流程(data, '2201', '9201', columns, DIP判斷日期)

    data = 全部剪下流程(data, '2202', '9202', columns, DIP判斷日期)

    data = 全部剪下流程(data, '2205', '9205', columns, DIP判斷日期_2205)

    return data
