#==============================================================================
# 輸入漢字，查詢《彙集雅俗通十五音》反切標音
#==============================================================================
import os
import sqlite3
import sys

from mod_十五音 import han_ji_ca_piau_im

if __name__ == "__main__":
    # 確認使用者有輸入反切之切語參數
    if len(sys.argv) != 2:
        print("請輸入欲查詢讀音之【漢字】(無需輸入切語之反切上字及下字)!")
        sys.exit(-1)

    # 取得使用者之輸入：欲查詢讀音的漢字
    beh_ca_e_han_ji = sys.argv[1]

    # 建立資料庫連線
    connection = sqlite3.connect('雅俗通十五音字典.db')
    cursor = connection.cursor()

    han_ji_piau_im = han_ji_ca_piau_im(cursor, beh_ca_e_han_ji)

    # 檢查是否查找到結果
    if not han_ji_piau_im:
        print("查找不到漢字：【{beh_ca_e_han_ji}】。")
        connection.close()
        sys.exit(-1)

    os.system('cls')
    for record in han_ji_piau_im:
        上字 = record['切音']
        下字 = record['字韻']
        雅俗通標音 = record['雅俗通標音']
        十五音標音 = record['十五音標音']
        台語音標 = record['漢字標音']

        # 顯示結果
        print('\n=======================================')
        print(f'查詢漢字：【{record["漢字"]}】[{台語音標}]  (字號：{record["識別號"]})')
        print(f'十五音：{十五音標音} （雅俗通：{雅俗通標音}）')
        # print(f'查詢漢字：【{record["漢字"]}】(字號：{record["識別號"]})')
        # print(f'切語：{上字}{下字}切')
        # print(f'台語音標：{台語音標}')
        # print(f'雅俗通切音：{雅俗通標音}')
        # print(f'十五音切音：{十五音標音}')

    # 關閉資料庫連線
    connection.close()