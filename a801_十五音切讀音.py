#==============================================================================
# 輸入《彙集雅俗通十五音》之【切音（切語上字）】【字韻（切語下字）】，查找漢字及反切標音
#==============================================================================
import os
import sqlite3
import sys

from mod_十五音 import huan_ciat_ca_piau_im, tiau_ho_tng_siann_tiau

if __name__ == "__main__":
    # 確認使用者有輸入反切之切語參數
    if len(sys.argv) != 2:
        print("請輸入欲查詢讀音之【切語】，如：君五求!")
        sys.exit(-1)

    # 取得使用者之輸入：欲查詢讀音的漢字
    切語 = sys.argv[1]

    # 檢查輸入是否為 3 個字元
    if len(切語) != 3:
        print("輸入格式錯誤！請輸入 3 個字元的【字韻】【調號】【切音】，如：君五求!")
        sys.exit(-1)

    # 將切語拆分為 字韻、調號 和 切音
    字韻 = 切語[0]
    調號 = 切語[1]
    切音 = 切語[2]

    # 建立資料庫連線
    connection = sqlite3.connect('雅俗通十五音字典.db')
    cursor = connection.cursor()

    聲調 = tiau_ho_tng_siann_tiau(調號)
    han_ji_piau_im = huan_ciat_ca_piau_im(cursor, 字韻, 聲調, 切音)

    # 檢查是否查找到結果
    if not han_ji_piau_im:
        print("查無結果，請確認輸入是否正確。")
        connection.close()
        sys.exit(-1)

    os.system('cls')
    # 顯示結果
    上字 = han_ji_piau_im[0]['切音']
    下字 = han_ji_piau_im[0]['字韻']
    台語音標 = han_ji_piau_im[0]['漢字標音']
    雅俗通標音 = han_ji_piau_im[0]['雅俗通標音']
    十五音標音 = han_ji_piau_im[0]['十五音標音']
    print('\n=======================================')
    print(f'查詢切音：【{字韻}{調號}{切音}】')
    print(f'台語音標：{台語音標}')
    print(f'十五音切語：{十五音標音}（雅俗通：{雅俗通標音}）')

    # 將所有漢字連接成同一個字串，並使用 "、" 分隔
    漢字列表 = "、".join(record['漢字'] for record in han_ji_piau_im)
    print(f'漢字：{漢字列表}')

    # 關閉資料庫連線
    connection.close()