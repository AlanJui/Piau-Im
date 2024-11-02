#==============================================================================
# 輸入《彙集雅俗通十五音》之【切音（切語上字）】【字韻（切語下字）】，查找漢字及反切標音
#==============================================================================
import os
import sqlite3
import sys

from mod_十五音 import huan_ciat_ca_piau_im, tiau_ho_tng_siann_tiau

if __name__ == "__main__":
    # 確認使用者有輸入反切之切語參數
    if len(sys.argv) != 4:
        print("請輸入欲查詢讀音之【字韻】、【調號】、【切音】，如：君 五 求!")
        sys.exit(-1)

    # 取得使用者之輸入：欲查詢讀音的漢字
    字韻 = sys.argv[1]
    調號 = sys.argv[2]
    切音 = sys.argv[3]

    # 建立資料庫連線
    connection = sqlite3.connect('雅俗通十五音字典.db')
    cursor = connection.cursor()

    聲調 = tiau_ho_tng_siann_tiau(調號)
    han_ji_piau_im = huan_ciat_ca_piau_im(cursor, 字韻, 聲調, 切音)

    os.system('cls')
    for record in han_ji_piau_im:
        上字 = record['切音']
        下字 = record['字韻']
        漢字 = record['漢字']
        雅俗通標音 = record['雅俗通標音']
        十五音標音 = record['十五音標音']
        台語音標 = record['漢字標音']

        # 顯示結果
        print('\n=======================================')
        # print(f'查詢切音：【{字韻}】【{調號}】【{切音}】')
        print(f'查詢切音：【{字韻}{調號}{切音}】')
        print(f'漢字：{漢字} [{台語音標}]')
        # print(f'十五音切音：{十五音標音}  （雅俗通切音：{雅俗通標音}）')
        print(f'十五音：{十五音標音}  （雅俗通：{雅俗通標音}）')

    # 關閉資料庫連線
    connection.close()