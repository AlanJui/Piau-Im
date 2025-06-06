#==============================================================================
# 輸入漢字，查詢廣韻反切標音
#==============================================================================
import os
import sqlite3
import sys

from mod_廣韻 import (  # Cu_Hong_Im_Hu_Ho,
    Kong_Un_Siann_Tiau_Tng_Tai_Loo,
    ca_siann_bu_piau_im,
    ca_un_bu_piau_im,
    han_ji_ca_piau_im,
)
from mod_標音 import PiauIm

if __name__ == "__main__":
    # 確認使用者有輸入反切之切語參數
    if len(sys.argv) != 2:
        print("請輸入欲查詢讀音之【漢字】(無需輸入切語之反切上字及下字)!")
        sys.exit(-1)

    # 取得使用者之輸入：欲查詢讀音的漢字
    beh_cha_e_han_ji = sys.argv[1]

    # 初始化 PiauIm 類別，産生標音物件
    # piau_im = PiauIm(han_ji_khoo='廣韻')
    piau_im = PiauIm()

    # 建立資料庫連線
    connection = sqlite3.connect('Kong_Un.db')
    cursor = connection.cursor()

    han_ji_piau_im = han_ji_ca_piau_im(cursor, beh_cha_e_han_ji)

    os.system('cls')
    for record in han_ji_piau_im:
        上字 = record['上字']
        下字 = record['下字']

        # 顯示結果
        print('\n=======================================')
        print(f'查詢漢字：【{record["漢字"]}】(字號：{record["字號"]})')
        print(f'切語：{上字}{下字}切')
        print(f'標音：{record["標音"]}')
        # print(f'字義：{record["字義"]}')

        # 查詢反切上字
        聲母 = record['聲母']
        聲母標音 = record['聲母標音']
        七聲類 = record['七聲類']
        清濁 = record['清濁']
        發送收 = record['發送收']
        聲母其它標音 = ca_siann_bu_piau_im(cursor, 聲母標音)
        聲母國際音標 = 聲母其它標音[0]['國際音標']
        聲母方音符號 = 聲母其它標音[0]['方音符號']
        print('\n---------------------------------------')
        print(f'【切語上字】：{上字}\n')
        if 聲母標音 == 'Ø':  # 若無聲母
            print(f"聲母：{聲母} [{聲母標音}]，國際音標：/{聲母國際音標}/，方音符號：(無聲母)")
        else:
            print(f"聲母：{聲母} [{聲母標音}]，國際音標：/{聲母國際音標}/，方音符號：{聲母方音符號}")
        print(f"(發音部位：{七聲類}，清濁：{清濁}，發送收：{發送收})")

        # 查詢反切下字
        韻母 = record['韻母']
        韻母標音 = record['韻母標音']
        攝 = record['攝']
        調 = record['調']
        目次 = record['目次']
        韻目 = record['韻目']
        等 = record['等']
        呼 = record['呼']
        等呼 = record['等呼']
        韻母其它標音 = ca_un_bu_piau_im(cursor, 韻母標音)
        韻母國際音標 = 韻母其它標音[0]['國際音標']
        韻母方音符號 = 韻母其它標音[0]['方音符號']
        print('\n---------------------------------------')
        print(f"【切語下字】 = {下字}\n")
        print(f"韻母：{韻母} [{韻母標音}]，國際音標：/{韻母國際音標}/，方音符號：{韻母方音符號}")
        print(f"攝：{攝}，調：{調}聲，目次：{目次}")
        print(f"韻目：{韻目}，{等}等，{呼}口 ({等呼})")

        # 組合拼音
        # 廣韻調名 = record['廣韻調名']
        # 台羅聲調 = record['台羅聲調']
        廣韻調名 = f'{清濁[-1]}{調}'
        台羅聲調 = Kong_Un_Siann_Tiau_Tng_Tai_Loo(廣韻調名)     # 將廣韻調名轉換成台羅聲調
        十五音 = piau_im.han_ji_piau_im_tng_huan(
            piau_im_huat='十五音',
            siann_bu=聲母標音,
            un_bu=韻母標音,
            tiau_ho=台羅聲調,
        )
        十五音切韻 = 十五音
        print('\n---------------------------------------')
        print('【聲調】：上字辨【清濁】，下字定【四聲】。\n')
        print(f' (1) 清濁：上字得【{清濁[-1]}】聲；')
        print(f' (2) 四聲：下字得【{調}】聲調；')
        print(f' (3) 由【{清濁[-1]}{調}】聲調，推導【台羅聲調】為：第【{台羅聲調}】調。')
        if 台羅聲調 == 6:
            台羅聲調 = 2
            print(f' (4) 台羅聲調：第【6】調，等同第【{台羅聲調}】調。')

        方音符號調符 = piau_im.Cu_Hong_Im_Hu_Ho_Tiau_Hu(台羅聲調)
        if 聲母標音 == 'Ø':  # 若無聲母
            聲母標音 = ''
            方音符號標音 = f'{韻母方音符號}{方音符號調符}'
        else:
            方音符號標音 = f'{聲母方音符號}{韻母方音符號}{方音符號調符}'
        print('\n---------------------------------------')
        print(f'【切語拼音】：\n')
        print(f'台語音標：{聲母標音}{韻母標音}{台羅聲調}')
        print(f'方音符號：{聲母方音符號}{韻母方音符號}{方音符號調符}')
        print(f'十五音切韻：{十五音切韻}')

    # 關閉資料庫連線
    connection.close()