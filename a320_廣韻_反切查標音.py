#==============================================================================
# 程式用於查詢反切拼音
#==============================================================================
import os
import sqlite3
import sys

from mod_廣韻 import (
    Cu_Hong_Im_Hu_Ho,
    Kong_Un_Tng_Tai_Loo,
    TL_Tng_Sip_Ngoo_Im,
    ca_siann_bu_piau_im,
    ca_un_bu_piau_im,
    han_ji_ca_piau_im,
)

if __name__ == "__main__":
    # 確認使用者有輸入反切之切語參數
    if len(sys.argv) != 2:
        print("請輸入欲查詢讀音之【切語】(反切上字及下字)!")
        sys.exit(-1)

    ciat_gu = sys.argv[1]

    # 檢查反切拼音是否有兩個字
    if len(ciat_gu) != 2:
        print("反切用的切語，必須有兩個漢字！")
        sys.exit(-1)

    siong_ji, ha_ji = ciat_gu[0], ciat_gu[1]

    # 建立資料庫連線
    connection = sqlite3.connect('Kong_Un.db')
    cursor = connection.cursor()

    os.system('cls')

    切語上字 = han_ji_ca_piau_im(cursor, siong_ji)
    if not 切語上字:
        print(f"切語上字：【{siong_ji}】找不到，無法反切出讀音！")
        sys.exit(-1)

    for record in 切語上字:
        切語上字 = record["漢字"]

        # 顯示結果
        print('\n=======================================')
        print(f'查詢切語：{ciat_gu}')

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
        print(f'切語上字：【{record["漢字"]}】(字號：{record["字號"]})\n')
        if 聲母標音 == 'Ø':  # 若無聲母
            print(f"聲母：{聲母} [{聲母標音}]，國際音標：/{聲母國際音標}/，方音符號：(無聲母)")
        else:
            print(f"聲母：{聲母} [{聲母標音}]，國際音標：/{聲母國際音標}/，方音符號：{聲母方音符號}")
        print(f"(發音部位：{七聲類}，清濁：{清濁}，發送收：{發送收})")

        切語下字 = han_ji_ca_piau_im(cursor, ha_ji)
        if not 切語下字:
            print('\n---------------------------------------')
            print(f"切語下字：【{ha_ji}】找不到，無法反切出讀音！")
            sys.exit(-1)

        for record in 切語下字:
            切語下字 = record["漢字"]

            # 查詢反切下字
            print('\n---------------------------------------')
            print(f'切語下字：【{record["漢字"]}】(字號：{record["字號"]})\n')
            韻母 = record['韻母']
            韻母標音 = record['韻母標音']
            攝 = record['攝']
            調 = record['調']
            韻系 = record['韻系']
            韻系列號 = record['韻系列號']
            目次 = record['目次']
            韻目 = record['韻目']
            等 = record['等']
            呼 = record['呼']
            等呼 = record['等呼']
            韻母其它標音 = ca_un_bu_piau_im(cursor, 韻母標音)
            韻母國際音標 = 韻母其它標音[0]['國際音標']
            韻母方音符號 = 韻母其它標音[0]['方音符號']
            print(f"韻母：{韻母} [{韻母標音}]，國際音標：/{韻母國際音標}/，方音符號：{韻母方音符號}")
            print(f"攝：{攝}，調：{調}聲，目次：{目次}，韻系：{韻系}，韻系列號：{韻系列號}")
            print(f"韻目：{韻目}，{等}等，{呼}口 ({等呼})")

            # 組合拼音
            廣韻調名 = f'{清濁[-1]}{調}'
            台羅聲調 = int(Kong_Un_Tng_Tai_Loo(廣韻調名))
            十五音 = TL_Tng_Sip_Ngoo_Im(聲母標音, 韻母標音, 台羅聲調, cursor)
            十五音切韻 = 十五音['標音']
            print('\n---------------------------------------')
            print('聲調：上字辨【清濁】，下字定【四聲】。\n')
            print(f' (1) 清濁：上字得【{清濁[-1]}】聲；')
            print(f' (2) 四聲：下字得【{調}】聲調；')
            print(f' (3) 由【{清濁[-1]}{調}】聲調，推導【台羅聲調】為：第【{台羅聲調}】調。')
            if 台羅聲調 == 6:
                台羅聲調 = 2
                print(f' (4) 台羅聲調：第【6】調，等同第【{台羅聲調}】調。')

            if 聲母標音 == 'Ø':  # 若無聲母
                聲母標音 = ''
                方音符號標音 = f'{韻母方音符號}{Cu_Hong_Im_Hu_Ho(台羅聲調)}'
            else:
                方音符號標音 = f'{聲母方音符號}{韻母方音符號}{Cu_Hong_Im_Hu_Ho(台羅聲調)}'
            print('\n---------------------------------------')
            print(f'【切語拼音】：\n')
            print(f'台語音標：{聲母標音}{韻母標音}{台羅聲調}')
            print(f'方音符號：{聲母方音符號}{韻母方音符號}{Cu_Hong_Im_Hu_Ho(台羅聲調)}')
            print(f'十五音切韻：{十五音切韻}')

    # 關閉資料庫連線
    connection.close()