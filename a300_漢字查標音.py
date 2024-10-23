#==============================================================================
# 輸入漢字，查詢廣韻反切標音
#==============================================================================
import os
import sys

from mod_廣韻 import (
    ca_siann_bu_piau_im,
    ca_un_bu_piau_im,
    connect_to_db2,
    han_ji_ca_piau_im,
)


def main():
    # 確認使用者有輸入反切之切語參數
    if len(sys.argv) != 2:
        print("請輸入欲查詢讀音之【漢字】(無需輸入切語之反切上字及下字)!")
        sys.exit(-1)

    # 取得使用者之輸入：欲查詢讀音的漢字
    beh_cha_e_han_ji = sys.argv[1]

    # 連上 DB
    with connect_to_db2('.\\Kong_Un.db') as conn:
        cursor = conn.cursor()

        # 查漢字之切語
        han_ji_piau_im = han_ji_ca_piau_im(cursor, beh_cha_e_han_ji)
        if not han_ji_piau_im:
            print("找不到此【漢字】!")
            sys.exit(-1)
        
        上字 = han_ji_piau_im[0]['上字']    
        下字 = han_ji_piau_im[0]['下字']    

        # 顯示結果
        os.system('cls')
        print('\n=================================================')
        print(f'查詢漢字：【{beh_cha_e_han_ji}】')
        print(f'切語：{上字}{下字}切')
        print(f'標音：{han_ji_piau_im[0]["標音"]}')
        # print(f'字義：{han_ji_piau_im[0]["字義"]}')

        # 查詢反切上字
        print('\n-------------------------------------------------')
        print('【切語上字】：\n')
        聲母 = han_ji_piau_im[0]['聲母']
        聲母標音 = han_ji_piau_im[0]['聲母標音']
        七聲類 = han_ji_piau_im[0]['七聲類']
        清濁 = han_ji_piau_im[0]['清濁']
        發送收 = han_ji_piau_im[0]['發送收']
        聲母其它標音 = ca_siann_bu_piau_im(cursor, 聲母標音)
        聲母國際音標 = 聲母其它標音[0]['國際音標聲母']
        聲母方音符號 = 聲母其它標音[0]['方音聲母']
        print(f"切語上字 = {上字}")
        print(f"聲母：{聲母} [{聲母標音}]，國際音標：/{聲母國際音標}/，方音符號：{聲母方音符號}")
        print(f"(發音部位：{七聲類}，清濁：{清濁}，發送收：{發送收})")

        # 查詢反切下字
        print('\n-------------------------------------------------')
        print('【切語下字】：\n')
        韻母 = han_ji_piau_im[0]['韻母']
        韻母標音 = han_ji_piau_im[0]['韻母標音']
        攝 = han_ji_piau_im[0]['攝']
        調 = han_ji_piau_im[0]['調']
        目次 = han_ji_piau_im[0]['目次']
        韻目 = han_ji_piau_im[0]['韻目']
        等 = han_ji_piau_im[0]['等']
        呼 = han_ji_piau_im[0]['呼']
        等呼 = han_ji_piau_im[0]['等呼']
        韻母其它標音 = ca_un_bu_piau_im(cursor, 韻母標音)
        韻母國際音標 = 韻母其它標音[0]['國際音標韻母']
        韻母方音符號 = 韻母其它標音[0]['方音韻母']
        print(f"切語下字 = {下字}") 
        print(f"韻母：{韻母} [{韻母標音}]，國際音標：/{韻母國際音標}/，方音符號：{韻母方音符號}")
        print(f"攝：{攝}，調：{調}聲，目次：{目次}")
        print(f"韻目：{韻目}，{等}等，{呼}口 ({等呼})")

        # 組合拼音
        廣韻調名 = han_ji_piau_im[0]['廣韻調名']
        台羅聲調 = han_ji_piau_im[0]['台羅聲調']
        print('\n-------------------------------------------------')
        print('【聲調】：\n')
        print(f'上字取聲之【清濁】，下字取調之【陰陽】，得：{廣韻調名}調；')
        print(f'推導【台羅聲調】得：第 {台羅聲調} 調。')
        print('\n-------------------------------------------------')
        print(f'【切語拼音】：\n')
        print(f'台語音標：{聲母標音}{韻母標音}{台羅聲調}')
        print(f'方音符號：{聲母方音符號}{韻母方音符號}{台羅聲調}\n')


if __name__ == "__main__":
    main()