#==============================================================================
# 輸入漢字，查詢廣韻反切標音
#==============================================================================
import os
import sys

from mod_廣韻 import (
    cha_ciat_gu_ha_ji,
    cha_ciat_gu_siong_ji,
    connect_to_db2,
    han_ji_cha_piau_im,
)


def main():
    # 確認使用者有輸入反切之切語參數
    if len(sys.argv) != 2:
        print("請輸入欲查詢讀音之【漢字】(無需輸入切語之反切上字及下字)!")
        sys.exit(-1)

    beh_cha_e_han_ji = sys.argv[1]

    # 連上 DB
    with connect_to_db2('.\\Kong_Un_V2.db') as conn:
        cursor = conn.cursor()

        # 查漢字之切語
        han_ji_piau_im = han_ji_cha_piau_im(cursor, beh_cha_e_han_ji)
        if not han_ji_piau_im:
            print("找不到此【漢字】!")
            sys.exit(-1)
        
        ciat_gu = han_ji_piau_im[0]['切語']    
        siong_ji = ciat_gu[0]
        ha_ji = ciat_gu[1]

        # 顯示結果
        os.system('cls')
        print('\n=================================================')
        print(f'查詢漢字：【{beh_cha_e_han_ji}】')
        print(f'切語：{ciat_gu}')
        print(f'標音：{han_ji_piau_im[0]["漢字標音"]}')
        print(f'字義：{han_ji_piau_im[0]["字義"]}')

        # 顯示廣韻參考資料
        print('\n-------------------------------------------------')
        print('【廣韻查找資料】：\n')
        print(f'目次：{han_ji_piau_im[0]["目次"]}，小韻字序號：{han_ji_piau_im[0]["小韻字序號"]}，小韻字：{han_ji_piau_im[0]["小韻字"]}')

        # 查詢反切上字
        print('\n-------------------------------------------------')
        print('【切語上字】：\n')
        siong_ji_im = han_ji_cha_piau_im(cursor, siong_ji)
        siong_ji_piau = cha_ciat_gu_siong_ji(cursor, siong_ji)
        if not siong_ji_piau:
            print(f'查不到【反切上字】：{siong_ji}')
        else:
            print(f"切語上字 = {siong_ji} (標音：{siong_ji_im[0]['漢字標音']})")
            print(f"聲母：{siong_ji_piau[0]['聲母']} [{siong_ji_piau[0]['聲母拼音碼']}]，國際音標：/{siong_ji_im[0]['聲母國際音標']}/ ")
            print(f"(發音部位：{siong_ji_piau[0]['發音部位']}，清濁：{siong_ji_piau[0]['清濁']}，發送收：{siong_ji_piau[0]['發送收']})")

        # 查詢反切下字
        print('\n-------------------------------------------------')
        print('【切語下字】：\n')
        ha_ji_im = han_ji_cha_piau_im(cursor, ha_ji)
        ha_ji_piau = cha_ciat_gu_ha_ji(cursor, ha_ji)
        if not ha_ji_piau:
            print(f'查不到【反切下字】：{ha_ji}')
        else:
            print(f"切語下字 = {ha_ji} (標音：{ha_ji_im[0]['漢字標音']})")
            print(f"韻母：{ha_ji_piau[0]['韻母']} [{ha_ji_piau[0]['韻母拼音碼']}]，國際音標：/{ha_ji_im[0]['韻母國際音標']}/")
            print(f"攝：{ha_ji_piau[0]['攝']}，調：{ha_ji_piau[0]['調']}聲，目次：{ha_ji_piau[0]['目次']}")
            print(f"{ha_ji_piau[0]['韻系']}韻系，{ha_ji_piau[0]['韻目']}韻，{ha_ji_piau[0]['呼']}口呼，{ha_ji_piau[0]['等']}等 ({ha_ji_piau[0]['等呼']})")

        # 組合拼音
        tiau_ho_list = {
            '清平': 1,
            '清上': 2,
            '清去': 3,
            '清入': 4,
            '濁平': 5,
            '濁上': 2,
            '濁去': 7,
            '濁入': 8,
        }
        siann = siong_ji_piau[0]['聲母拼音碼']
        cing_tok_str = siong_ji_piau[0]['清濁']
        cing_tok = cing_tok_str[-1]
        un = ha_ji_piau[0]['韻母拼音碼']
        tiau_ho = tiau_ho_list[ f"{cing_tok}{ha_ji_piau[0]['調']}" ]

        print('\n-------------------------------------------------')
        print(f'【切語拼音】：{ciat_gu} [{siann}{un}{tiau_ho}]\n')


if __name__ == "__main__":
    main()