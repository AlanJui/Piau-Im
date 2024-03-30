import sys, os 
from mod_Query_for_Tshiat_Gu import query_ji_piau, query_siau_un, query_ciat_gu_siong_ji, query_ciat_gu_ha_ji


def main():
    # 檢查是否有提供足夠的參數
    if len(sys.argv) != 2:
        print("請輸入欲查詢讀音之漢字!")
        return

    # 從命令列參數取得查詢漢字
    han_ji = sys.argv[1]

    # 從反切拼音取得反切上字和反切下字
    try:
        ji_piau = query_ji_piau(han_ji)
    except Exception:
        print("查不到這個漢字！")
        sys.exit()

    # 顯示結果
    os.system('cls')
    print('\n=================================================')
    print(f"欲查詢拼音之漢字：{han_ji}")
    print("\n")
    print(f'字= {ji_piau[0]["字"]}')
    print(f'切語= {ji_piau[0]["小韻切語"]}')
    print(f'拼音= {ji_piau[0]["拼音"]}')
    print(f'字義 = {ji_piau[0]["字義"]}')

    for record in ji_piau:
        ciat_gu = record["小韻切語"]
        siau_un = query_siau_un(ciat_gu)

        # 根據反切上字和反切下字來查詢台羅拼音
        siong_ji = ciat_gu[0]
        ha_ji = ciat_gu[1]
        siann_bu = query_ciat_gu_siong_ji(siong_ji)
        un_bu = query_ciat_gu_ha_ji(ha_ji)

        # 顯示結果
        print('\n=================================================')
        if not siau_un:
            print(f'查不到【{record['小韻切語']}】小韻！')
            print(f'小韻識別號 = {record["小韻識別號"]}')
        else:
            print(f"小韻：{siau_un[0]['切語']} (拼音：{siau_un[0]['聲母拼音碼']}{siau_un[0]['韻母拼音碼']}{siau_un[0]['拼音調號']})")
            print(f'聲母= {siau_un[0]["聲母"]} (清濁= {siau_un[0]["清濁"]})')
            print(f'韻母= {siau_un[0]["韻母"]} (調/韻/等/呼 = {siau_un[0]["調"]} {siau_un[0]["韻"]} {siau_un[0]["等"]} {siau_un[0]["呼"]})')
            print(f'聲母拼音碼= {siau_un[0]["聲母拼音碼"]}')
            print(f'韻母拼音碼= {siau_un[0]["韻母拼音碼"]}')
            print(f'拼音調號= {siau_un[0]["拼音調號"]}')
            print('\n-------------------------------------------------')
            print(f"反切上字：{siong_ji}")
            print(f'聲母 = {siann_bu[0]["聲母"]} (發音部位：{siann_bu[0]['發音部位']}, 清濁: {siann_bu[0]["清濁"]})')
            print(f'(切語上字：{siann_bu[0]["切語上字"]})')
            print('\n-------------------------------------------------')
            print(f"反切下字：{ha_ji}")
            print(f'韻母 = {un_bu[0]["韻母"]} (攝：{un_bu[0]["攝"]}, 調：{un_bu[0]["調"]}, 韻：{un_bu[0]["韻"]}, 等：{un_bu[0]["等"]}, 呼：{un_bu[0]["呼"]})')
            print(f'(切語下字：{un_bu[0]["切語下字"]})')

        # 暫停，避免視窗一閃而過
        print("\n")
        input("按下換行鍵以繼續...")

if __name__ == "__main__":
    main()
