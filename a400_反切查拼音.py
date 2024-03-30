import sys
from mod_Query_for_Tshiat_Gu import query_ji_piau, query_siau_un, query_ciat_gu_siong_ji, query_ciat_gu_ha_ji

def yong_ciat_gu_za_piau_im(han_ji, ciat_gu):
    # 檢查反切拼音是否有兩個字
    if len(ciat_gu) != 2:
        print("反切拼音必須是兩個字")
        return

    return query_ji_piau(han_ji)
    

def main():
    # 檢查是否有提供足夠的參數
    if len(sys.argv) != 3:
        print("請提供兩個參數：查詢漢字和反切拼音")
        return

    # 從命令列參數取得查詢漢字和反切拼音
    han_ji = sys.argv[1]
    ciat_gu = sys.argv[2]

    # 從反切拼音取得反切上字和反切下字
    han_ji_piau_im = yong_ciat_gu_za_piau_im(han_ji, ciat_gu)
    ciat_gu_siong_ji = ciat_gu[0]
    ciat_gu_ha_ji = ciat_gu[1]

    # 根據反切上字和反切下字來查詢台羅拼音
    siann_bu = query_ciat_gu_siong_ji(ciat_gu_siong_ji)
    un_bu = query_ciat_gu_ha_ji(ciat_gu_ha_ji)

    # 印出結果
    print(f"欲查詢拼音之漢字：{han_ji}")
    print(f"反切拼音為：{han_ji_piau_im}")
    print(f"反切上字為：{siann_bu}")
    print(f"反切下字為：{un_bu}")

if __name__ == "__main__":
    main()
