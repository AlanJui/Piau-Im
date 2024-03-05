import sys
import p200_Huan_Tshiat_Tsa_Han_Ji_Thak_Im as tsa_huan_thsiat

def main():
    # 檢查是否有提供足夠的參數
    if len(sys.argv) != 3:
        print("請提供兩個參數：查詢漢字和反切拼音")
        return

    # 從命令列參數取得查詢漢字和反切拼音
    han_ji = sys.argv[1]
    fan_qie = sys.argv[2]

    # 檢查反切拼音是否有兩個字
    if len(fan_qie) != 2:
        print("反切拼音必須是兩個字")
        return

    # 從反切拼音取得反切上字和反切下字
    fan_qie_shang = fan_qie[0]
    fan_qie_xia = fan_qie[1]

    # TODO: 根據反切上字和反切下字來查詢台羅拼音
    tsa_huan_thsiat.main(fan_qie_shang, fan_qie_xia)

    # 印出結果
    print(f"欲查詢拼音之漢字：{han_ji}")
    print(f"反切拼音為：{fan_qie}")
    print(f"反切上字為：{fan_qie_shang}")
    print(f"反切下字為：{fan_qie_xia}")

if __name__ == "__main__":
    main()
