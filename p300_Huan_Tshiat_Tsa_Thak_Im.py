# ```bash
# py a300_反切查拼音.py [參數1：查詢漢字] [參數2：反切拼]
# ```
# 
# 參數：
# 
# 1. 查詢漢字： 1 個中文字
# 2. 反切拼音： 2 個中文字
#    2.1 反切上字：反切拼音參數的第 1 個中文字
#    2.2 反切下字：反切拼音參數的第 2 個中文字
# 
# 案例：
# 
# ```bash
# py a300_反切查拼音.py 東 德紅
# ```
# 漢字= "東"
# 
# 上字= "德" --> 台羅拚音：tik4  --> 聲母 = "t"   --> 調號 = 4 --> 清音
# 下字= "紅" --> 台羅拚音：hong5 --> 聲母 = "ong" --> 調號 = 5 --> 平聲
# 由清音+平聲 --> 調號 = 5
# 
# 台羅拼音 = t + ong + 5 = tong5

def main():
    # 取得輸入
    han_ji = "東"
    huan_tshiat = "德紅"

    siong_ji = huan_tshiat[0]
    e_ji = huan_tshiat[1]

    # 分析輸入
    print("=========================================")
    print(f"欲查漢字：{han_ji}")
    print(f"反切讀音：{huan_tshiat}\n")
    print(f"反切上字：{siong_ji}")
    print(f"反切下字：{e_ji}\n")

    # ## 查字典取拼音
    # 
    # 在【漢字典】，查詢反切上字、下字之標音。
    # 
    # - 上字：德 --> tik4
    # - 下字：紅 --> hong5
    siong_ji_piau_im = {
        "han_ji": "德",
        "piau_im": "tik4",
        "sian_bu": "t",
        "un_bu": "ik",
        "tiau_ho": 4,
    }

    e_ji_piau_im = {
        "han_ji": "紅",
        "piau_im": "hong5",
        "sian_bu": "h",
        "un_bu": "ong",
        "tiau_ho": 5,
    }

    # ## 分析上字：上取聲母分清濁
    # 
    # - (1) 上字定聲理：雙聲取聲母，上一字祗取發聲
    # - (2) 上字分清濁：依據字韻聲調辦清/濁：聲調小於5為清(陰)、否則為濁(陽)
    # 

    # 0:清 1:濁
    TSHING = 0 
    LO = 1

    piau_im_sian_bu = siong_ji_piau_im["sian_bu"]
    piau_im_tshing_lo = TSHING if siong_ji_piau_im["tiau_ho"] < 5 else LO 

    print("-----------------------------------------")
    print(f"上字：{siong_ji}\n")
    print(f"上字定聲理，因上字標音為：{siong_ji_piau_im['piau_im']}")
    print(f"故得聲母：{piau_im_sian_bu}\n")
    print(f"上字分清濁，因聲調為：{siong_ji_piau_im['tiau_ho']}")
    print(f"故清濁為：{'清' if piau_im_tshing_lo == TSHING else '濁'}\n")


    # ## 分析下字：下取韻母定開合
    # 
    # - (3) 下字定韻律：疊韻取韻母，下一字祗取其收韻
    # - (4) 下字定開合：依據韻母聲調辦四聲（平/上/去/入）：平[1,5]、上[2,6]、去[3,7]、入[4,8]

    piau_im_un_bu = e_ji_piau_im["sian_bu"]
    piau_im_tshing_lo = TSHING if siong_ji_piau_im["tiau_ho"] < 5 else LO 

    print("-----------------------------------------")
    print(f"下字：{e_ji}\n")
    print(f"下字定韻律，因下字標音為：{e_ji_piau_im['piau_im']}")
    print(f"故得韻母：{piau_im_un_bu}\n")
    print(f"下字定開合，因聲調為：{e_ji_piau_im['tiau_ho']}")
    print(f"故清濁為：{'清' if piau_im_tshing_lo == TSHING else '濁'}\n")


    # ## 切出四聲八調之聲調
    # 
    # - 依據 (2) 清濁 、(4) 開合，定四聲八調：
    #     - 上字為清聲，下字為濁聲，切成之字仍為清聲；
    #     - 下一字為合口，上一字為開口，切成之字仍為合口。
    # | **清濁聲** | ** 平聲韻** | ** 上聲韻** | ** 去聲韻** | ** 入聲韻** |
    # |:-------:|:--------:|:--------:|:--------:|:--------:|
    # | **清**   | 1        | 2        | 3        | 4        |
    # | **濁**   | 5        | 6        | 7        | 8        |
    # 

    sian_tiau = 1

    # ### 輸出反切拼音 
    # 
    # 依據輸入的 2 個參數進行處理，完成後輸出如下格式資料：
    # 
    # `格式`：
    # 
    # ```bash
    # 欲查詢拼音之漢字：[參數1：查詢漢字]
    # 
    # 反切拼音為：[參數2：反切拼]
    # 
    # 反切上字為：[反切上字]
    # 反切下字為：[反切下字]
    # ```
    # 
    # `舉例`：
    # 
    # ```bash
    # 欲查詢拼音之漢字：東
    # 反切拼音為：tong1
    # 
    # 反切上字為：德 (tik4)
    #  - 聲母：t
    #  - 清濁：清音
    # 
    # 反切下字為：紅（hong5）
    #  - 韻母：ong
    #  - 四聲：平聲
    # 
    # 清音配平聲，得聲調為：1
    # ```
    print("N/A")


if __name__ == "__main__":
    main()

