# -*- coding: utf-8 -*-
import re

import xlwings as xw

import p210_ngoo_siok_thong_zu_im as zu_im


# ==========================================================
# 設定輸出使用的注音方法
# ==========================================================
zu_im_huat_list = {
    "SNI": [
        "fifteen_yin",  # <div class="">
        "rt",  # Ruby Tag: <rt> / <rtc>
        "十五音切語",  # 輸出工作表名稱
    ],
    "TPS": [
        "zhu_yin",  # <div class="">
        "rtc",  # Ruby Tag: <rt> / <rtc>
        "方音符號注音",  # 輸出工作表名稱
    ],
    "POJ": [
        "pin_yin",  # <div class="">
        "rt",  # Ruby Tag: <rt> / <rtc>
        "白話字拼音",  # 輸出工作表名稱
    ],
    "TL": [
        "pin_yin",  # <div class="">
        "rt",  # Ruby Tag: <rt> / <rtc>
        "台羅拼音",  # 輸出工作表名稱
    ],
    "BP": [
        "pin_yin",  # <div class="">
        "rt",  # Ruby Tag: <rt> / <rtc>
        "閩拼標音",  # 輸出工作表名稱
    ],
    "TLPA_Plus": [
        "pin_yin",  # <div class="">
        "rt",  # Ruby Tag: <rt> / <rtc>
        "台羅改良式",  # 輸出工作表名稱
    ],
    "DBL": [
        "zhu_yin",  # <div class="">
        "rtc",  # Ruby Tag: <rt> / <rtc>
        "雙排注音",  # 輸出工作表名稱
    ],
}

# ==========================================================
# 設定共用變數
# ==========================================================
wb = None
end_of_source_row = 0
source_sheet = None
source_sheet_name = ""

# 使用 SQLite 資料庫，設定聲母及韻母之注音對照表
try:
    siann_bu_dict = zu_im.init_siann_bu_dict()
    un_bu_dict = zu_im.init_un_bu_dict()
except Exception as e:
    print(e)
    
#================================================================
# 雅俗通十五音(SNI:Nga-Siok-Thong)
#================================================================
def SNI_piau_im(siann_bu, un_bu, tiau_ho):
    siann = siann_bu_dict[siann_bu]["sni"]
    un = un_bu_dict[un_bu]["sni"]
    tiau = tiau_ho_remap_for_sip_ngoo_im[tiau_ho]
    piau_im = f"{un}{tiau}{siann}"
    return piau_im

tiau_ho_remap_for_sip_ngoo_im = {
    1: "一",
    2: "二",
    3: "三",
    4: "四",
    5: "五",
    7: "七",
    8: "八",
} 

#================================================================
# 方音符號注音（TPS）
# TPS_mapping_dict = {
#     "p": "ㆴ˙",
#     "t": "ㆵ˙",
#     "k": "ㆻ˙",
#     "h": "ㆷ˙",
# }
#================================================================
def TPS_piau_im(siann_bu, un_bu, tiau_ho):
    tiau_ho_remap_for_TPS = {
        1: "",
        2: "ˋ",
        3: "˪",
        4: "",
        5: "ˊ",
        7: "˫",
        8: "\u02D9",
    }

    TPS_piau_im_remap_dict = {
        "ㄗㄧ": "ㄐㄧ",
        "ㄘㄧ": "ㄑㄧ",
        "ㄙㄧ": "ㄒㄧ",
        "ㆡㄧ": "ㆢㄧ",
    }

    siann = siann_bu_dict[siann_bu]["tps"]
    un = un_bu_dict[un_bu]["tps"]
    tiau = tiau_ho_remap_for_TPS[tiau_ho]
    piau_im = f"{siann}{un}{tiau}"

    pattern = r"(ㄗㄧ|ㄘㄧ|ㄙㄧ|ㆡㄧ)"
    searchObj = re.search(pattern, piau_im, re.M | re.I)
    if searchObj:
        key_value = searchObj.group(1)
        piau_im = piau_im.replace(key_value, TPS_piau_im_remap_dict[key_value])

    return piau_im

#================================================================
# 在韻母加調號：白話字(POJ)與台羅(TL)同
#================================================================
def un_bu_ga_tiau_ho(guan_im, tiau):
    tiau_hu_dict = {
        1: "",
        2: "\u0301",
        3: "\u0300",
        4: "",
        5: "\u0302",
        6: "\u030C",
        7: "\u0304",
        8: "\u030D",
        9: "\u030B",
    }
    guan_im_u_ga_tiau_ho = f"{guan_im}{tiau_hu_dict[tiau]}"
    return guan_im_u_ga_tiau_ho

#================================================================
# 白話字（POJ）
# 順序：《o＞e＞a＞u＞i＞ng＞m》；而 ng 標示在字母 n 上。
# 例外：
#  - oai、oan、oat、oah 標在 a 上。
#  - oeh 標在 e 上。
#================================================================
def POJ_piau_im(siann_bu, un_bu, tiau_ho):
    siann = siann_bu_dict[siann_bu]["poj"]
    un = un_bu_dict[un_bu]["poj"]
    piau_im = f"{siann}{un}"
    
    # 韻母為複元音
    # pattern1 = r"(oai|oan|oah|oeh|ee|ei)"
    pattern1 = r"(oai|oan|oah|oeh)"
    searchObj = re.search(pattern1, piau_im, re.M | re.I)
    if searchObj:
        found = searchObj.group(1)
        un_chars = list(found)
        idx = 0
        if found == "ee" or found == "ei":
            idx = 0
        else:
            # found = oai/oan/oah/oeh
            idx = 1
        guan_im = un_chars[idx]
        un_chars[idx] = un_bu_ga_tiau_ho(guan_im, tiau_ho)
        un_str = "".join(un_chars)
        piau_im = piau_im.replace(found, un_str)
    else:
        # 韻母為單元音或鼻音韻
        pattern2 = r"(o|e|a|u|i|ng|m)"
        searchObj2 = re.search(pattern2, piau_im, re.M | re.I)
        if searchObj2:
            found = searchObj2.group(1)
            guan_im = found
            new_un = un_bu_ga_tiau_ho(guan_im, tiau_ho)
            piau_im = piau_im.replace(found, new_un)

    return piau_im

#================================================================
# 台羅拼音（TL）
# 順序：《o＞e＞a＞u＞i＞ng＞m》；而 ng 標示在字母 n 上。
#================================================================
def TL_piau_im(siann_bu, un_bu, tiau_ho):
    siann = siann_bu_dict[siann_bu]["tl"]
    un = un_bu_dict[un_bu]["tl"]
    piau_im = f"{siann}{un}"

    # 韻母為複元音
    pattern1 = r"(uai|uan|uah|ueh|ee|ei)"
    searchObj = re.search(pattern1, piau_im, re.M | re.I)
    if searchObj:
        found = searchObj.group(1)
        un_chars = list(found)
        idx = 0
        if found == "ee" or found == "ei":
            idx = 0
        else:
            # found = uai/uan/uah/ueh
            idx = 1
        guan_im = un_chars[idx]
        un_chars[idx] = un_bu_ga_tiau_ho(guan_im, tiau_ho)
        un_str = "".join(un_chars)
        piau_im = piau_im.replace(found, un_str)
    else:
        # 韻母為單元音或鼻音韻
        pattern2 = r"(o|e|a|u|i|ng|m)"
        searchObj2 = re.search(pattern2, piau_im, re.M | re.I)
        if searchObj2:
            found = searchObj2.group(1)
            guan_im = found
            new_un = un_bu_ga_tiau_ho(guan_im, tiau_ho)
            piau_im = piau_im.replace(found, new_un)

        # pattern = r"(oo|ee|ei|o|e|a|u|i|ng|m)"
        # searchObj = re.search(pattern, piau_im, re.M | re.I)
        # if searchObj:
        #     found = searchObj.group(1)
        #     un_chars = list(found)
        #     guan_im = un_chars[0]
        #     un_chars[0] = un_bu_ga_tiau_ho(guan_im, tiau_ho)
        #     un_str = "".join(un_chars)
        #     piau_im = piau_im.replace(found, un_str)

    return piau_im
    
#================================================================
# 閩拼（BP）
#
# 【調號標示規則】
# 當一個音節有多個字母時，調號得標示在響度最大的字母上面（通常在韻腹）。由規則可以判定確切的字母：
# 
#  - 響度優先順序： a > oo > (e = o) > (i = u)〈低元音 > 高元音 > 無擦通音 > 擦音 > 塞音〉
#  - 二合字母 iu 及 ui ，調號都標在後一個字母上；因為前一個字母是介音。
#  - m 作韻腹時則標於字母 m 上。
#  - 二合字母 oo 及 ng，標於前一個字母上；比如 ng 標示在字母 n 上。
#  - 三合字母 ere，標於最後的字母 e 上。
#================================================================

# 將「台羅八聲調」轉換成閩拼使用的調號
tiau_ho_remap_for_BP = {
    1: 1,  # 陰平: 44
    2: 3,  # 上聲：53
    3: 5,  # 陰去：21
    4: 7,  # 上聲：53
    5: 2,  # 陽平：24
    7: 6,  # 陰入：3?
    8: 8,  # 陽入：4?
}

def bp_un_bu_ga_tiau_ho(guan_im, tiau):
    tiau_hu_dict = {
        1: "\u0304",  # 陰平
        2: "\u0341",  # 陽平
        3: "\u030C",  # 上声
        5: "\u0300",  # 陰去
        6: "\u0302",  # 陽去
        7: "\u0304",  # 陰入
        8: "\u0341",  # 陽入
    }
    return f"{guan_im}{tiau_hu_dict[tiau]}"

def BP_piau_im(siann_bu, un_bu, tiau_ho):
    siann = siann_bu_dict[siann_bu]["bp"]
    un = un_bu_dict[un_bu]["bp"]
    piau_im = f"{siann}{un}"

    # 當聲母為「空白」，韻母為：i 或 u 時，調整聲母
    un_chars = list(un)
    if siann == "":
        if un_chars[0] == "i":
            siann = "y"
        elif un_chars[0] == "u":
            siann = "w"

    pattern = r"(a|oo|ere|iu|ui|ng|e|o|i|u|m)"
    searchObj = re.search(pattern, piau_im, re.M | re.I)

    if searchObj:
        found = searchObj.group(1)
        un_chars = list(found)
        idx = 0
        if found == "iu" or found == "ui":
            idx = 1
        elif found == "oo" or found == "ng":
            idx = 0
        elif found == "ere":
            idx = 2

        # 處理韻母加聲調符號
        guan_im = un_chars[idx]
        tiau = tiau_ho_remap_for_BP[tiau_ho]  # 將「傳統八聲調」轉換成閩拼使用的調號
        un_chars[idx] = bp_un_bu_ga_tiau_ho(guan_im, tiau)
        un_str = "".join(un_chars)
        piau_im = piau_im.replace(found, un_str)

    return piau_im

# =========================================================
# 檢查工作表是否已存在；若否：則建立
# =========================================================
def get_sheet_ready_to_work(wb, sheet_name_list):
    for sheet_name in sheet_name_list:
        sheets =  [sheet.name for sheet in wb.sheets]  # 獲取所有工作表的名稱
        if sheet_name in sheets:
            sheet = wb.sheets[sheet_name]
            try:
                sheet.select()
                sheet.clear()
                continue
            except Exception as e:
                print(e)
        else:
            # CommandError 的 Exception 發生時，表工作表不存在
            # 新增程式需使用之工作表
            print(f"工作表【{sheet_name}】已新增！")
            wb.sheets.add(name=sheet_name)

# =========================================================
# 依據指定的【注音方法】，輸出含 Ruby Tags 之 HTML 網頁
# =========================================================
def build_web_page(target_sheet, zu_im_huat, div_class, rt_tag):
    source_index = 1  # index for source sheet
    target_index = 1  # index for target sheet

    # =========================================================
    # 輸出放置圖片的 HTML Tag
    # =========================================================
    title = wb.sheets["env"].range("TITLE").value
    image_url = wb.sheets["env"].range("IMAGE_URL").value

    # ruff: noqa: E501
    div_tag = (
        "《%s》【%s】\n"
        "<div class='separator' style='clear: both'>\n"
        "  <a href='圖片' style='display: block; padding: 1em 0; text-align: center'>\n"
        "    <img alt='%s' border='0' width='400' data-original-height='630' data-original-width='1200'\n"
        "      src='%s' />\n"
        "  </a>\n"
        "</div>\n"
    )
    html_str = ""
    html_str += div_tag % (title, source_sheet_name, title, image_url)
    target_sheet.range("A" + str(target_index)).value = html_str
    target_index += 1

    # =========================================================
    # 輸出 <div> tag
    # =========================================================
    html_str = f"<div class='{div_class}'><p>"
    target_sheet.range("A" + str(target_index)).value = html_str
    target_index += 1

    while source_index <= end_of_source_row:
        # 自 source_sheet 取待注音漢字
        han_ji = str(source_sheet.range("A" + str(source_index)).value)
        han_ji.strip()

        # =========================================================
        # 如是空白或換行，處理換行
        # =========================================================
        if han_ji == "" or han_ji == "\n":
            html_str = "</p><p>"
            target_sheet.range("A" + str(target_index)).value = html_str
            target_index += 1
            source_index += 1
            continue

        # =========================================================
        # 如只是標點符號，不必處理為漢字注音的工作
        # =========================================================
        # 比對是否為標點符號
        piau_tiam = r"[；：？！\uFF0C\uFF08-\uFF09\u2013-\u2014\u2026\\u2018-\u201D\u3000\u3001-\u303F]"
        tshue_tioh_piau_tiam = re.search(piau_tiam, han_ji, re.M | re.I)

        # 若是標點符號，則直接寫入目標工作表
        if tshue_tioh_piau_tiam:
            # 將取到的「標點符號」，寫入目標工作表
            target_sheet.range("A" + str(target_index)).value = han_ji
            target_index += 1
            source_index += 1
            continue

        # =========================================================
        # 在字庫中查不到注音的漢字，略過注音處理
        # =========================================================
        piau_im = ""
        ji_e_piau_im = str(source_sheet.range("B" + str(source_index)).value).strip()

        if ji_e_piau_im == "None":
            # 讀到空白儲存格，視為使用者：「欲終止一個段落」；故於目標工作表寫入一個「換行」字元。
            ruby_tag = f"  <ruby><rb>{han_ji}</rb><rp>(</rp><{rt_tag}>{piau_im}</{rt_tag}><rp>)</rp></ruby>"
            target_sheet.range("A" + str(target_index)).value = ruby_tag
            target_index += 1
            source_index += 1
            continue

        # =========================================================
        # 備妥注音時需參考用的資料
        # =========================================================
        # 取得聲母之聲母碼
        siann_bu = source_sheet.range("C" + str(source_index)).value
        if not siann_bu_dict[siann_bu]:  
            # 記錄沒找到之聲母
            print(f"漢字：【{han_ji}】，找不到【聲母】：{siann_bu}！")

        # 取得韻母之韻母碼
        un_bu = source_sheet.range("D" + str(source_index)).value
        if not un_bu_dict[un_bu]:
            # 記錄沒找到之韻母
            print(f"漢字：【{han_ji}】，找不到【韻母】：{un_bu}！")

        # 取得調號
        tiau_ho = int(source_sheet.range("E" + str(source_index)).value)

        # =========================================================
        # 將漢字的「注音碼」，依指定的〖注音法〗，轉換為注音／拼音
        # =========================================================
        if zu_im_huat == "SNI":  # 輸出十五音
            piau_im = SNI_piau_im(siann_bu, un_bu, tiau_ho)
        elif zu_im_huat == "POJ":  # 輸出白話字拼音
            piau_im = POJ_piau_im(siann_bu, un_bu, tiau_ho)
        elif zu_im_huat == "TL":  # 輸出羅馬拼音
            piau_im = TL_piau_im(siann_bu, un_bu, tiau_ho)
        elif zu_im_huat == "BP":  # 輸出閩拼拼音
            piau_im = BP_piau_im(siann_bu, un_bu, tiau_ho)
        elif zu_im_huat == "TPS":  # 方音符號注音
            piau_im = TPS_piau_im(siann_bu, un_bu, tiau_ho)
        elif zu_im_huat == "TLPA_Plus":  # 台羅改良式
            siann = siann_bu_dict[siann_bu]["code"]
            # 若是空聲母，則不輸出聲母
            siann = "" if siann == "q" else siann
            un = un_bu_dict[un_bu]["code"]
            tiau = tiau_ho
            piau_im = f"{siann}{un}{tiau}"
        else:
            piau_im = TPS_piau_im(siann_bu, un_bu, tiau_ho)
            siann = siann_bu_dict[siann_bu]["code"]
            # 若是空聲母，則不輸出聲母
            siann = "" if siann == "q" else siann
            un = un_bu_dict[un_bu]["code"]
            tiau = tiau_ho
            piau_im2 = f"{siann}{un}{tiau}"

        # =========================================================
        # 將已注音之漢字加入【漢字注音表】
        # =========================================================
        if zu_im_huat != "DBL":
            ruby_tag = f"  <ruby><rb>{han_ji}</rb><rp>(</rp><{rt_tag}>{piau_im}</{rt_tag}><rp>)</rp></ruby>"
        else:
            ruby_tag = f"  <ruby><rb>{han_ji}</rb><rp>(</rp><{rt_tag}>{piau_im}</{rt_tag}><rp>)</rp><rt>{piau_im2}</rt></ruby>"

        target_sheet.range("A" + str(target_index)).value = ruby_tag

        # =========================================================
        # 調整讀取來源；寫入標的各工作表
        # =========================================================
        print(f"row = {source_index}，漢字：{han_ji}，注音：[{piau_im}]")
        target_index += 1
        source_index += 1

    # =========================================================
    # 輸出 </div>
    # =========================================================
    html_str = "</p></div>"
    target_sheet.range("A" + str(target_index)).value = html_str


def main_run(CONVERT_FILE_NAME):
    global wb  # 宣告 wb 為全域變數
    global source_sheet  # 宣告 source_sheet 為全域變數
    global source_sheet_name  # 宣告 source_sheet_name 為全域變數
    global end_of_source_row  # 宣告 end_of_source_row 為全域變數

    # ==========================================================
    # 打開 Excel 檔案
    # ==========================================================
    file_path = CONVERT_FILE_NAME
    wb = xw.Book(file_path)

    source_sheet = wb.sheets["漢字注音表"]
    end_of_source_row = (
        source_sheet.range("A" + str(source_sheet.cells.last_cell.row)).end("up").row
    )
    print(f"End of Row = {end_of_source_row}")

    # ==========================================================
    # 自「漢字注音表」，製作各種注音法之 HTML 網頁
    # ==========================================================
    for zu_im_huat in zu_im_huat_list:
        # 取得 <div> tag 使用的 class 名稱
        div_class = zu_im_huat_list[zu_im_huat][0]

        # 取得 Ruby Tag 應使用的 tag 類別( <rt> 或 <rtc>)
        rt_tag = zu_im_huat_list[zu_im_huat][1]

        # 取得輸出工作表，使用之名稱
        zu_im_piau_e_mia = zu_im_huat_list[zu_im_huat][2]

        # -----------------------------------------------------
        # 檢查工作表是否已存在
        # -----------------------------------------------------
        get_sheet_ready_to_work(wb, [zu_im_piau_e_mia])
        beh_zu_im_e_piau = wb.sheets[zu_im_piau_e_mia]
        source_sheet_name = beh_zu_im_e_piau.name

        # -----------------------------------------------------
        # 製作 HTML 網頁
        # -----------------------------------------------------
        # 设置A列的列宽为128
        beh_zu_im_e_piau.range("A:A").column_width = 128

        # 启用A列单元格的自动换行
        beh_zu_im_e_piau.range("A:A").api.WrapText = True

        print(f"開始製作【{zu_im_piau_e_mia}】網頁！")
        build_web_page(beh_zu_im_e_piau, zu_im_huat, div_class, rt_tag)
        print(f"【{zu_im_piau_e_mia}】網頁製作完畢！")

if __name__ == "__main__":
    #=====================================================================
    # 方音符號
    #=====================================================================
    han_ji = "時"
    siann_bu = "s"
    un_bu = "i"
    tiau_ho = 5
    
    piau_im = TPS_piau_im(siann_bu, un_bu, tiau_ho) 
    print("測試方音符號注音：")
    print(f"han_ji = {han_ji}")
    print(f"siann_bu = {siann_bu}, un_bu = {un_bu}, tiau_ho = {tiau_ho}")
    print(f"piau_im = {piau_im}")
    assert piau_im == "ㄒㄧˊ", "測試失敗!"

    #=====================================================================
    # 雅俗通十五
    #=====================================================================
    piau_im = SNI_piau_im(siann_bu, un_bu, tiau_ho) 
    print("\n測試雅俗通十五音注音：")
    print(f"han_ji = {han_ji}")
    print(f"siann_bu = {siann_bu}, un_bu = {un_bu}, tiau_ho = {tiau_ho}")
    print(f"piau_im = {piau_im}")
    assert piau_im == "居五時", "測試失敗!"

    #=====================================================================
    # 閩拼
    #=====================================================================
    han_ji = "字"
    siann_bu = "j"
    un_bu = "i"
    tiau_ho = 7

    piau_im = BP_piau_im(siann_bu, un_bu, tiau_ho) 
    print("\n測試閩拼注音：")
    print(f"han_ji = {han_ji}")
    print(f"siann_bu = {siann_bu}, un_bu = {un_bu}, tiau_ho = {tiau_ho}")
    print(f"piau_im = {piau_im}")
    assert piau_im == "zzî", "測試失敗!"

    #=====================================================================
    # 白話字
    #=====================================================================
    han_ji = "轉"
    siann_bu = "z"
    un_bu = "uan"
    tiau_ho = 2

    piau_im = POJ_piau_im(siann_bu, un_bu, tiau_ho) 
    print("\n測試白話字注音：")
    print(f"han_ji = {han_ji}")
    print(f"siann_bu = {siann_bu}, un_bu = {un_bu}, tiau_ho = {tiau_ho}")
    print(f"piau_im = {piau_im}")
    # assert piau_im == "choán", "測試失敗!"
    assert piau_im == "choán", "測試失敗!"

    #=====================================================================
    # 台羅拼音
    #=====================================================================
    han_ji = "轉"
    siann_bu = "z"
    un_bu = "uan"
    tiau_ho = 2

    piau_im = TL_piau_im(siann_bu, un_bu, tiau_ho) 
    print("\n測試白話字注音：")
    print(f"han_ji = {han_ji}")
    print(f"siann_bu = {siann_bu}, un_bu = {un_bu}, tiau_ho = {tiau_ho}")
    print(f"piau_im = {piau_im}")
    # assert piau_im == "tsuán", "測試失敗!"