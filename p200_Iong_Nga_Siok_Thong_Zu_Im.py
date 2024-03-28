# coding=utf-8
import re

import xlwings as xw

import p210_ngoo_siok_thong_zu_im as zu_im


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
    
# 將「傳統八聲調」轉換成閩拼使用的調號
tiau_ho_remap_for_BP = {
    1: 1,  # 陰平: 44
    2: 3,  # 上聲：53
    3: 5,  # 陰去：21
    4: 7,  # 上聲：53
    5: 2,  # 陽平：24
    7: 6,  # 陰入：3?
    8: 8,  # 陽入：4?
}

tiau_ho_remap_for_sip_ngoo_im = {
    1: "一",
    2: "二",
    3: "三",
    4: "四",
    5: "五",
    7: "七",
    8: "八",
} 

tiau_ho_remap_for_TPS = {
    1: "",
    2: "ˋ",
    3: "˪",
    4: "",
    5: "ˊ",
    7: "˫",
    8: "\u02D9",
}

TPS_mapping_dict = {
    "p": "ㆴ˙",
    "t": "ㆵ˙",
    "k": "ㆻ˙",
    "h": "ㆷ˙",
}

TPS_remap_dict = {
    "ㄗㄧ": "ㄐㄧ",
    "ㄘㄧ": "ㄑㄧ",
    "ㄙㄧ": "ㄒㄧ",
    "ㆡㄧ": "ㆢㄧ",
}



# ==========================================================
# 設定輸出使用的注音方法
# ==========================================================
tsu_im_huat_list = {
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
}


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
def build_web_page(target_sheet, tsu_im_huat, div_class, rt_tag):
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
        if tsu_im_huat == "SNI":  # 輸出十五音
            siann = siann_bu_dict[siann_bu]["sni"]
            un = un_bu_dict[un_bu]["sni"]
            tiau = tiau_ho_remap_for_sip_ngoo_im[tiau_ho]
            piau_im = f"{un}{tiau}{siann}"
        elif tsu_im_huat == "POJ":  # 輸出白話字拼音
            siann = siann_bu_dict[siann_bu]["poj"]
            un = un_bu_dict[un_bu]["poj"]
            tiau = tiau_ho
            piau_im = f"{siann}{un}{tiau}"
        elif tsu_im_huat == "TL":  # 輸出羅馬拼音
            siann = siann_bu_dict[siann_bu]["tl"]
            un = un_bu_dict[un_bu]["tl"]
            tiau = tiau_ho
            piau_im = f"{siann}{un}{tiau}"
        elif tsu_im_huat == "BP":  # 輸出閩拼拼音
            siann = siann_bu_dict[siann_bu]["bp"]
            un = un_bu_dict[un_bu]["bp"]
            tiau = tiau_ho_remap_for_BP[tiau_ho]
            piau_im = f"{siann}{un}{tiau}"
        elif tsu_im_huat == "TPS":  # 方音符號注音
            siann = siann_bu_dict[siann_bu]["tps"]
            un = un_bu_dict[un_bu]["tps"]
            tiau = tiau_ho_remap_for_TPS[tiau_ho]
            piau_im = f"{siann}{un}{tiau}"

            pattern = r"(ㄗㄧ|ㄘㄧ|ㄙㄧ|ㆡㄧ)"
            searchObj = re.search(pattern, piau_im, re.M | re.I)
            if searchObj:
                key_value = searchObj.group(1)
                piau_im = piau_im.replace(key_value, TPS_remap_dict[key_value])


        # =========================================================
        # 將已注音之漢字加入【漢字注音表】
        # =========================================================
        ruby_tag = f"  <ruby><rb>{han_ji}</rb><rp>(</rp><{rt_tag}>{piau_im}</{rt_tag}><rp>)</rp></ruby>"
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
    for tsu_im_huat in tsu_im_huat_list:
        # 取得 <div> tag 使用的 class 名稱
        div_class = tsu_im_huat_list[tsu_im_huat][0]

        # 取得 Ruby Tag 應使用的 tag 類別( <rt> 或 <rtc>)
        rt_tag = tsu_im_huat_list[tsu_im_huat][1]

        # 取得輸出工作表，使用之名稱
        tsu_im_piau_e_mia = tsu_im_huat_list[tsu_im_huat][2]

        # -----------------------------------------------------
        # 檢查工作表是否已存在
        # -----------------------------------------------------
        get_sheet_ready_to_work(wb, [tsu_im_piau_e_mia])
        beh_tsu_im_e_piau = wb.sheets[tsu_im_piau_e_mia]
        source_sheet_name = beh_tsu_im_e_piau.name

        # -----------------------------------------------------
        # 製作 HTML 網頁
        # -----------------------------------------------------
        # 设置A列的列宽为128
        beh_tsu_im_e_piau.range("A:A").column_width = 128

        # 启用A列单元格的自动换行
        beh_tsu_im_e_piau.range("A:A").api.WrapText = True

        print(f"開始製作【{tsu_im_piau_e_mia}】網頁！")
        build_web_page(beh_tsu_im_e_piau, tsu_im_huat, div_class, rt_tag)
        print(f"【{tsu_im_piau_e_mia}】網頁製作完畢！")

if __name__ == "__main__":
    # 方音符號
    han_ji = "時"
    siann_bu = "s"
    un_bu = "i"
    tiau_ho = 5
    
    siann = siann_bu_dict[siann_bu]["tps"]
    un = un_bu_dict[un_bu]["tps"]
    tiau = tiau_ho_remap_for_TPS[tiau_ho]
    piau_im = f"{siann}{un}{tiau}"

    pattern = r"(ㄗㄧ|ㄘㄧ|ㄙㄧ|ㆡㄧ)"
    searchObj = re.search(pattern, piau_im, re.M | re.I)
    if searchObj:
        key_value = searchObj.group(1)
        piau_im = piau_im.replace(key_value, TPS_remap_dict[key_value])
        print(f"han_ji = {han_ji}")
        print(f"siann_bu = {siann_bu}, un_bu = {un_bu}, tiau_ho = {tiau_ho}")
        print(f"piau_im = {piau_im}")