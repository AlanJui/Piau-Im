# coding=utf-8
import re

import xlwings as xw

import modules.han_ji_chu_im as ji

# ==========================================================
# 設定共用變數
# ==========================================================
end_of_source_row = 0
source_sheet = None

# ==========================================================
# 設定輸出使用的注音方法
# ==========================================================
tsu_im_huat_list = {
    "SNI": [
        "fifteen_yin",  # <div class="">
        "rt",  # Ruby Tag: <rt> / <rtc>
        "十五音注音",  # 輸出工作表名稱
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
        "閩拼拼音",  # 輸出工作表名稱
    ],
}


# =========================================================
# 檢查工作表是否已存在；若否：則建立
# =========================================================
def get_sheet_ready_to_work(wb, sheet_name_list):
    for sheet_name in sheet_name_list:
        sheet = wb.sheets[sheet_name]
        try:
            sheet.select()
            sheet.clear()
            continue
        except Exception as e:
            # 當 Exception 為 CommandError 時，表工作表不存在
            print(e)
            # 新增程式需使用之工作表
            wb.sheets.add(name=sheet_name)


# =========================================================
# 依據指定的【注音方法】，輸出含 Ruby Tags 之 HTML 網頁
# =========================================================
def build_web_page(target_sheet, tsu_im_huat, div_class, rt_tag):
    # =========================================================
    # 輸出 <div> tag
    # =========================================================
    source_index = 1  # index for source sheet
    target_index = 1  # index for target sheet
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
        siann_index = 0
        if siann_bu.strip() != "":
            siann_index = ji.get_siann_idx(siann_bu)
            if siann_index == -1:
                # 記錄沒找到之聲母
                print(f"漢字：【{han_ji}】，找不到【聲母】：{siann_bu}！")

        # 取得韻母之韻母碼
        un_bu = source_sheet.range("D" + str(source_index)).value
        un_index = ji.get_un_idx(un_bu)
        if un_index == -1:
            # 記錄沒找到之韻母
            print(f"漢字：【{han_ji}】，找不到【韻母】：{un_bu}！")

        # 取得調號
        tiau_ho = int(source_sheet.range("E" + str(source_index)).value)

        # =========================================================
        # 將漢字的「注音碼」，依指定的〖注音法〗，轉換為注音／拼音
        # =========================================================
        if tsu_im_huat == "SNI":  # 輸出十五音
            piau_im = ji.get_sip_ngoo_im_chu_im(siann_index, un_index, tiau_ho)
        elif tsu_im_huat == "TPS":  # 方音符號注音
            piau_im = ji.get_TPS_chu_im(siann_index, un_index, tiau_ho)
        elif tsu_im_huat == "POJ":  # 輸出白話字拼音
            piau_im = ji.get_POJ_chu_im(siann_index, un_index, tiau_ho)
        elif tsu_im_huat == "TL":  # 輸出羅馬拼音
            piau_im = ji.get_TL_chu_im(siann_index, un_index, tiau_ho)
        elif tsu_im_huat == "BP":  # 輸出閩拼拼音
            piau_im = ji.get_BP_chu_im(siann_index, un_index, tiau_ho)

        # =========================================================
        # 將已注音之漢字加入【漢字注音表】
        # =========================================================
        ruby_tag = f"  <ruby><rb>{han_ji}</rb><rp>(</rp><{rt_tag}>{piau_im}</{rt_tag}><rp>)</rp></ruby>"
        target_sheet.range("A" + str(target_index)).value = ruby_tag

        # =========================================================
        # 調整讀取來源；寫入標的各工作表
        # =========================================================
        target_index += 1
        source_index += 1

    # =========================================================
    # 輸出 </div>
    # =========================================================
    html_str = "</p></div>"
    target_sheet.range("A" + str(target_index)).value = html_str


def main_run(CONVERT_FILE_NAME):
    global source_sheet  # 宣告 source_sheet 為全域變數
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
        zhu_im_piau_mia = tsu_im_huat_list[tsu_im_huat][2]

        # -----------------------------------------------------
        # 檢查工作表是否已存在
        # -----------------------------------------------------
        get_sheet_ready_to_work(wb, [zhu_im_piau_mia])
        beh_tsu_im_e_piau = wb.sheets[zhu_im_piau_mia]

        # -----------------------------------------------------
        # 製作 HTML 網頁
        # -----------------------------------------------------
        build_web_page(beh_tsu_im_e_piau, tsu_im_huat, div_class, rt_tag)
