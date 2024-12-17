# Tng_Sing_Bang_Iah.py (轉成網頁)
# 用途：將【漢字注音】工作表中的漢字、台語音標及台語注音符號，轉成 HTML 網頁格式。
import os
import re
import sqlite3

import xlwings as xw

from mod_file_access import get_named_value
from mod_標音 import split_tai_gi_im_piau  # 分解台語音標
from mod_標音 import PiauIm, is_punctuation, split_hong_im_hu_ho

# ==========================================================
# 注音法設定和共用變數
# ==========================================================
zu_im_huat_list = {
    "SNI": ["fifteen_yin", "rt", "十五音切語"],
    "TPS": ["Piau_Im", "rt", "方音符號注音"],
    "POJ": ["pin_yin", "rt", "白話字拼音"],
    "TL": ["pin_yin", "rt", "台羅拼音"],
    "BP": ["pin_yin", "rt", "閩拼標音"],
    "TLPA_Plus": ["pin_yin", "rt", "台羅改良式"],
    "DBL": ["Siang_Pai", "rtc", "雙排注音"],
}


def create_html_file(output_path, content, title='您的標題'):
    template = f"""
<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <title>{title}</title>
    <meta charset="UTF-8">
    <link rel="stylesheet" href="assets/styles/styles.css">
</head>
<body>
    {content}
</body>
</html>
    """

    # Write to HTML file
    with open(output_path, 'w', encoding='utf-8') as file:
        file.write(template)

    # 顯示輸出之網頁檔案及其存放目錄路徑
    print(f"\n輸出網頁檔案：{output_path}")

def put_picture(wb, source_sheet_name):
    html_str = ""

    title = wb.sheets["env"].range("TITLE").value
    # web_page_title = f"《{title}》【{source_sheet_name}】\n"
    web_page_title = f"《{title}》\n"
    image_url = wb.sheets["env"].range("IMAGE_URL").value

    # ruff: noqa: E501
    div_tag = (
        "<div class='separator' style='clear: both'>\n"
        "  <a href='圖片' style='display: block; padding: 1em 0; text-align: center'>\n"
        "    <img alt='%s' border='0' width='400' data-original-height='630' data-original-width='1200'\n"
        "      src='%s' />\n"
        "  </a>\n"
        "</div>\n"
    )
    # 寫入文章附圖
    # html_str = f"《{title}》【{source_sheet_name}】\n"
    html_str = f"{title}\n"
    # html_str += div_tag % (title, image_url)
    html_str += (div_tag % (title, image_url) + "\n")
    return html_str

def choose_piau_im_method(piau_im, zu_im_huat, siann_bu, un_bu, tiau_ho):
    """選擇並執行對應的注音方法"""
    if zu_im_huat == "十五音":
        return piau_im.SNI_piau_im(siann_bu, un_bu, tiau_ho)
    elif zu_im_huat == "白話字":
        return piau_im.POJ_piau_im(siann_bu, un_bu, tiau_ho)
    elif zu_im_huat == "台羅拼音":
        return piau_im.TL_piau_im(siann_bu, un_bu, tiau_ho)
    elif zu_im_huat == "閩拼方案":
        return piau_im.BP_piau_im(siann_bu, un_bu, tiau_ho)
    elif zu_im_huat == "方音符號":
        return piau_im.TPS_piau_im(siann_bu, un_bu, tiau_ho)
    elif zu_im_huat == "台語音標":
        siann = piau_im.Siann_Bu_Dict[siann_bu]["台語音標"] or ""
        un = piau_im.Un_Bu_Dict[un_bu]["台語音標"]
        return f"{siann}{un}{tiau_ho}"
    return ""

def concat_ruby_tag(style, han_ji, tlpa_im_piau, han_ji_piau_im):
    """將漢字、台語音標及台語注音符號，合併成一個 Ruby Tag"""
    if style == "DBL":
        ruby_tag = f"  <ruby><rb>{han_ji}</rb><rt>{tlpa_im_piau}</rt><rp>(</rp><rtc>{han_ji_piau_im}</rtc><rp>)</rp></ruby>"
    elif style == "TPS":
        ruby_tag = f"  <ruby><rb>{han_ji}</rb><rp>(</rp><rtc>{han_ji_piau_im}</rtc><rp>)</rp></ruby>"
    elif style == "SNI":
        ruby_tag = f"  <ruby><rb>{han_ji}</rb><rp>(</rp><rt>{han_ji_piau_im}</rt><rp>)</rp></ruby>"
    else:
        ruby_tag = f"  <ruby><rb>{han_ji}</rb><rp>(</rp><rt>{han_ji_piau_im}</rt><rp>)</rp></ruby>"
    return ruby_tag


# =========================================================
# 依據指定的【注音方法】，輸出含 Ruby Tags 之 HTML 網頁
# =========================================================
def build_web_page(wb, sheet, source_chars, total_length, page_type='含頁頭', piau_im_huat='方音符號', piau_im=None):
    write_buffer = ""

    # =========================================================
    # 輸出放置圖片的 HTML Tag
    # =========================================================
    # 寫入文章附圖
    if page_type == '含頁頭':
        write_buffer += put_picture(wb, sheet.name)
    # =========================================================
    # 輸出 <div> tag
    # =========================================================
    div_class = zu_im_huat_list[Web_Page_Style][0]
    html_str = f"<div class='{div_class}'><p>"
    write_buffer += (html_str + "\n")

    # 每頁最多處理 20 列
    TOTAL_ROWS = int(wb.names['每頁總列數'].refers_to_range.value) # 自名稱為【每頁總列數】之儲存格，取得【每頁最多處理幾列】之值
    # 每列最多處理 15 字元
    CHARS_PER_ROW = int(wb.names['每列總字數'].refers_to_range.value)  # 自名稱為【每列總字數】之儲存格，取得【每列最多處理幾個字元】之值
    # 設定起始及結束的欄位  （【D欄=4】到【R欄=18】）
    start = 4
    end = start + CHARS_PER_ROW

    if total_length and total_length < (CHARS_PER_ROW * TOTAL_ROWS):
        row = 5
        index = 0  # 用來追蹤處理到哪個字元

        # 在 Console 顯示待處理的字串
        print(f"待處理的漢字 = {source_chars}")

        # 逐字處理字串
        while index < total_length:
            # 設定當前作用儲存格，根據 `row` 和 `col` 動態選取
            sheet.range((row, 1)).select()

            for col in range(start, end):  # 【D欄=4】到【R欄=18】
                col_name = xw.utils.col_name(col)
                if index < total_length:
                    ruby_tag = ""
                    src_char = source_chars[index]  # 取得目前欲處理的【漢字】
                    if src_char == "\n":
                        # 若遇到換行字元，退出迴圈
                        write_buffer += ("</p><p>\n")
                        index += 1
                        print("\n")
                        break
                    else:
                        han_ji = sheet.range((row, col)).value  # 取得漢字
                        # 當 han_ji 是標點符號時，不需要注音
                        if is_punctuation(han_ji):
                            ruby_tag = f"<span>{han_ji}</span>\n"
                            # 在 Console 顯示目前處理的漢字，以便使用者可知目前進度
                            print(f"({row}, {col_name}) = {han_ji}")
                        else:
                            # 取得漢字的【台語音標】
                            lo_ma_im_piau = sheet.range((row - 1, col)).value  # 取得漢字的台語音標
                            # 當儲存格寫入之資料為 None 情況時之處理作法：給予空字串
                            lo_ma_im_piau = lo_ma_im_piau if lo_ma_im_piau is not None else ""

                            # zu_im_hu_ho = sheet.range((row + 1, col)).value  # 取得漢字的台語注音符號
                            if piau_im_huat == "台語音標":
                                han_ji_piau_im = lo_ma_im_piau
                            else:
                                # zu_im_list = split_hong_im_hu_ho(lo_ma_im_piau)
                                zu_im_list = split_tai_gi_im_piau(lo_ma_im_piau)
                                if zu_im_list[0] == "" or zu_im_list[0] == None:
                                    siann_bu = "Ø"
                                else:
                                    siann_bu = zu_im_list[0]

                                han_ji_piau_im = choose_piau_im_method(
                                    piau_im,
                                    piau_im_huat,
                                    siann_bu,
                                    zu_im_list[1],
                                    zu_im_list[2]
                                )

                            # 在 Console 顯示目前處理的漢字，以便使用者可知目前進度
                            print(f"({row}, {col_name}) = {han_ji} [{lo_ma_im_piau}] 【{han_ji_piau_im}】")
                            # =========================================================
                            # 將已注音之漢字加入【漢字注音表】
                            # =========================================================
                            # ruby_tag = f"<ruby><rb>{han_ji}</rb><rt>{lo_ma_im_piau}</rt><rtc>{han_ji_piau_im}</rtc></ruby>\n"
                            ruby_tag = concat_ruby_tag(Web_Page_Style, han_ji, lo_ma_im_piau, han_ji_piau_im)

                    write_buffer += ruby_tag
                    index += 1
                else:
                    break  # 若已處理完畢，退出欄位迴圈

            # 每處理一行後，換到下一行
            print("\n")
            row += 4

        # =========================================================
        # 輸出 </div>
        # =========================================================
        html_str = "</p></div>"
        write_buffer += html_str

    # 返回網頁輸出暫存區
    return write_buffer


def tng_sing_bang_iah(wb, sheet_name='漢字注音', cell='V3', page_type='含頁頭'):
    global source_sheet  # 宣告 source_sheet 為全域變數
    global source_sheet_name  # 宣告 source_sheet_name 為全域變數
    global total_length  # 宣告 total_length 為全域變數
    global Web_Page_Style

    # -------------------------------------------------------------------------
    # 連接指定資料庫
    # -------------------------------------------------------------------------
    han_ji_khoo = get_named_value(wb, '漢字庫', '河洛話')
    Web_Page_Style = get_named_value(wb, '網頁格式', 'DBL')
    piau_im = PiauIm(han_ji_khoo)

    # -------------------------------------------------------------------------
    # 選擇指定的工作表
    # -------------------------------------------------------------------------
    sheet = wb.sheets[sheet_name]   # 選擇工作表
    sheet.activate()               # 將「漢字注音」工作表設為作用中工作表
    sheet.range('A1').select()     # 將 A1 儲存格設為作用儲存格
    source_sheet_name = sheet.name

    han_ji_piau_im_huat = wb.names['標音方法'].refers_to_range.value

    # -----------------------------------------------------
    # 產生 HTML 網頁用文字檔
    # -----------------------------------------------------
    title = wb.sheets["env"].range("TITLE").value
    # web_page_title = f"《{title}》【{source_sheet_name}】"
    web_page_title = f"{title}"

    # 確保 output 子目錄存在
    siann_lui = get_named_value(wb, '語音類型', '文讀音')
    output_dir = 'docs'
    # output_file = f"{title}_{siann_lui}.html"
    output_file = f"{title}_{han_ji_piau_im_huat}.html"
    output_path = os.path.join(output_dir, output_file)

    # 開啟文字檔，準備寫入網頁內容
    f = open(output_path, 'w', encoding='utf-8')

    # 取得 V3 儲存格的字串
    source_chars = sheet.range(cell).value
    if source_chars:
        # 計算字串的總長度
        total_length = len(source_chars)

        # ==========================================================
        # 自「漢字注音表」，製作各種注音法之 HTML 網頁
        # ==========================================================
        print(f"開始製作【漢字注音】網頁！")
        html_content = build_web_page(
            wb, sheet, source_chars, total_length, page_type, han_ji_piau_im_huat, piau_im
        )

        # 輸出到網頁檔案
        create_html_file(output_path, html_content, web_page_title)

        print(f"【漢字注音】網頁製作完畢！")
