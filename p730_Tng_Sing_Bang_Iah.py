# Tng_Sing_Bang_Iah.py (轉成網頁)
# 用途：將【漢字注音】工作表中的漢字、台語音標及台語注音符號，轉成 HTML 網頁格式。
import math
import os

import xlwings as xw


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
    <div>
        <p>
            為能正確顯示「注音符號」，請點擊以下連結，下載注音符號專用字型：
            <a href="https://github.com/cmex-30/Bopomofo_on_Web/tree/master/font/BopomofoRuby1909-v1-Regular.ttf">
                BopomofoRuby1909-v1-Regular.ttf
            </a>
            ，並於使用之電腦端安裝此字型。
        </p>
    </div>
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
    web_page_title = f"《{title}》【{source_sheet_name}】\n"
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
    html_str = f"《{title}》【{source_sheet_name}】\n"
    # html_str += div_tag % (title, image_url)
    html_str += (div_tag % (title, image_url) + "\n")
    return html_str 


# =========================================================
# 判斷是否為標點符號的輔助函數
# =========================================================
def is_punctuation(char):
    # 如果 char 是 None，直接返回 False
    if char is None:
        return False
    
    # 可以根據需要擴充此列表以判斷各種標點符號
    punctuation_marks = "，。！？；：、（）「」『』《》……"
    return char in punctuation_marks


# =========================================================
# 依據指定的【注音方法】，輸出含 Ruby Tags 之 HTML 網頁
# =========================================================
def build_web_page(wb, sheet, total_length):
    write_buffer = ""
    
    # =========================================================
    # 輸出放置圖片的 HTML Tag
    # =========================================================
    # 寫入文章附圖
    write_buffer += put_picture(wb, sheet.name)
    # =========================================================
    # 輸出 <div> tag
    # =========================================================
    div_class = "Sing_Pai"
    html_str = f"<div class='{div_class}'><p>"
    write_buffer += (html_str + "\n")

    # 每列最多處理 15 個字元
    chars_per_row = 15

    if total_length:
        row = 5
        index = 0  # 用來追蹤處理到哪個字元

        # 逐字處理字串
        while index < total_length:
            for col in range(4, 19):  # 【D欄=4】到【R欄=18】
                if index < total_length:
                    han_ji = sheet.range((row, col)).value  # 取得漢字

                    # 當 han_ji 是標點符號時，不需要注音
                    if is_punctuation(han_ji):
                        ruby_tag = f"<span>{han_ji}</span>"
                    else:
                        lo_ma_im_piau = sheet.range((row - 1, col)).value  # 取得漢字的台語音標
                        zu_im_hu_ho = sheet.range((row + 1, col)).value  # 取得漢字的台語注音符號

                        # 處理拼音或注音是 None 的情況
                        lo_ma_im_piau = lo_ma_im_piau if lo_ma_im_piau is not None else ""
                        zu_im_hu_ho = zu_im_hu_ho if zu_im_hu_ho is not None else ""

                        # =========================================================
                        # 將已注音之漢字加入【漢字注音表】
                        # =========================================================
                        # ruby_tag = f"""  
                        # <ruby>
                        #     <rb>{han_ji}</rb>
                        #     <rt>{lo_ma_im_piau}</rt>
                        #     <rp>(</rp>
                        #         <rtc>{zu_im_hu_ho}</rtc>
                        #     <rp>)</rp>
                        # </ruby>
                        # """
                        ruby_tag = f"<ruby><rb>{han_ji}</rb><rt>{lo_ma_im_piau}</rt><rtc>{zu_im_hu_ho}</rtc></ruby>"

                    write_buffer += (ruby_tag + "\n")
                    index += 1
                else:
                    break  # 若已處理完畢，退出欄位迴圈

            # 每處理一行後，換到下一行
            write_buffer += "<br>\n"
            row += 4

        # =========================================================
        # 輸出 </div>
        # =========================================================
        html_str = "</p></div>"
        write_buffer += html_str        

    # 返回網頁輸出暫存區
    return write_buffer


def tng_sing_bang_iah(file_name, sheet_name='漢字注音', cell='V3'):
    global source_sheet  # 宣告 source_sheet 為全域變數
    global source_sheet_name  # 宣告 source_sheet_name 為全域變數
    global total_length  # 宣告 end_of_source_row 為全域變數

   # 打開 Excel 檔案
    wb = xw.Book(file_name)

    # 選擇指定的工作表
    sheet = wb.sheets[sheet_name]
    source_sheet_name = sheet.name

    # -----------------------------------------------------
    # 產生 HTML 網頁用文字檔
    # -----------------------------------------------------
    title = wb.sheets["env"].range("TITLE").value
    web_page_title = f"《{title}》【{source_sheet_name}】"

    # 確保 output 子目錄存在
    output_dir = 'docs'
    output_file = f"{title}_{source_sheet_name}.html"
    output_path = os.path.join(output_dir, output_file)

    # 開啟文字檔，準備寫入網頁內容
    f = open(output_path, 'w', encoding='utf-8')

    # 取得 V3 儲存格的字串
    v3_value = sheet.range(cell).value
    if v3_value:
        # 計算字串的總長度
        total_length = len(v3_value)

        # ==========================================================
        # 自「漢字注音表」，製作各種注音法之 HTML 網頁
        # ==========================================================
        print(f"開始製作【漢字注音】網頁！")
        html_content = build_web_page(wb, sheet, total_length)

        # 輸出到網頁檔案
        create_html_file(output_path, html_content, web_page_title)
        
        print(f"【漢字注音】網頁製作完畢！")