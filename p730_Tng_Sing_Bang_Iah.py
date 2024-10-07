# Tng_Sing_Bang_Iah.py (轉成網頁)
# 用途：將【漢字注音】工作表中的漢字、台語音標及台語注音符號，轉成 HTML 網頁格式。
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
def build_web_page(wb, sheet, source_chars, total_length):
    write_buffer = ""
    
    # =========================================================
    # 輸出放置圖片的 HTML Tag
    # =========================================================
    # 寫入文章附圖
    write_buffer += put_picture(wb, sheet.name)
    # =========================================================
    # 輸出 <div> tag
    # =========================================================
    div_class = "Siang_Pai"
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
                        break;  
                    else: 
                        han_ji = sheet.range((row, col)).value  # 取得漢字
                        # 當 han_ji 是標點符號時，不需要注音
                        if is_punctuation(han_ji):
                            ruby_tag = f"<span>{han_ji}</span>\n"
                            # 在 Console 顯示目前處理的漢字，以便使用者可知目前進度
                            print(f"({row}, {col_name}) = {han_ji}")
                        else:
                            lo_ma_im_piau = sheet.range((row - 1, col)).value  # 取得漢字的台語音標
                            zu_im_hu_ho = sheet.range((row + 1, col)).value  # 取得漢字的台語注音符號

                            # 處理拼音或注音是 None 的情況
                            lo_ma_im_piau = lo_ma_im_piau if lo_ma_im_piau is not None else ""
                            zu_im_hu_ho = zu_im_hu_ho if zu_im_hu_ho is not None else ""

                            # 在 Console 顯示目前處理的漢字，以便使用者可知目前進度
                            print(f"({row}, {col_name}) = {han_ji} [{lo_ma_im_piau}] 【{zu_im_hu_ho}】")
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
                            ruby_tag = f"<ruby><rb>{han_ji}</rb><rt>{lo_ma_im_piau}</rt><rtc>{zu_im_hu_ho}</rtc></ruby>\n"
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


def tng_sing_bang_iah(wb, sheet_name='漢字注音', cell='V3'):
    global source_sheet  # 宣告 source_sheet 為全域變數
    global source_sheet_name  # 宣告 source_sheet_name 為全域變數
    global total_length  # 宣告 end_of_source_row 為全域變數

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
    source_chars = sheet.range(cell).value
    if source_chars:
        # 計算字串的總長度
        total_length = len(source_chars)

        # ==========================================================
        # 自「漢字注音表」，製作各種注音法之 HTML 網頁
        # ==========================================================
        print(f"開始製作【漢字注音】網頁！")
        html_content = build_web_page(wb, sheet, source_chars, total_length)

        # 輸出到網頁檔案
        create_html_file(output_path, html_content, web_page_title)
        
        print(f"【漢字注音】網頁製作完畢！")
