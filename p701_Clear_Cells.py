# 填漢字等標音：將整段的文字拆解，個別填入儲存格，以便後續人工手動填入台語音標、注音符號。
import xlwings as xw


def clear_han_ji_kap_piau_im(wb, sheet_name='漢字注音'):
    sheet = wb.sheets[sheet_name]   # 選擇工作表
    sheet.activate()               # 將「漢字注音」工作表設為作用中工作表
    sheet.range('A1').select()     # 將 A1 儲存格設為作用儲存格

    total_rows = wb.names['每頁總列數'].refers_to_range.value
    cells_per_row = 4
    end_of_rows = int((total_rows * cells_per_row ) + 2)
    cells_range = f'D3:R{end_of_rows}'

    sheet.range(cells_range).clear_contents()     # 清除 C3:R{end_of_row} 範圍的內容



def clear_hanji_in_cells(wb, sheet_name='漢字注音', source_cell='V3', clear_source=False):
    # 選擇指定的工作表
    sheet = wb.sheets[sheet_name]

    # 取得 V3 儲存格的字串
    v3_value = sheet.range(source_cell).value

    # 計算字串的總長度
    total_length = len(v3_value)
    print(f" {total_length} 個字元")

    # 每頁最多處理 20 列
    TOTAL_ROWS = int(wb.names['每頁總列數'].refers_to_range.value) # 自名稱為【每頁總列數】之儲存格，取得【每頁最多處理幾列】之值
    # 每列最多處理 15 字元
    CHARS_PER_ROW = int(wb.names['每列總字數'].refers_to_range.value)  # 自名稱為【每列總字數】之儲存格，取得【每列最多處理幾個字元】之值
    # 設定起始及結束的欄位  （【D欄=4】到【R欄=18】）
    start = 4
    end = start + CHARS_PER_ROW

    # 迴圈清空所有漢字的上下方儲存格 (羅馬拼音和台語注音符號)
    row = 5
    for i in range(TOTAL_ROWS):
        for col in range(start, end):  # 【D欄=4】到【R欄=18】
            # 清空漢字儲存格 (Row)
            sheet.range((row, col)).value = None
            # 清空上方的台語拼音儲存格 (Row-1)
            sheet.range((row - 1, col)).value = None
            # 清空下方的台語注音儲存格 (Row+1)
            sheet.range((row + 1, col)).value = None
            # 清空填入注音的儲存格 (Row-2)
            sheet.range((row - 2, col)).value = None

            # 顯示清空的儲存格
            col_name = xw.utils.col_name(col)
            print(f"清空第 {row} 列，第 {col_name} 欄")

        # 每處理 15 個字元後，換到下一行
        row += 4
        print("\n")

    # =========================================================================
    # (2) 清除原先已填入的漢字
    # =========================================================================
    if clear_source:
        sheet.range(source_cell).value = ""
        print(f"清空原先的漢字：{source_cell}")
