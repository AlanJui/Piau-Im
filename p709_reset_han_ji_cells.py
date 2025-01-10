import xlwings as xw


#--------------------------------------------------------------------------
# 將待注音的【漢字儲存格】，文字顏色重設為黑色（自動 RGB: 0, 0, 0）；填漢顏色重設為無填滿
#--------------------------------------------------------------------------
def reset_han_ji_cells(wb, sheet_name='漢字注音'):
    # 選擇指定的工作表
    sheet = wb.sheets[sheet_name]
    sheet.activate()  # 將「漢字注音」工作表設為作用中工作表
    sheet.range('A1').select()     # 將 A1 儲存格設為作用儲存格

    # 每頁最多處理的列數
    TOTAL_ROWS = int(wb.names['每頁總列數'].refers_to_range.value)  # 從名稱【每頁總列數】取得值
    # 每列最多處理的字數
    CHARS_PER_ROW = int(wb.names['每列總字數'].refers_to_range.value)  # 從名稱【每列總字數】取得值

    # 設定起始及結束的欄位（【D欄=4】到【R欄=18】）
    start_col = 4
    end_col = start_col + CHARS_PER_ROW

    # 從第 5 列開始，每隔 4 列進行重置（5, 9, 13, ...）
    for row in range(5, 5 + 4 * TOTAL_ROWS, 4):
        for col in range(start_col, end_col):
            cell = sheet.range((row, col))
            # 將文字顏色設為【自動】（黑色）
            cell.font.color = (0, 0, 0)  # 設定為黑色
            # 將儲存格的填滿色彩設為【無填滿】
            cell.color = None

    print("漢字儲存格已成功重置，文字顏色設為自動，填滿色彩設為無填滿。")

    return 0