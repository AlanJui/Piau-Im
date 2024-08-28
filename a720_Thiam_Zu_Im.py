import math

import xlwings as xw


def thiam_zu_im(file_name, sheet_name='漢字注音', cell='V3'):
    # 打開 Excel 檔案
    wb = xw.Book(file_name)  

    # 選擇工作表
    # sheet = wb.sheets[0]  # 選擇第一個工作表
    sheet = wb.sheets[sheet_name]

    # 取得 V3 儲存格的字串
    v3_value = sheet.range(cell).value
    
    # 確認 V3 不為空
    if v3_value:
        # 計算字串的總長度
        total_length = len(v3_value)

        # 每列最多處理 15 個字元，計算總共需要多少列
        chars_per_row = 15
        total_rows_needed = math.ceil(total_length / chars_per_row)  # 無條件進位

        # 清空 Row: 5, 9, 13, ... 漢字所在儲存格，上方的台語音標儲存格，及下方的台語注音符號儲存格
        row = 5
        for i in range(total_rows_needed):
            for col in range(4, 19):  # 【D欄=4】到【R欄=18】
                # 清空漢字所在儲存格
                # sheet.range((row, col)).value = None

                # 清空上方的台語音標儲存格
                sheet.range((row - 1, col)).value = None

                # 清空下方的台語注音符號儲存格
                sheet.range((row + 1, col)).value = None

            # 每處理 15 個字元後，換到下一行
            row += 4

        # 逐行處理資料，從 Row 3 開始，每列處理 15 個字元
        row = 3
        for i in range(total_rows_needed):
            for col in range(4, 19):  # 【D欄=4】到【R欄=18】
                cell_value = sheet.range((row, col)).value  # 取得 D4, E4, ..., R4 的內容

                # 確認內容不為空
                if cell_value:
                    # 分割字串來提取羅馬拼音和台語注音
                    lo_ma_ji = cell_value.split('〔')[1].split('〕')[0]  # 取得〔羅馬拼音〕
                    zu_im_hu_ho = cell_value.split('【')[1].split('】')[0]  # 取得【台語注音】

                    # 將羅馬拼音填入當前 row + 1 的儲存格
                    sheet.range((row + 1, col)).value = lo_ma_ji

                    # 將台語注音填入當前 row + 3 的儲存格
                    sheet.range((row + 3, col)).value = zu_im_hu_ho

            # 每處理 15 個字元後，換到下一行
            row += 4
            
    print("已完成【台語音標】和【台語注音符號】標註工作。")

    # 保存 Excel 檔案
    wb.save('Tai_Gi_Zu_Im_Bun.xlsx')
    # wb.close()
