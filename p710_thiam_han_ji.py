import math

import xlwings as xw


def fill_hanji_in_cells(file_name, sheet_name='漢字注音', cell='V3'):
    # 打開 Excel 檔案
    wb = xw.Book(file_name)

    # 選擇指定的工作表
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

        # 迴圈清空所有漢字的上下方儲存格 (羅馬拼音和台語注音符號)
        row = 5
        for i in range(total_rows_needed):
            for col in range(4, 19):  # 【D欄=4】到【R欄=18】
                # 清空漢字儲存格 (Row)
                sheet.range((row, col)).value = None
                # 清空上方的台語拼音儲存格 (Row-1)
                sheet.range((row - 1, col)).value = None
                # 清空下方的台語注音儲存格 (Row+1)
                sheet.range((row + 1, col)).value = None
                # 清空填入注音的儲存格 (Row-2)
                sheet.range((row - 2, col)).value = None

            # 每處理 15 個字元後，換到下一行
            row += 4

        # 逐字處理字串，並填入對應的儲存格
        row = 5
        index = 0  # 用來追蹤目前處理到的字元位置
        for i in range(total_rows_needed):
            for col in range(4, 19):  # 【D欄=4】到【R欄=18】
                # 確認是否還有字元可以處理
                if index < total_length:
                    # 取得當前字元
                    char = v3_value[index]

                    # 將字元填入對應的儲存格
                    sheet.range((row, col)).value = char

                    # 更新索引，處理下一個字元
                    index += 1
                else:
                    break  # 若字串已處理完畢，退出迴圈

            # 每處理 15 個字元後，換到下一行
            row += 4

    # 保存 Excel 檔案
    wb.save(file_name)
    # wb.close()

    print(f"{file_name} 已成功更新，漢字已填入對應儲存格，上下方儲存格已清空。")
