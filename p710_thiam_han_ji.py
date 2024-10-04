import xlwings as xw


def fill_hanji_in_cells(wb, sheet_name='漢字注音', cell='V3'):
    # 選擇指定的工作表
    sheet = wb.sheets[sheet_name]

    # 清空儲存格內容
    sheet.range('D3:R166').clear_contents()    # 清除 C3:R166 範圍的內容

    # 取得 V3 儲存格的字串
    v3_value = sheet.range(cell).value

    # 確認 V3 不為空
    if v3_value:
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

        # 逐字處理字串，並填入對應的儲存格
        row = 5
        index = 0  # 用來追蹤目前處理到的字元位置

        # 逐字處理字串 
        while index < total_length:     # 使用 while 而非 for，確保處理完整個字串
            for col in range(start, end):  # 【D欄=4】到【R欄=18】
                # 確認是否還有字元可以處理
                if index < total_length:
                    # 取得當前字元
                    char = v3_value[index]

                    if char != "\n":
                        # 將字元填入對應的儲存格
                        sheet.range((row, col)).value = char

                        col_name = xw.utils.col_name(col)
                        print(f"【{row} 列， {col_name} 欄】：{char}")
                    else:
                        # 若遇到換行字元，直接跳過
                        index += 1
                        break  

                    # 更新索引，處理下一個字元
                    index += 1
                else:
                    break  # 若字串已處理完畢，退出迴圈

            # 每處理 15 個字元後，換到下一行
            print("\n")
            row += 4

    # 保存 Excel 檔案
    wb.save()

    # 選擇名為 "顯示注音輸入" 的命名範圍
    named_range = wb.names['顯示注音輸入']
    named_range.refers_to_range.value = True

    print(f"已成功更新，漢字已填入對應儲存格，上下方儲存格已清空。")