def fill_hanji_in_cells(wb, sheet_name='漢字注音', cell='V3'):
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

        # 逐字處理字串，並填入對應的儲存格
        row = 5
        index = 0  # 用來追蹤目前處理到的字元位置

        # 逐字處理字串 
        while index < total_length:     # 使用 while 而非 for，確保處理完整個字串
            for col in range(4, 19):  # 【D欄=4】到【R欄=18】
                # 確認是否還有字元可以處理
                if index < total_length:
                    # 取得當前字元
                    char = v3_value[index]

                    if char != "\n":
                        # 將字元填入對應的儲存格
                        sheet.range((row, col)).value = char
                    else:
                        # 若遇到換行字元，直接跳過
                        index += 1
                        break  

                    # 更新索引，處理下一個字元
                    index += 1
                else:
                    break  # 若字串已處理完畢，退出迴圈

            # 每處理 15 個字元後，換到下一行
            row += 4

    # 保存 Excel 檔案
    wb.save()

    # 選擇名為 "顯示注音輸入" 的命名範圍
    named_range = wb.names['顯示注音輸入']
    named_range.refers_to_range.value = True

    print(f"已成功更新，漢字已填入對應儲存格，上下方儲存格已清空。")