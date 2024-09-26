import xlwings as xw


def thiam_zu_im(wb, sheet_name='漢字注音', cell='V3'):
    # 顯示「已輸入之拼音字母及注音符號」 
    named_range = wb.names['顯示注音輸入']  # 選擇名為 "顯示注音輸入" 的命名範圍# 選擇名為 "顯示注音輸入" 的命名範圍
    named_range.refers_to_range.value = True

    # 選擇工作表
    sheet = wb.sheets[sheet_name]

    # 取得 V3 儲存格的字串
    v3_value = sheet.range(cell).value
    
    # 每頁最多處理 20 列
    TOTAL_ROWS = int(wb.names['每頁總列數'].refers_to_range.value) # 自名稱為【每頁總列數】之儲存格，取得【每頁最多處理幾列】之值
    # 每列最多處理 15 字元
    CHARS_PER_ROW = int(wb.names['每列總字數'].refers_to_range.value)  # 自名稱為【每列總字數】之儲存格，取得【每列最多處理幾個字元】之值
    # 設定起始及結束的欄位  （【D欄=4】到【R欄=18】）
    start = 4
    end = start + CHARS_PER_ROW

    # 計算字串的總長度
    total_length = len(v3_value)

    # 確認 V3 不為空
    if total_length and total_length < (CHARS_PER_ROW * TOTAL_ROWS):
        # # 清空 Row: 5, 9, 13, ... 漢字所在儲存格，上方的台語音標儲存格，及下方的台語注音符號儲存格
        # row = 5
        # index = 0  # 漢字處理指標

        # # 逐字處理字串 
        # while index < total_length:     # 使用 while 而非 for，確保處理完整個字串
        #     for col in range(start, end):  # 【D欄=4】到【R欄=18】
        #         # 確認是否還有字元可以處理
        #         if index < total_length:
        #             # 取得當前字元
        #             char = v3_value[index]

        #             if char != "\n":
        #                 # 清空上方的台語音標儲存格
        #                 sheet.range((row - 1, col)).value = None

        #                 # 清空下方的台語注音符號儲存格
        #                 sheet.range((row + 1, col)).value = None

        #             else:
        #                 # 若遇到換行字元，退出迴圈 
        #                 index += 1
        #                 break;  

        #             # 更新索引，處理下一個字元
        #             index += 1
        #         else:
        #             break  # 若字串已處理完畢，退出迴圈
        #     # 每處理 15 個字元後，換到下一行
        #     row += 4

        # 逐行處理資料，從 Row 3 開始，每列處理 15 個字元
        row = 3
        index = 0  # 漢字處理指標
        while index < total_length:     # 使用 while 而非 for，確保處理完整個字串
            for col in range(start, end):  # 【D欄=4】到【R欄=18】
                col_name = xw.utils.col_name(col)
                char = None
                han_ji = ""
                lo_ma_im_piau = ""
                zu_im_hu_ho = ""
                # 確認是否還有字元可以處理
                if index < total_length:
                    char = v3_value[index]  # 取得目前欲處理的【漢字】
                    if char != "\n":    # 確認待處理的【漢字】不是【換行字元: \n】
                        cell_value = sheet.range((row, col)).value  # 取得 D4, E4, ..., R4 的內容
                        if cell_value:  # 確認儲存格有填入【拚音字母/注音符號】
                            # 取得正在注音的漢字
                            han_ji = sheet.range((row + 2, col)).value

                            # 分割字串來提取羅馬拼音和台語注音
                            lo_ma_im_piau = cell_value.split('〔')[1].split('〕')[0]  # 取得〔羅馬拼音〕
                            zu_im_hu_ho = cell_value.split('【')[1].split('】')[0]  # 取得【台語注音】

                            # 將羅馬拼音填入當前 row + 1 的儲存格
                            sheet.range((row + 1, col)).value = lo_ma_im_piau

                            # 將台語注音填入當前 row + 3 的儲存格
                            sheet.range((row + 3, col)).value = zu_im_hu_ho
                    else:
                        # 若遇到換行字元，退出迴圈 
                        index += 1
                        break;  

                    # 顯示當前處理的【漢字】、【羅馬拼音】和【台語注音】
                    if lo_ma_im_piau and zu_im_hu_ho:
                        print(f"({row}, {col_name}) = {han_ji} [{lo_ma_im_piau}] 【{zu_im_hu_ho}】")
                    else:
                        print(f"({row}, {col_name}) = {char}")

                    # 更新索引，處理下一個字元
                    index += 1
                else:
                    break  # 若字串已處理完畢，退出迴圈
            # 每處理 15 個字元後，換到下一行
            row += 4
            print("\n")
            
    print("已完成【台語音標】和【台語注音符號】標註工作。")

    # 保存 Excel 檔案
    # wb.save('Tai_Gi_Zu_Im_Bun.xlsx')
    wb.save()
    # wb.close()

    # 令人工手動填入的台語音標和注音符號不要顯示

    # 不要顯示「已輸入之拼音字母及注音符號」 
    named_range = wb.names['顯示注音輸入']  # 選擇名為 "顯示注音輸入" 的命名範圍# 選擇名為 "顯示注音輸入" 的命名範圍
    named_range.refers_to_range.value = False