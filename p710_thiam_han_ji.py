# 用途：將漢字填入對應的儲存格
# 詳述：待加註讀音的漢字文置於 V3 儲存格。本程式將漢字逐字填入對應的儲存格：
# 【第一列】D5, E5, F5,... ,R5；
# 【第二列】D9, E9, F9,... ,R9；
# 【第三列】D13, E13, F13,... ,R13；
# 每個漢字佔一格，每格最多容納一個漢字。
# 漢字上方的儲存格（如：D4）為：【台語音標】欄，由【羅馬拼音字母】組成拼音。
# 漢字下方的儲存格（如：D6）為：【台語注音符號】欄，由【台語方音符號】組成注音。
# 漢字上上方的儲存格（如：D3）為：【人工標音】欄，可以只輸入【台語音標】；或
# 【台語音標】和【台語注音符號】皆輸入。
import xlwings as xw


def fill_hanji_in_cells(wb, sheet_name='漢字注音', cell='V3'):
    # 選擇指定的工作表
    sheet = wb.sheets[sheet_name]
    sheet.activate()               # 將「漢字注音」工作表設為作用中工作表
    sheet.range('A1').select()     # 將 A1 儲存格設為作用儲存格

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
            # 設定當前作用儲存格，根據 `row` 和 `col` 動態選取
            sheet.range((row, 1)).select()

            for col in range(start, end):  # 【D欄=4】到【R欄=18】
                # 確認是否還有字元可以處理
                if index < total_length:
                    # 取得當前字元
                    char = v3_value[index]

                    # 檢查下一個字元（確保 index + 1 在範圍內）
                    next_char = v3_value[index + 1] if index + 1 < total_length else None

                    # 自動檢查是否為【不】字，且下一個字元為【？】
                    if char == '不' and next_char == '？':
                        # 在【人工標音】欄自動輸入【hiu2】
                        # 【人工標音】欄位為漢字上方的儲存格（如 D3）
                        sheet.range((row - 2, col)).value = "hiu2"
                        print(f"自動填入【hiu2】於 {xw.utils.col_name(col)}{row - 2}")

                    if char == "\n":
                        char = "=CHAR(10)"  # 換行字元

                    # 重置儲存格：文字顏色（黑色）及填滿色彩（無填滿）
                    sheet.range((row-2, col), (row+1, col)).color = None
                    sheet.range((row, col)).font.color = (0, 0, 0)
                    sheet.range((row, col)).font.color = (0, 0, 0)
                    sheet.range((row-2, col)).font.color = (255, 0, 0)
                    sheet.range((row-1, col)).font.color = 0x3399FF # 藍色
                    sheet.range((row+1, col)).font.color = 0x009900 # 綠色

                    # 將字元填入對應的儲存格
                    sheet.range((row, col)).value = char

                    col_name = xw.utils.col_name(col)
                    print(f"【{row} 列， {col_name} 欄】：{char}")

                    # 更新索引，處理下一個字元
                    index += 1

                    # 換行：列數加一，並從下一列的第一個字元開始
                    if char == "=CHAR(10)":
                        break
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