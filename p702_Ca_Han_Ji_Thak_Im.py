# 查找漢字讀音，並標註台語音標和注音符號
import sqlite3

import xlwings as xw

# from mod_雅俗通 import han_ji_ca_piau_im
from mod_台羅音標漢字庫 import han_ji_ca_piau_im, split_zu_im
from p711_TL_Tng_Zu_Im import TL_Tng_Zu_Im


# =========================================================
# 判斷是否為標點符號的輔助函數
# =========================================================
def is_valid_han_ji(char):
    # 如果 char 是 None，直接返回 False
    if char is None:
        return False
    else:
        char = char.strip()
    
    # 可以根據需要擴充此列表以判斷各種標點符號
    punctuation_marks = "，。！？；：、（）「」『』《》……"
    return char not in punctuation_marks


def ca_han_ji_thak_im(wb, sheet_name='漢字注音', cell='V3'):
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

    #--------------------------------------------------------------------------------------------------------
    # 漢字所在儲存： 5, 9, 13, ... 
    # 漢字上方儲存：台語音標
    # 漢字下方儲存：台語注音符號
    #--------------------------------------------------------------------------------------------------------
    # 在所有測試開始前，連接資料庫
    conn = sqlite3.connect('Tai_Loo_Han_Ji_Khoo.db')  # 替換為實際資料庫路徑
    cursor = conn.cursor()

    # 確認 V3 不為空
    if total_length and total_length < (CHARS_PER_ROW * TOTAL_ROWS):
        # 逐行處理資料，從 Row 5 開始，每列處理 15 個字元
        row = 5
        index = 0  # 漢字處理指標
        while index < total_length:     # 使用 while 而非 for，確保處理完整個字串
            for col in range(start, end):  # 【D欄=4】到【R欄=18】
                if index == total_length:  # 確認是否已處理完整個字串
                    break
                
                col_name = xw.utils.col_name(col)
                char = None
                cell_value = ""
                han_ji = ''
                lo_ma_im_piau = ""
                zu_im_hu_ho = ""
                result = None
                msg = ""

                char = v3_value[index]  # 取得目前欲處理的【漢字】
                if char == "\n":    # 確認待處理的【漢字】不是【換行字元: \n】
                    # 若遇到換行字元，退出迴圈 
                    index += 1
                    break;  

                # 取得正在注音的漢字
                cell_value = sheet.range((row, col)).value  # 取得 D4, E4, ..., R4 的內容
                if not is_valid_han_ji(cell_value):  # 確認儲存格有填入【拚音字母/注音符號】
                    msg = cell_value  # 儲存格內容不是【漢字】，是【空白】，或是【標點符號】
                    print(f"({row}, {col_name}) = {msg}")
                    index += 1
                    continue  # 跳過此次迴圈，處理下一個字元   
                else:
                    han_ji = cell_value  # 取得【漢字】

                # 檢查有無手動填入【拚音字母/注音符號】(儲存格： D3, E3, ..., R3)
                manual_input = sheet.range((row-2, col)).value
                if manual_input:
                    # 檢查是否為完整格式〔put⁴〕【ㄅㄨㆵ】
                    if '〔' in manual_input and '〕' in manual_input and '【' in manual_input and '】' in manual_input:
                        lo_ma_im_piau = manual_input.split('〔')[1].split('〕')[0]
                        zu_im_hu_ho = manual_input.split('【')[1].split('】')[0]
                    else:
                        # 如果是簡化格式，例如 "put4"
                        zu_im_list = split_zu_im(manual_input)  # 將人工輸入的【羅馬拼音】拆分為【聲母】、【韻母】和【聲調】
                        zu_im_hu_ho = TL_Tng_Zu_Im(
                            siann_bu=zu_im_list[0],  # 聲母是拼音的前部分
                            un_bu=zu_im_list[1],  # 韻母是拼音的後部分
                            siann_tiau=zu_im_list[2],  # 最後一位是聲調
                            cursor=cursor
                        )['注音符號']  # 從 TL_Tng_Zu_Im 取得注音符號
                        lo_ma_im_piau = manual_input

                    sheet.range((row - 1, col)).value = lo_ma_im_piau
                    sheet.range((row + 1, col)).value = zu_im_hu_ho
                else:
                    # 查找漢字讀音 
                    result = han_ji_ca_piau_im(cursor, han_ji)

                    # 取羅馬拼音和台語注音
                    if result:
                        lo_ma_im_piau = result[0]['台語音標']  # 取得【台語音標】拼音羅馬字母
                        # 將台羅音標轉換為台語注音符號
                        zu_im_hu_ho = TL_Tng_Zu_Im(
                            siann_bu=result[0]['聲母'],
                            un_bu=result[0]['韻母'],
                            siann_tiau=result[0]['聲調'],
                            cursor=cursor
                        )

                        # 將羅馬拼音填入當前 row + 1 的儲存格
                        sheet.range((row - 1, col)).value = lo_ma_im_piau

                        # 將台語注音填入當前 row + 3 的儲存格
                        sheet.range((row + 1, col)).value = zu_im_hu_ho['注音符號']
                    else:
                        msg = f"【{cell_value}】查無此字！"
                # 顯示當前處理的【漢字】、【羅馬拼音】和【台語注音】
                if lo_ma_im_piau and zu_im_hu_ho:
                    print(f"({row}, {col_name}) = {han_ji} [{lo_ma_im_piau}] 【{zu_im_hu_ho}】")
                else:
                    print(f"({row}, {col_name}) = {msg}")

                # 更新索引，處理下一個字元
                index += 1
            # 每處理 15 個字元後，換到下一行
            row += 4
            print("\n")
    print("已完成【台語音標】和【台語注音符號】標註工作。")

    # 在所有測試結束後，關閉資料庫連接
    conn.close()

    # 保存 Excel 檔案
    # wb.save('Tai_Gi_Zu_Im_Bun.xlsx')
    wb.save()
    # wb.close()

    # 令人工手動填入的台語音標和注音符號不要顯示

    # 不要顯示「已輸入之拼音字母及注音符號」 
    # named_range = wb.names['顯示注音輸入']  # 選擇名為 "顯示注音輸入" 的命名範圍# 選擇名為 "顯示注音輸入" 的命名範圍
    # named_range.refers_to_range.value = False