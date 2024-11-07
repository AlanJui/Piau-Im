# 查找漢字讀音，並標註台語音標和注音符號
import sqlite3

import xlwings as xw

from mod_file_access import load_module_function
from mod_標音 import is_valid_han_ji, split_zu_im


# 十五音標注音
def zap_goo_im_piau_im(wb, sheet_name='十五音', cell='V3', hue_im="白話音", han_ji_khoo="河洛話", db_name='Ho_Lok_Ue.db', module_name='mod_標音', function_name='TLPA_Tng_Zap_Goo_Im'):
    # 顯示「已輸入之拼音字母及注音符號」
    named_range = wb.names['顯示注音輸入']
    named_range.refers_to_range.value = True

    # 選擇工作表
    sheet = wb.sheets[sheet_name]
    sheet.activate()
    sheet.range('A1').select()

    # 取得 V3 儲存格的字串及總漢字數
    v3_value = sheet.range(cell).value
    total_length = len(v3_value)

    # 每頁最多處理 20 列
    TOTAL_ROWS = int(wb.names['每頁總列數'].refers_to_range.value)
    # 每列最多處理 15 字元
    CHARS_PER_ROW = int(wb.names['每列總字數'].refers_to_range.value)
    # 設定起始及結束的欄位  （【D欄=4】到【R欄=18】）
    start = 4
    end = start + CHARS_PER_ROW

    # 檢查儲存格數是否足夠
    if total_length > (CHARS_PER_ROW * TOTAL_ROWS):
        print("可供作業之儲存格數太少，無法進行作業！")
        return

    # 動態載入查找函數：漢字標音轉換
    han_ji_tng_piau_im = load_module_function(module_name, function_name)

    # 連接指定資料庫
    conn = sqlite3.connect(db_name)
    cursor = conn.cursor()

    # 初始化行數和索引
    row = 5
    index = 0

    while index < total_length:
        sheet.range((row, 1)).select()  # 選擇行起始儲存格

        for col in range(start, end):
            if index >= total_length:
                break  # 當所有漢字處理完畢時，停止內部迴圈

            # han_ji = v3_value[index]  # 從 V3 字串中取得漢字
            col_name = xw.utils.col_name(col)
            char = v3_value[index]
            if char == "\n":
                index += 1
                break

            cell_value = sheet.range((row, col)).value
            if not is_valid_han_ji(cell_value):
                msg = cell_value
                print(f"({row}, {col_name}) = {msg}")
                index += 1
                continue
            else:
                han_ji = cell_value

            # 從【台語音標】儲存格讀取聲母、韻母、聲調
            lo_ma_im_piau = sheet.range((row - 1, col)).value
            if lo_ma_im_piau:
                try:
                    siann_bu, un_bu, siann_tiau = split_zu_im(lo_ma_im_piau)
                except ValueError as e:
                    print(f"【台語音標】資料格式錯誤於({row - 1}, {col}): {e}")
                    index += 1
                    continue

                # 使用 TLPA_Tng_Zap_Goo_Im 函數轉換成十五音
                zu_im_hu_ho = han_ji_tng_piau_im(siann_bu=siann_bu, un_bu=un_bu, siann_tiau=siann_tiau, cursor=cursor)

                # 將十五音結果寫入【漢字注音】儲存格
                sheet.range((row + 1, col)).value = zu_im_hu_ho['漢字標音']
                print(f"({row + 1}, {xw.utils.col_name(col)}) = 【{han_ji}】{zu_im_hu_ho['漢字標音']}")
            else:
                print(f"缺少【台語音標】於({row - 1}, {col})")

            index += 1  # 處理下一個漢字

        row += 4  # 移動到下一組儲存格區域

    print("已完成【十五音】標注音工作。")

    # 關閉資料庫連線
    conn.close()

    # 保存活頁簿
    wb.save()
