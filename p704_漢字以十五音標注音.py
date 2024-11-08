import sqlite3

import xlwings as xw

from mod_file_access import load_module_function
from mod_標音 import is_valid_han_ji, split_zu_im


# 十五音標注音
def zap_goo_im_piau_im(wb, sheet_name='十五音', cell='V3', hue_im="白話音", han_ji_khoo="河洛話", db_name='Ho_Lok_Ue.db', module_name='mod_標音', function_name='TLPA_Tng_Zap_Goo_Im'):
    named_range = wb.names['顯示注音輸入']
    named_range.refers_to_range.value = True

    sheet = wb.sheets[sheet_name]
    sheet.activate()
    v3_value = sheet.range(cell).value
    total_length = len(v3_value)

    TOTAL_ROWS = int(wb.names['每頁總列數'].refers_to_range.value)
    CHARS_PER_ROW = int(wb.names['每列總字數'].refers_to_range.value)
    start, end = 4, 4 + CHARS_PER_ROW

    if total_length > (CHARS_PER_ROW * TOTAL_ROWS):
        print("可供作業之儲存格數太少，無法進行作業！")
        return

    han_ji_tng_piau_im = load_module_function(module_name, function_name)
    conn = sqlite3.connect(db_name)
    cursor = conn.cursor()

    row, index = 5, 0

    def process_cell(row, col, char):
        """處理單一儲存格的標音轉換和填寫"""
        col_name = xw.utils.col_name(col)
        cell_value = sheet.range((row, col)).value

        if not is_valid_han_ji(cell_value):
            print(f"({row}, {col_name}) = 無效漢字或標音缺失")
            return

        lo_ma_im_piau = sheet.range((row - 1, col)).value
        if not lo_ma_im_piau:
            print(f"缺少【台語音標】於({row - 1}, {col_name})")
            return

        try:
            siann_bu, un_bu, siann_tiau = split_zu_im(lo_ma_im_piau)
            zu_im_hu_ho = han_ji_tng_piau_im(siann_bu=siann_bu, un_bu=un_bu, siann_tiau=siann_tiau, cursor=cursor)
            sheet.range((row + 1, col)).value = zu_im_hu_ho['漢字標音']
            print(f"({row + 1}, {col_name}) = 【{char}】{zu_im_hu_ho['漢字標音']}")
        except ValueError as e:
            print(f"【台語音標】資料格式錯誤於({row - 1}, {col_name}): {e}")

    while index < total_length:
        for col in range(start, end):
            if index >= total_length:
                break
            char = v3_value[index]
            if char == "\n":
                index += 1
                continue  # 忽略換行符號

            process_cell(row, col, char)
            index += 1  # 處理下一個漢字

        row += 4  # 移動到下一組儲存格區域

    print("已完成【十五音】標注音工作。")
    conn.close()
    wb.save()
