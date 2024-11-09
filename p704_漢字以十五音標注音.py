import sqlite3

import xlwings as xw

# from mod_file_access import load_module_function
from mod_標音 import PiauIm, is_valid_han_ji, split_zu_im


def choose_piau_im_method(piau_im, zu_im_huat, siann_bu, un_bu, tiau_ho):
    """選擇並執行對應的注音方法"""
    if zu_im_huat == "十五音":
        return piau_im.SNI_piau_im(siann_bu, un_bu, tiau_ho)
    elif zu_im_huat == "白話字":
        return piau_im.POJ_piau_im(siann_bu, un_bu, tiau_ho)
    elif zu_im_huat == "台羅拼音":
        return piau_im.TL_piau_im(siann_bu, un_bu, tiau_ho)
    elif zu_im_huat == "閩拼方案":
        return piau_im.BP_piau_im(siann_bu, un_bu, tiau_ho)
    elif zu_im_huat == "方音符號":
        return piau_im.TPS_piau_im(siann_bu, un_bu, tiau_ho)
    elif zu_im_huat == "台語音標":
        siann = piau_im.Siann_Bu_Dict[siann_bu]["台語音標"] or ""
        un = piau_im.Un_Bu_Dict[un_bu]["台語音標"]
        return f"{siann}{un}{tiau_ho}"
    return ""

# 十五音標注音
def han_ji_piau_im(wb, sheet_name='十五音', cell='V3', hue_im="白話音", han_ji_khoo="河洛話", db_name='Ho_Lok_Ue.db'):
    # 初始化 PiauIm 類別，産生標音物件
    piau_im = PiauIm(han_ji_khoo=han_ji_khoo)
    piau_im_huat = wb.names['標音方法'].refers_to_range.value

    # 顯示「已輸入之拼音字母及注音符號」
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

    conn = sqlite3.connect(db_name)
    cursor = conn.cursor()

    row, index = 5, 0

    def process_cell(row, col, char, piau_im=piau_im):
        """處理單一儲存格的標音轉換和填寫"""
        col_name = xw.utils.col_name(col)
        cell_value = sheet.range((row, col)).value

        if not is_valid_han_ji(cell_value):
            # 若儲存格內容不是漢字，應是：標點符號或空白，故將其顯示
            print(f"({row}, {col_name}) = {cell_value}")
            return

        lo_ma_im_piau = sheet.range((row - 1, col)).value
        if not lo_ma_im_piau:
            print(f"缺少【台語音標】於({row - 1}, {col_name})")
            return

        try:
            siann_bu, un_bu, siann_tiau = split_zu_im(lo_ma_im_piau)

            if siann_bu == "" or siann_bu == None:
                siann_bu = "Ø"

            han_ji_piau_im = choose_piau_im_method(
                piau_im,
                piau_im_huat,
                siann_bu,
                un_bu,
                siann_tiau
            )
            sheet.range((row + 1, col)).value = han_ji_piau_im
            print(f"({row + 1}, {col_name}) = 【{char}】{han_ji_piau_im}")
        except ValueError as e:
            print(f"【台語音標】資料格式錯誤於({row - 1}, {col_name}): {e}")

    while index < total_length:
        sheet.range((row, 1)).select()
        for col in range(start, end):
            if index >= total_length:
                break
            char = v3_value[index]
            if char == "\n":
                index += 1
                break  # 跳出內部 for 迴圈，繼續處理下一列

            process_cell(row, col, char, piau_im=piau_im)
            index += 1  # 處理下一個漢字

        row += 4  # 移動到下一組儲存格區域

    print("已完成【十五音】標注音工作。")
    conn.close()
    wb.save()
