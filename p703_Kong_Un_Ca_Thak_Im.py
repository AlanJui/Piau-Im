# 查找漢字讀音，並標註台語音標和注音符號
import importlib
import sqlite3

import xlwings as xw

from mod_file_access import load_module_function
from mod_廣韻 import TL_Tng_Zu_Im
from mod_標音 import is_valid_han_ji, split_tai_gi_im_piau

# =========================================================
# 動態載入模組和函數
# =========================================================
# def load_module_function(module_name, function_name):
#     module = importlib.import_module(module_name)
#     return getattr(module, function_name)

def ca_han_ji_thak_im(wb, sheet_name='漢字注音', cell='V3', type="白話音", db_name='Tai_Loo_Han_Ji_Khoo.db', module_name='mod_台羅音標漢字庫', function_name='han_ji_ca_piau_im'):
    # 顯示「已輸入之拼音字母及注音符號」
    named_range = wb.names['顯示注音輸入']
    named_range.refers_to_range.value = True

    # 選擇工作表
    sheet = wb.sheets[sheet_name]
    sheet.activate()
    sheet.range('A1').select()

    # 取得 V3 儲存格的字串
    v3_value = sheet.range(cell).value

    # 每頁最多處理 20 列
    TOTAL_ROWS = int(wb.names['每頁總列數'].refers_to_range.value)
    # 每列最多處理 15 字元
    CHARS_PER_ROW = int(wb.names['每列總字數'].refers_to_range.value)
    # 設定起始及結束的欄位  （【D欄=4】到【R欄=18】）
    start = 4
    end = start + CHARS_PER_ROW

    total_length = len(v3_value)

    # 動態載入查找函數
    han_ji_ca_piau_im = load_module_function(module_name, function_name)

    # 連接指定資料庫
    conn = sqlite3.connect(db_name)
    cursor = conn.cursor()

    if total_length:
        if total_length > (CHARS_PER_ROW * TOTAL_ROWS):
            print("可供作業之儲存格數太少，無法進行作業！")
        else:
            row = 5
            index = 0
            while index < total_length:
                sheet.range((row, 1)).select()

                for col in range(start, end):
                    if index == total_length:
                        break

                    col_name = xw.utils.col_name(col)
                    char = None
                    cell_value = ""
                    han_ji = ''
                    lo_ma_im_piau = ""
                    zu_im_hu_ho = ""
                    result = None
                    msg = ""

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

                    manual_input = sheet.range((row-2, col)).value
                    if manual_input:
                        if '〔' in manual_input and '〕' in manual_input and '【' in manual_input and '】' in manual_input:
                            lo_ma_im_piau = manual_input.split('〔')[1].split('〕')[0]
                            zu_im_hu_ho = manual_input.split('【')[1].split('】')[0]
                        else:
                            zu_im_list = split_tai_gi_im_piau(manual_input)
                            zu_im_hu_ho = TL_Tng_Zu_Im(
                                siann_bu=zu_im_list[0],
                                un_bu=zu_im_list[1],
                                siann_tiau=zu_im_list[2],
                                cursor=cursor
                            )['注音符號']
                            lo_ma_im_piau = manual_input

                        sheet.range((row - 1, col)).value = lo_ma_im_piau
                        sheet.range((row + 1, col)).value = zu_im_hu_ho
                    else:
                        # result = han_ji_ca_piau_im(cursor, han_ji, type)
                        result = han_ji_ca_piau_im(cursor, han_ji)

                        if result:
                            # lo_ma_im_piau = result[0]['台語音標']
                            # zu_im_hu_ho = TL_Tng_Zu_Im(
                            #     siann_bu=result[0]['聲母'],
                            #     un_bu=result[0]['韻母'],
                            #     siann_tiau=result[0]['聲調'],
                            #     cursor=cursor
                            # )
                            lo_ma_im_piau = split_tai_gi_im_piau(result[0]['標音'])
                            zu_im_hu_ho = TL_Tng_Zu_Im(
                                siann_bu=lo_ma_im_piau[0],
                                un_bu=lo_ma_im_piau[1],
                                siann_tiau=lo_ma_im_piau[2],
                                cursor=cursor
                            )
                            sheet.range((row - 1, col)).value = ''.join(lo_ma_im_piau)
                            sheet.range((row + 1, col)).value = zu_im_hu_ho['注音符號']
                        else:
                            msg = f"【{cell_value}】查無此字！"
                    if lo_ma_im_piau and zu_im_hu_ho:
                        print(f"({row}, {col_name}) = {han_ji} [{lo_ma_im_piau}] 【{zu_im_hu_ho}】")
                    else:
                        print(f"({row}, {col_name}) = {msg}")

                    index += 1

                row += 4
                print("\n")
        print("已完成【台語音標】和【台語注音符號】標註工作。")

    conn.close()

    wb.save()
