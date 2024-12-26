# 查找漢字讀音，並標註台語音標和注音符號
import sqlite3

import xlwings as xw

from mod_file_access import load_module_function
from mod_標音 import hong_im_tng_tai_gi_im_piau  # 方音符號轉台語音標
from mod_標音 import is_punctuation  # 是否為標點符號
from mod_標音 import split_hong_im_hu_ho  # 分解漢字標音
from mod_標音 import split_tai_gi_im_piau  # 分解台語音標
from mod_標音 import tlpa_tng_han_ji_piau_im  # 台語音標轉漢字標音
from mod_標音 import PiauIm

# ==========================================================
# 注音法設定和共用變數
# ==========================================================
zu_im_huat_list = {
    "SNI": ["fifteen_yin", "rt", "十五音切語"],
    "TPS": ["Piau_Im", "rt", "方音符號注音"],
    "POJ": ["pin_yin", "rt", "白話字拼音"],
    "TL": ["pin_yin", "rt", "台羅拼音"],
    "BP": ["pin_yin", "rt", "閩拼標音"],
    "TLPA_Plus": ["pin_yin", "rt", "台羅改良式"],
    "DBL": ["Siang_Pai", "rtc", "雙排注音"],
}


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


def ca_han_ji_thak_im(wb, sheet_name='漢字注音', cell='V3', hue_im="白話音", han_ji_khoo="河洛話", db_name='Ho_Lok_Ue.db', module_name='mod_河洛話', function_name='han_ji_ca_piau_im'):
    # 初始化 PiauIm 類別，産生標音物件
    piau_im = PiauIm(han_ji_khoo=han_ji_khoo)
    piau_im_huat = wb.names['標音方法'].refers_to_range.value
    # piau_im_huat = '方音符號'

    # 顯示「已輸入之拼音字母及注音符號」
    named_range = wb.names['顯示注音輸入']
    named_range.refers_to_range.value = True

    # 選擇工作表
    sheet = wb.sheets[sheet_name]
    sheet.activate()
    sheet.range('A1').select()

    # 取得 V3 儲存格的字串
    v3_value = sheet.range(cell).value

    # 取得工作表能處理最多列數： 20 列
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
                    tai_gi_im_piau = ""
                    han_ji_piau_im = ""
                    result = None
                    msg = ""

                    char = v3_value[index]
                    if char == "\n":
                        index += 1
                        break

                    cell_value = sheet.range((row, col)).value
                    if is_punctuation(cell_value):
                        msg = cell_value
                        print(f"({row}, {col_name}) = {msg}")
                        index += 1
                        continue
                    else:
                        han_ji = cell_value

                    manual_input = sheet.range((row-2, col)).value
                    if manual_input:
                        if '〔' in manual_input and '〕' in manual_input:
                            # 將人工輸入的〔台語音標〕轉換成【方音符號】
                            im_piau = manual_input.split('〔')[1].split('〕')[0]
                            siann, un, tiau = split_tai_gi_im_piau(im_piau)
                            tai_gi_im_piau = ''.join([siann, un, tiau])
                            # 依使用者指定之【標音方法】，將【台語音標】轉換成其所需之【漢字標音】
                            han_ji_piau_im = tlpa_tng_han_ji_piau_im(
                                piau_im=piau_im,
                                piau_im_huat=piau_im_huat,
                                tai_gi_im_piau=tai_gi_im_piau
                            )
                        elif '【' in manual_input and '】' in manual_input:
                            # 將人工輸入的【方音符號】轉換成【台語音標】
                            han_ji_piau_im = manual_input.split('【')[1].split('】')[0]
                            siann, un, tiau = split_hong_im_hu_ho(han_ji_piau_im)
                            # 依使用者指定之【標音方法】，將【台語音標】轉換成其所需之【漢字標音】
                            tai_gi_im_piau = hong_im_tng_tai_gi_im_piau(
                                siann=siann,
                                un=un,
                                tiau=tiau,
                                cursor=cursor,
                            )['台語音標']
                        else:
                            # 將人工輸入，沒有以[中括號]標示的【台語音標】轉換成【方音符號】
                            im_piau = manual_input
                            siann, un, tiau = split_tai_gi_im_piau(im_piau)
                            tai_gi_im_piau = ''.join([siann, un, tiau])
                            # 依使用者指定之【標音方法】，將【台語音標】轉換成其所需之【漢字標音】
                            han_ji_piau_im = tlpa_tng_han_ji_piau_im(
                                piau_im=piau_im,
                                piau_im_huat=piau_im_huat,
                                tai_gi_im_piau=tai_gi_im_piau
                            )

                        sheet.range((row - 1, col)).value = tai_gi_im_piau
                        sheet.range((row + 1, col)).value = han_ji_piau_im
                    else:
                        result = han_ji_ca_piau_im(cursor=cursor, han_ji=han_ji, hue_im=hue_im)

                        if result:
                            if han_ji_khoo == "河洛話":
                                #-----------------------------------------------------------------
                                # 【白話音】：依《河洛話漢字庫》標注【台語音標】和【方音符號】
                                #-----------------------------------------------------------------
                                # 將【台語音標】分解為【聲母】、【韻母】、【聲調】
                                siann_bu = result[0]['聲母']
                                un_bu = result[0]['韻母']
                                tiau_ho = result[0]['聲調']
                                # if siann_bu == "" or siann_bu == None:
                                #     siann_bu = "Ø"

                                if tiau_ho == "6":
                                    # 若【聲調】為【6】，則將【聲調】改為【7】
                                    tiau_ho = "7"

                                # 將【聲母】、【韻母】、【聲調】，合併成【台語音標】
                                # tai_gi_im_piau = siann_bu + un_bu + tiau_ho
                                tai_gi_im_piau = ''.join([siann_bu, un_bu, tiau_ho])

                                # 依使用者指定之【標音方法】，將【台語音標】轉換成其所需之【漢字標音】
                                # han_ji_piau_im = tlpa_tng_han_ji_piau_im(
                                #     piau_im=piau_im,
                                #     piau_im_huat=piau_im_huat,
                                #     tai_gi_im_piau=tai_gi_im_piau
                                # )

                                zu_im_list = split_tai_gi_im_piau(tai_gi_im_piau)
                                if zu_im_list[0] == "" or zu_im_list[0] == None:
                                    siann_bu = "Ø"
                                else:
                                    siann_bu = zu_im_list[0]

                                han_ji_piau_im = choose_piau_im_method(
                                    piau_im,
                                    piau_im_huat,
                                    siann_bu,
                                    zu_im_list[1],
                                    zu_im_list[2]
                                )
                            else:
                                #-----------------------------------------------------------------
                                # 【文讀音】：依《廣韻字庫》標注【台語音標】和【方音符號】
                                #-----------------------------------------------------------------
                                siann_bu, un_bu, tiau_ho = split_tai_gi_im_piau(result[0]['標音'])

                                # 將【台語音標】分解為【聲母】、【韻母】、【聲調】
                                if siann_bu == "" or siann_bu == None:
                                    siann_bu = "Ø"
                                tai_gi_im_piau = siann_bu + un_bu + tiau_ho

                                # 依使用者指定之【標音方法】，將【台語音標】轉換成其所需之【漢字標音】
                                han_ji_piau_im = tlpa_tng_han_ji_piau_im(
                                    piau_im=piau_im,
                                    piau_im_huat=piau_im_huat,
                                    tai_gi_im_piau=tai_gi_im_piau
                                )
                            sheet.range((row - 1, col)).value = tai_gi_im_piau
                            sheet.range((row + 1, col)).value = han_ji_piau_im
                        else:
                            msg = f"【{cell_value}】查無此字！"
                    if tai_gi_im_piau and han_ji_piau_im:
                        print(f"({row}, {col_name}) = {han_ji} [{tai_gi_im_piau}] 【{han_ji_piau_im}】")
                    else:
                        print(f"({row}, {col_name}) = {msg}")

                    index += 1

                row += 4
                print("\n")
        print("已完成【台語音標】和【方音符號】標注工作。")

    conn.close()

    wb.save()
