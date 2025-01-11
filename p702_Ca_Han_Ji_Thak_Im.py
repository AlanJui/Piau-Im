# 查找漢字讀音，並標註台語音標和注音符號
import sqlite3

import xlwings as xw

from mod_file_access import load_module_function
from mod_標音 import hong_im_tng_tai_gi_im_piau  # 方音符號轉台語音標
from mod_標音 import is_punctuation  # 是否為標點符號
from mod_標音 import split_hong_im_hu_ho  # 分解漢字標音
from mod_標音 import split_tai_gi_im_piau  # 分解台語音標
from mod_標音 import tlpa_tng_han_ji_piau_im  # 台語音標轉漢字標音
from mod_標音 import tng_uann_han_ji_piau_im  # 台語音標轉台語音標
from mod_標音 import PiauIm
from p740_Phua_Im_Ji import PhuaImJi


def za_ji_kiat_ko_cut_piau_im(result, han_ji_khoo, piau_im, piau_im_huat):
    """查字結果出標音：查詢【漢字庫】取得之【查找結果】，將之切分：聲、韻、調"""
    if han_ji_khoo == "河洛話":
        #-----------------------------------------------------------------
        # 【白話音】：依《河洛話漢字庫》標注【台語音標】和【方音符號】
        #-----------------------------------------------------------------
        # 將【台語音標】分解為【聲母】、【韻母】、【聲調】
        siann_bu = result[0]['聲母']
        un_bu = result[0]['韻母']
        tiau_ho = result[0]['聲調']
        if tiau_ho == "6":
            # 若【聲調】為【6】，則將【聲調】改為【7】
            tiau_ho = "7"
    else:
        #-----------------------------------------------------------------
        # 【文讀音】：依《廣韻字庫》標注【台語音標】和【方音符號】
        #-----------------------------------------------------------------
        siann_bu, un_bu, tiau_ho = split_tai_gi_im_piau(result[0]['標音'])
        if siann_bu == "" or siann_bu == None:
            siann_bu = "ø"

    # 將【聲母】、【韻母】、【聲調】，合併成【台語音標】
    # tai_gi_im_piau = ''.join([siann_bu, un_bu, tiau_ho])
    tai_gi_im_piau = siann_bu + un_bu + tiau_ho

    # 標音法為：【十五音】或【雅俗通】，且【聲母】為空值，則將【聲母】設為【ø】
    if (piau_im_huat == "十五音" or piau_im_huat == "雅俗通") and (siann_bu == "" or siann_bu == None):
        siann_bu = "ø"
    han_ji_piau_im = tng_uann_han_ji_piau_im(
        piau_im,
        piau_im_huat,
        siann_bu,
        un_bu,
        tiau_ho
    )
    return tai_gi_im_piau, han_ji_piau_im


def ca_han_ji_thak_im(wb, sheet_name='漢字注音', cell='V3', ue_im_lui_piat="白話音", han_ji_khoo="河洛話", db_name='Ho_Lok_Ue.db', module_name='mod_河洛話', function_name='han_ji_ca_piau_im'):
    """查漢字讀音：依【漢字】查找【台語音標】，並依指定之【標音方法】輸出【漢字標音】"""
    # 動態載入查找函數
    han_ji_ca_piau_im = load_module_function(module_name, function_name)

    # 連接指定資料庫
    conn = sqlite3.connect(db_name)
    cursor = conn.cursor()

    # 初始化 PiauIm 類別，産生標音物件
    piau_im = PiauIm(han_ji_khoo=han_ji_khoo)
    piau_im_huat = wb.names['標音方法'].refers_to_range.value
    # piau_im_huat = '方音符號'
    phua_im_ji = PhuaImJi()

    # 顯示「已輸入之拼音字母及注音符號」
    named_range = wb.names['顯示注音輸入']
    named_range.refers_to_range.value = True

    # 選擇工作表
    sheet = wb.sheets[sheet_name]
    sheet.activate()
    sheet.range('A1').select()

    # 取得工作表能處理最多列數： 20 列
    TOTAL_LINES = int(wb.names['每頁總列數'].refers_to_range.value)
    row = 5

    # 每列最多處理 15 字元
    CHARS_PER_ROW = int(wb.names['每列總字數'].refers_to_range.value)
    # 設定起始及結束的欄位  （【D欄=4】到【R欄=18】）
    start = 4
    end = start + CHARS_PER_ROW

    # 逐字處理字串，並填入對應的儲存格
    EOF = False
    line = 1
    while line < TOTAL_LINES and not EOF:
        # 設定【作用儲存格】為列首
        sheet.range((row, 1)).select()
        Two_Empty_Cells = 0
        for col in range(start, end):
            msg = ""
            col_name = xw.utils.col_name(col)

            # 取得當前字元
            han_ji_u_piau_im = False
            cell_value = sheet.range((row, col)).value

            if cell_value == 'φ':
                EOF = True
                msg = "【文字終結】"
            elif cell_value == '\n':
                msg = "【換行】"
            elif cell_value == None:
                if Two_Empty_Cells == 0:
                    Two_Empty_Cells += 1
                elif Two_Empty_Cells == 1:
                    EOF = True
                msg = "【缺空】"    # 表【儲存格】未填入任何字/符，不同於【空白】字元
            else:
                # 若不為【標點符號】，則以【漢字】處理
                if is_punctuation(cell_value):
                    msg = f"{cell_value}"
                else:
                    # 查找漢字讀音
                    han_ji = cell_value

                    # 依據【人工標音】欄是否有輸入，決定【漢字標音】之處理方式
                    manual_input = sheet.range((row-2, col)).value

                    if manual_input:    # 若有人工輸入之處理作業
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
                            han_ji_u_piau_im = True
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
                            han_ji_u_piau_im = True
                        else:
                            # 將人工輸入的【台語音標】，解構為【聲母】、【韻母】、【聲調】
                            tai_gi_im_piau = manual_input
                            siann, un, tiau = split_tai_gi_im_piau(tai_gi_im_piau)
                            # 依指定之【標音方法】，將【台語音標】轉換成其所需之【漢字標音】
                            han_ji_piau_im = tng_uann_han_ji_piau_im(
                                piau_im,
                                piau_im_huat,
                                siann,
                                un,
                                tiau
                            )
                            han_ji_u_piau_im = True

                        # 將人工輸入的【台語音標】置入【破音字庫】Dict
                        phua_im_ji.ka_phua_im_ji(han_ji, tai_gi_im_piau)
                    else:               # 無人工輸入，則自【漢字庫】查找作業
                        # 查找【破音字庫】，確認是否有此漢字
                        found = phua_im_ji.ca_phua_im_ji(han_ji)
                        # 若【破音字庫】有此漢字
                        if found:
                            siann_bu, un_bu, tiau_ho = split_tai_gi_im_piau(found)
                            tai_gi_im_piau = siann_bu + un_bu + tiau_ho
                            han_ji_piau_im = tng_uann_han_ji_piau_im(
                                piau_im,
                                piau_im_huat,
                                siann_bu,
                                un_bu,
                                tiau_ho
                            )
                            han_ji_u_piau_im = True
                            sheet.range((row, col)).font.color = (255, 0, 0)    # 將文字顏色設為【紅色】
                            sheet.range((row, col)).color = (255, 255, 0)       # 將底色設為【黄色】
                            print(f"漢字：【{han_ji}】之注音【{tai_gi_im_piau}】取自【人工注音字典】。")
                        # 若【破音字庫】無此漢字，則在資料庫中查找
                        else:
                            result = han_ji_ca_piau_im(cursor=cursor, han_ji=han_ji, hue_im=ue_im_lui_piat)
                            if not result:
                                msg = f"【{han_ji}】查無此字！"
                            else:
                                # 依【漢字庫】查找結果，輸出【台語音標】和【漢字標音】
                                tai_gi_im_piau, han_ji_piau_im = za_ji_kiat_ko_cut_piau_im(
                                    result=result,
                                    han_ji_khoo=han_ji_khoo,
                                    piau_im=piau_im,
                                    piau_im_huat=piau_im_huat
                                )
                                han_ji_u_piau_im = True

                if han_ji_u_piau_im:
                    sheet.range((row - 1, col)).value = tai_gi_im_piau
                    sheet.range((row + 1, col)).value = han_ji_piau_im
                    msg = f"{han_ji}： [{tai_gi_im_piau}] /【{han_ji_piau_im}】"

            # 顯示處理進度
            print(f"({row}, {col_name}) = {msg}")

            # 若讀到【換行】或【文字終結】，跳出逐欄取字迴圈
            if msg == "【換行】" or EOF:
                break

        # 每當處理一行 15 個漢字後，亦換到下一行
        print("\n")
        row += 4

    print("已完成【台語音標】和【方音符號】標注工作。")

    # 關閉資料庫連線
    conn.close()

    wb.save()
