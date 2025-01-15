# 查找漢字讀音，並標註台語音標和注音符號
import logging
import sqlite3

import xlwings as xw

from mod_excel_access import get_han_ji_khoo, get_tai_gi_by_han_ji, maintain_han_ji_koo
from mod_file_access import load_module_function
from mod_標音 import PiauIm  # 漢字之【漢字標音】轉換物件
from mod_標音 import hong_im_tng_tai_gi_im_piau  # 方音符號轉台語音標
from mod_標音 import is_punctuation  # 是否為標點符號
from mod_標音 import siann_un_tiau_tng_piau_im  # 台語音標轉台語音標
from mod_標音 import split_hong_im_hu_ho  # 分解漢字標音
from mod_標音 import split_tai_gi_im_piau  # 分解台語音標
from mod_標音 import tlpa_tng_han_ji_piau_im  # 台語音標轉漢字標音
from mod_標音 import tai_gi_im_piau_tng_un_bu, un_bu_tng_huan  # 韻母轉換
from p740_Phua_Im_Ji import PhuaImJi

# =========================================================================
# 常數定義
# =========================================================================
# 定義 Exit Code
EXIT_CODE_SUCCESS = 0  # 成功
EXIT_CODE_NO_FILE = 1  # 無法找到檔案
EXIT_CODE_INVALID_INPUT = 2  # 輸入錯誤
EXIT_CODE_PROCESS_FAILURE = 3  # 過程失敗
EXIT_CODE_UNKNOWN_ERROR = 99  # 未知錯誤

# =========================================================================
# 設定日誌
# =========================================================================
logging.basicConfig(
    filename='process_log.txt',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def logging_process_step(msg):
    print(msg)
    logging.info(msg)


# =========================================================================
# 程式函數
# =========================================================================
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
        #-----------------------------------------------：------------------
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
    han_ji_piau_im = siann_un_tiau_tng_piau_im(
        piau_im,
        piau_im_huat,
        siann_bu,
        un_bu,
        tiau_ho
    )
    return tai_gi_im_piau, han_ji_piau_im


def ca_han_ji_thak_im(wb, sheet_name='漢字注音', cell='V3', ue_im_lui_piat="白話音", han_ji_khoo="河洛話", db_name='Ho_Lok_Ue.db', module_name='mod_河洛話', function_name='han_ji_ca_piau_im'):
    """查漢字讀音：依【漢字】查找【台語音標】，並依指定之【標音方法】輸出【漢字標音】"""
    # 取得【漢字庫】工作表物件
    han_ji_koo_sheet = get_han_ji_khoo(wb)
    jin_kang_piau_im = get_han_ji_khoo(wb, sheet_name='人工標音字庫')
    khuat_ji_piau_sheet = get_han_ji_khoo(wb, sheet_name='缺字表')

    # 動態載入查找函數
    han_ji_ca_piau_im = load_module_function(module_name, function_name)

    # 連接指定資料庫
    conn = sqlite3.connect(db_name)
    cursor = conn.cursor()

    # 初始化 PiauIm 類別，産生標音物件
    piau_im = PiauIm(han_ji_khoo=han_ji_khoo)
    piau_im_huat = wb.names['標音方法'].refers_to_range.value
    # piau_im_huat = '方音符號'
    # phua_im_ji = PhuaImJi()

    # 選擇工作表
    sheet = wb.sheets[sheet_name]
    sheet.activate()
    sheet.range('A1').select()

    # 設定起始及結束的【列】位址（【第5列】、【第9列】、【第13列】等列）
    TOTAL_LINES = int(wb.names['每頁總列數'].refers_to_range.value)
    ROWS_PER_LINE = 4
    start_row = 5
    end_row = start_row + (TOTAL_LINES * ROWS_PER_LINE)
    line = 1

    # 設定起始及結束的【欄】位址（【D欄=4】到【R欄=18】）
    CHARS_PER_ROW = int(wb.names['每列總字數'].refers_to_range.value)
    start_col = 4
    end_col = start_col + CHARS_PER_ROW

    # 逐列處理作業
    EOF = False
    for row in range(start_row, end_row, ROWS_PER_LINE):
        # 若已到【結尾】或【超過總行數】，則跳出迴圈
        if EOF or line > TOTAL_LINES:
            break

        # 設定【作用儲存格】為列首
        Two_Empty_Cells = 0
        sheet.range((row, 1)).select()

        # 逐欄取出漢字處理
        for col in range(start_col, end_col):
            # 取得當前儲存格內含值
            han_ji_u_piau_im = False
            msg = ""
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
                            han_ji_piau_im = siann_un_tiau_tng_piau_im(
                                piau_im,
                                piau_im_huat,
                                siann,
                                un,
                                tiau
                            )
                            han_ji_u_piau_im = True

                        # 將人工輸入的【台語音標】置入【標音字庫】Dict
                        # phua_im_ji.ka_phua_im_ji(han_ji, tai_gi_im_piau)
                        maintain_han_ji_koo(sheet=jin_kang_piau_im,
                                            han_ji=han_ji,
                                            tai_gi=tai_gi_im_piau,
                                            show_msg=False)
                    else:               # 無人工輸入，則自【漢字庫】查找作業
                        # 查找【標音字庫】，確認是否有此漢字
                        # found = phua_im_ji.ca_phua_im_ji(han_ji)
                        tai_gi_im_piau = get_tai_gi_by_han_ji(jin_kang_piau_im, han_ji)
                        found = tai_gi_im_piau if tai_gi_im_piau else False
                        # 若【標音字庫】有此漢字
                        if found:
                            siann_bu, un_bu, tiau_ho = split_tai_gi_im_piau(found)
                            tai_gi_im_piau = siann_bu + un_bu + tiau_ho
                            han_ji_piau_im = siann_un_tiau_tng_piau_im(
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
                        # 若【標音字庫】無此漢字，則在資料庫中查找
                        else:
                            result = han_ji_ca_piau_im(cursor=cursor, han_ji=han_ji, ue_im_lui_piat=ue_im_lui_piat)
                            if not result:
                                maintain_han_ji_koo(sheet=khuat_ji_piau_sheet,
                                                    han_ji=han_ji,
                                                    tai_gi='',
                                                    show_msg=False)
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
                    maintain_han_ji_koo(sheet=han_ji_koo_sheet,
                                        han_ji=han_ji,
                                        tai_gi=tai_gi_im_piau,
                                        show_msg=False)
                    sheet.range((row - 1, col)).value = tai_gi_im_piau
                    sheet.range((row + 1, col)).value = han_ji_piau_im
                    msg = f"{han_ji}： [{tai_gi_im_piau}] /【{han_ji_piau_im}】"

            # 顯示處理進度
            col_name = xw.utils.col_name(col)   # 取得欄位名稱
            print(f"({row}, {col_name}) = {msg}")

            # 若讀到【換行】或【文字終結】，跳出逐欄取字迴圈
            if msg == "【換行】" or EOF:
                break

        # 每當處理一行 15 個漢字後，亦換到下一行
        print("\n")
        line += 1
        row += 4

    #----------------------------------------------------------------------
    # 作業處理用的 row 迴圈與 col 迴圈己終結
    #----------------------------------------------------------------------
    # 關閉資料庫連線
    conn.close()

    # 作業結束前處理
    wb.save()
    print("已完成【台語音標】和【漢字標音】標注工作。")
    return EXIT_CODE_SUCCESS



def update_han_ji_piau_im(wb, han_ji_koo_sheet_name='漢字庫', han_ji_piau_im_sheet_name='漢字注音'):
    """
    更新【漢字注音】表中【台語音標】儲存格的內容，依據【漢字庫】中的【校正】欄位進行更新。
    wb: Excel 活頁簿物件
    han_ji_koo_sheet_name: 【漢字庫】工作表名稱
    han_ji_zhu_yin_sheet_name: 【漢字注音】工作表名稱
    """
    # 取得工作表
    han_ji_koo_sheet = wb.sheets[han_ji_koo_sheet_name]
    han_ji_piau_im_sheet = wb.sheets[han_ji_piau_im_sheet_name]

    # 取得【漢字庫】表格範圍的所有資料
    data = han_ji_koo_sheet.range("A2").expand("table").value

    if data is None:
        print("【漢字庫】工作表無資料")
        return EXIT_CODE_INVALID_INPUT

    # 確保資料為 2D 列表
    if not isinstance(data[0], list):
        data = [data]

    # 將資料轉為字典格式，key: 漢字, value: (台語音標, 校正, 次數)
    han_ji_dict = {}
    for row in data:
        han_ji = row[0] or ""
        tai_gi_im_piau = row[1] or ""
        total_count = int(row[2]) if len(row) > 2 and isinstance(row[2], (int, float)) else 0
        corrected_tai_gi = row[3] if len(row) > 3 else ""  # 若無 D 欄資料則設為空字串

        if corrected_tai_gi and (corrected_tai_gi != tai_gi_im_piau):
            han_ji_dict[han_ji] = (tai_gi_im_piau, corrected_tai_gi, total_count)

    # 若無需更新的資料，結束函數
    if not han_ji_dict:
        print("【漢字庫】工作表中，【校正音標】欄，均未填入需更新之台語音標！")
        return EXIT_CODE_SUCCESS

    # 逐列處理【漢字注音】表
    TOTAL_LINES = int(wb.names['每頁總列數'].refers_to_range.value)
    ROWS_PER_LINE = 4
    CHARS_PER_ROW = int(wb.names['每列總字數'].refers_to_range.value)

    start_row = 5
    end_row = start_row + (TOTAL_LINES * ROWS_PER_LINE)
    start_col = 4
    end_col = start_col + CHARS_PER_ROW

    for row in range(start_row, end_row, ROWS_PER_LINE):
        for col in range(start_col, end_col):
            han_ji_cell = han_ji_piau_im_sheet.range((row, col))
            han_ji = han_ji_cell.value or ""

            if han_ji in han_ji_dict:
                _, corrected_tai_gi, total_count = han_ji_dict[han_ji]
                tai_gi_cell = han_ji_piau_im_sheet.range((row - 1, col))
                original_tai_gi = tai_gi_cell.value or ""

                # 更新多次，直到總數用完
                if corrected_tai_gi != original_tai_gi and total_count > 0:
                    tai_gi_cell.value = corrected_tai_gi  # 更新儲存格
                    han_ji_cell.color = (255, 255, 0)       # 將底色設為【黄色】
                    han_ji_cell.font.color = (255, 0, 0)    # 將文字顏色設為【紅色】

                    msg = f"({row}, {xw.utils.col_name(col)}) = {han_ji}，台語音標由【{original_tai_gi}】改為【{corrected_tai_gi}】"
                    print(msg)
                    total_count -= 1  # 減少剩餘更新次數

                    # 更新完畢後，減少【漢字庫】的總數
                    han_ji_koo_sheet.range(f"C{row + 1}").value = total_count
                    if total_count == 0:
                        print(f"漢字【{han_ji}】的更新次數已用完")

    print("【漢字注音】表的台語音標更新作業已完成")

    # 作業結束前處理
    wb.save()
    print("已完成【漢字注音】工作表，漢字之【台語音標】更新作業。")
    return EXIT_CODE_SUCCESS

