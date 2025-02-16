# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import logging
import os
import sqlite3
import sys
from pathlib import Path
from typing import Callable

# 載入第三方套件
import xlwings as xw
from dotenv import load_dotenv

# 載入自訂模組/函式
from a100_作業中活頁檔填入漢字 import process as fill_hanji_in_cells

# 載入自訂模組/函式
from mod_excel_access import (
    check_and_update_pronunciation,
    ensure_sheet_exists,
    get_value_by_name,
    strip_cell,
)
from mod_字庫 import JiKhooDict  # 漢字字庫物件
from mod_標音 import PiauIm  # 漢字標音物件
from mod_標音 import hong_im_tng_tai_gi_im_piau  # 方音符號轉台語音標
from mod_標音 import is_punctuation  # 是否為標點符號
from mod_標音 import split_hong_im_hu_ho  # 分解漢字標音
from mod_標音 import tlpa_tng_han_ji_piau_im  # 台語音標轉台語音標

# =========================================================================
# 載入環境變數
# =========================================================================
load_dotenv()

# 預設檔案名稱從環境變數讀取
DB_HO_LOK_UE = os.getenv('DB_HO_LOK_UE', 'Ho_Lok_Ue.db')
DB_KONG_UN = os.getenv('DB_KONG_UN', 'Kong_Un.db')

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
# 常數定義
# =========================================================================
# 定義 Exit Code
EXIT_CODE_SUCCESS = 0  # 成功
EXIT_CODE_FAILURE = 1  # 失敗
EXIT_CODE_NO_FILE = 1  # 無法找到檔案
EXIT_CODE_INVALID_INPUT = 2  # 輸入錯誤
EXIT_CODE_PROCESS_FAILURE = 3  # 過程失敗
EXIT_CODE_UNKNOWN_ERROR = 99  # 未知錯誤


# =========================================================================
# 本程式主要處理作業程序
# =========================================================================
def han_ji_ti_piau_im_ji_khoo(wb, position, han_ji: str, jin_kang_piau_im: str) -> bool:
    """
    檢查【漢字注音】工作表中，某【漢字】的【人工標音】欄位是否有填入？若有填入，則需檢查在【標音字庫】
    工作表，是否有重複狀況？若結果為有重複，則以【人工標音】之值代【標音字庫】工作表中的【校正音標】。
    """
    try:
        # 確保工作表存在
        han_ji_piau_im_sheet_name = '標音字庫'
        ensure_sheet_exists(wb, han_ji_piau_im_sheet_name)
        han_ji_piau_im_sheet = wb.sheets[han_ji_piau_im_sheet_name]
    except Exception as e:
        raise ValueError(f"無法找到或建立工作表 '{han_ji_piau_im_sheet_name}'：{e}")

    # 使用【漢字】的【人工標音】填入【標音字庫】工作表中的【校正音標】欄位
    return check_and_update_pronunciation(wb, han_ji, position, jin_kang_piau_im)


def jin_kang_piau_im_cu_han_ji_piau_im(wb, han_ji: str, jin_kang_piau_im: str, piau_im: PiauIm, piau_im_huat: str) -> str:
    """人工標音取【台語音標】"""
    cursor = piau_im.get_cursor()   # 取得【資料庫】系統之 cursor 物件
    jin_kang_piau_im_sheet = wb.sheets['人工標音字庫']  # 取得【人工標音字庫】工作表

    if '〔' in jin_kang_piau_im and '〕' in jin_kang_piau_im:
        # 將人工輸入的〔台語音標〕轉換成【方音符號】
        im_piau = jin_kang_piau_im.split('〔')[1].split('〕')[0]
        tai_gi_im_piau = im_piau
        # 依使用者指定之【標音方法】，將【台語音標】轉換成其所需之【漢字標音】
        han_ji_piau_im = tlpa_tng_han_ji_piau_im(
            piau_im=piau_im,
            piau_im_huat=piau_im_huat,
            tai_gi_im_piau=tai_gi_im_piau
        )
    elif '【' in jin_kang_piau_im and '】' in jin_kang_piau_im:
        # 將人工輸入的【方音符號】轉換成【台語音標】
        han_ji_piau_im = jin_kang_piau_im.split('【')[1].split('】')[0]
        siann, un, tiau = split_hong_im_hu_ho(han_ji_piau_im)
        # 依使用者指定之【標音方法】，將【台語音標】轉換成其所需之【漢字標音】
        tai_gi_im_piau = hong_im_tng_tai_gi_im_piau(
            siann=siann,
            un=un,
            tiau=tiau,
            cursor=cursor,
        )['台語音標']
    else:
        # 將人工輸入的【台語音標】，解構為【聲母】、【韻母】、【聲調】
        tai_gi_im_piau = jin_kang_piau_im
        # 依指定之【標音方法】，將【台語音標】轉換成其所需之【漢字標音】
        han_ji_piau_im = tlpa_tng_han_ji_piau_im(
            piau_im=piau_im,
            piau_im_huat=piau_im_huat,
            tai_gi_im_piau=tai_gi_im_piau
        )

    return han_ji_piau_im


def write_ji_khoo_dict_to_sheet(wb, sheet_name: str, ji_khoo_dict: JiKhooDict):
    """
    將 khuat_ji_piau 字典的資料寫回【缺字表】工作表。

    :param wb: Excel 活頁簿物件。
    :param sheet_name: 工作表名稱（例如「缺字表」）。
    :param khuat_ji_piau: 基於【缺字表】工作表建置的字典。
    """
    try:
        # 確保工作表存在
        ensure_sheet_exists(wb, sheet_name)
        sheet = wb.sheets[sheet_name]
    except Exception as e:
        raise ValueError(f"無法找到或建立工作表 '{sheet_name}'：{e}")

    # 清空工作表內容
    sheet.clear()

    # 寫入標題列
    headers = ["漢字", "總數", "台語音標", "校正音標", "座標"]
    sheet.range("A1").value = headers

    # 寫入字典內容
    data = []
    for han_ji, (total_count, tai_gi_im_piau, kenn_ziann_im_piau, coordinates) in ji_khoo_dict.items():
        coords_str = "; ".join([f"({row}, {col})" for row, col in coordinates])
        data.append([han_ji, total_count, tai_gi_im_piau, kenn_ziann_im_piau, coords_str])

    sheet.range("A2").value = data
    print(f"\n完成【{sheet_name}】工作表內容更新...")


def update_by_khuat_ji_piau(wb, sheet_name: str, piau_im: PiauIm, piau_im_huat: str):
    """
    將字典中的所有漢字資料寫入 Excel 的「漢字注音」工作表。

    :param wb: Excel 活頁簿物件。
    :param sheet_name: 工作表名稱（例如「漢字注音」）。
    """
    try:
        # 確保工作表存在
        piau_im_ji_khoo_sheet_name = '漢字注音'
        ensure_sheet_exists(wb, piau_im_ji_khoo_sheet_name)
        han_ji_piau_im_sheet = wb.sheets[piau_im_ji_khoo_sheet_name]

        # 依【工作表】內容建立【字庫字典】
        khuat_ji_piau_sheet_name = '缺字表'
        ji_khoo_dict = JiKhooDict.create_ji_khoo_dict_from_sheet(wb=wb, sheet_name=khuat_ji_piau_sheet_name)

        piau_im_ji_khoo_sheet_name = '標音字庫'
        piau_im_ji_khoo_dict = JiKhooDict.create_ji_khoo_dict_from_sheet(wb=wb, sheet_name=piau_im_ji_khoo_sheet_name)
    except Exception as e:
        raise ValueError(f"無法找到或建立工作表 '{sheet_name}'：{e}")

    # 遍歷字典中的每個漢字
    for han_ji, (total_count, tai_gi_im_piau, kenn_ziann_im_piau, coordinates) in ji_khoo_dict.items():
        # 若【校正音標】為空，則略過
        if total_count == 0:
            row_no, col_no = coordinates[0]
            print(f"（{row_no}, {xw.utils.col_name(col_no)}）= {han_ji}【{tai_gi_im_piau}】/【{kenn_ziann_im_piau}】：待校正【總數】為 {total_count}，略過！")
            continue
        # 遍歷每個座標
        for row, col in coordinates:
            # 將漢字和台語音標寫入指定座標
            han_ji_piau_im_sheet.range((row, col)).select()
            original_total_count = total_count
            # 取得【漢字注音】表中的【漢字】儲存格內容
            han_ji_cell = han_ji_piau_im_sheet.range((row, col))
            # 取得【漢字注音】表中的【台語音標】儲存格內容
            tai_gi_cell = han_ji_piau_im_sheet.range((row - 1, col))
            # tai_gi_im_piau = tai_gi_cell.value or ""
            original_tai_gi = tai_gi_im_piau
            # 取得【漢字注音】表中的【人工標音】儲存格內容
            jin_kang_piau_im_cell = han_ji_piau_im_sheet.range((row - 2, col))
            # jin_kang_piau_im = jin_kang_piau_im_cell.value or ""
            # 取得【漢字注音】表中的【漢字標音】儲存格
            han_ji_piau_im_cell = han_ji_piau_im_sheet.range((row + 1, col))
            han_ji_piau_im = han_ji_piau_im_cell.value or ""

            # sheet.range((row, col)).value = han_ji
            # 將【台語音標】寫入【漢字注音】表的【台語音標】儲存格
            han_ji_piau_im_sheet.range((row-1, col)).value = tai_gi_im_piau
            # 檢查是否符合更新條件：
            # 若【漢字標音】儲存格亦空缺，則用【台語音標】生成【漢字標音】
            if total_count > 0:
                han_ji_piau_im_cell.value = tlpa_tng_han_ji_piau_im(
                    piau_im=piau_im,
                    piau_im_huat=piau_im_huat,
                    tai_gi_im_piau=original_tai_gi
                )
                # 將【缺字表】已填入【台語音標】之資料，回填【標音字庫】工作表，補登紀錄
                piau_im_ji_khoo_dict.add_or_update_entry(
                    han_ji=han_ji,
                    tai_gi_im_piau=han_ji_piau_im_cell.value,
                    kenn_ziann_im_piau='N/A',
                    coordinates=(row, col)
                )
                # 減少剩餘更新次數，並同步回缺字表
                total_count -= 1
                # 每寫入一次，total_count 減 1
                ji_khoo_dict[han_ji][0] = total_count
                # 重置【漢字】儲存格的底色和文字顏色
                han_ji_cell.color = (255, 255, 0)       # 將底色設為【黄色】
                han_ji_cell.font.color = (255, 0, 0)    # 將文字顏色設為【紅色】
            # 顯示更新訊息
            print(f"({row}, {xw.utils.col_name(col)}) = {han_ji}：【{tai_gi_im_piau}】/【{kenn_ziann_im_piau}】"
                f"（原有：{original_total_count} 字；尚有 {total_count} 字待補上）")

    #-----------------------------------------------------------------------------------------
    # 作業結束前處理
    #-----------------------------------------------------------------------------------------
    # 將【缺字表】字典保存之資料，回填【缺字表】工作表
    # write_ji_khoo_dict_to_sheet(wb=wb, sheet_name=sheet_name, ji_khoo_dict=ji_khoo_dict)
    ji_khoo_dict.write_to_excel_sheet(wb=wb, sheet_name=khuat_ji_piau_sheet_name)
    piau_im_ji_khoo_dict.write_to_excel_sheet(wb=wb, sheet_name=piau_im_ji_khoo_sheet_name)
    # 顯示【漢字注音】工作表
    han_ji_piau_im_sheet.activate()
    han_ji_piau_im_sheet.range('A1').select()
    return EXIT_CODE_SUCCESS


def update_by_jin_kang_piau_im(wb, sheet_name: str, piau_im: PiauIm, piau_im_huat: str):
    """
    更新【漢字注音】表中【台語音標】儲存格的內容，依據【標音字庫】中的【校正音標】欄位進行更新，並將【校正音標】覆蓋至原【台語音標】。
    """
    try:
        # 確保工作表存在
        han_ji_piau_im_sheet_name = '漢字注音'
        ensure_sheet_exists(wb, han_ji_piau_im_sheet_name)
        han_ji_piau_im_sheet = wb.sheets[han_ji_piau_im_sheet_name]

        han_ji_khoo = wb.names['漢字庫'].refers_to_range.value
        piau_im = PiauIm(han_ji_khoo)

        han_ji_piau_im_huat = wb.names['標音方法'].refers_to_range.value
    except Exception as e:
        raise ValueError(f"無法找到或建立工作表 '{sheet_name}'：{e}")

    # 依據【人工標音字庫】及【標音字庫】工作表，建置【字庫】物件
    jin_kang_piau_im_sheet_name='人工標音字庫'
    jin_kang_piau_im_ji_khoo = JiKhooDict.create_ji_khoo_dict_from_sheet(wb, jin_kang_piau_im_sheet_name)

    piau_im_sheet_name='標音字庫'
    piau_im_ji_khoo = JiKhooDict.create_ji_khoo_dict_from_sheet(wb, piau_im_sheet_name)

    # 逐列處理【漢字注音】表
    TOTAL_LINES = int(wb.names['每頁總列數'].refers_to_range.value)
    CHARS_PER_ROW = int(wb.names['每列總字數'].refers_to_range.value)
    ROWS_PER_LINE = 4

    start_row = 5
    end_row = start_row + (TOTAL_LINES * ROWS_PER_LINE)
    start_col = 4
    end_col = start_col + CHARS_PER_ROW

    # 選擇工作表
    EOF = False # 是否已到達【漢字注音】表的結尾
    line = 1
    for row in range(start_row, end_row, ROWS_PER_LINE):
        # 設定【作用儲存格】為列首
        Empty_Cells_Total = 0
        han_ji_piau_im_sheet.activate()
        han_ji_piau_im_sheet.range((row, 1)).select()

        # 逐欄取出漢字處理
        for col in range(start_col, end_col):
            status = ""
            # 取得【漢字注音】表中的【漢字】儲存格內容
            han_ji_cell = han_ji_piau_im_sheet.range((row, col))
            tai_gi_cell = han_ji_piau_im_sheet.range((row - 1, col))
            jin_kang_piau_im_cell = han_ji_piau_im_sheet.range((row - 2, col))

            # 設定【漢字】儲存格的底色和文字顏色
            han_ji_cell.color = (255, 255, 255)       # 將底色設為【白色】
            han_ji_cell.font.color = (0, 0, 0)    # 將文字顏色設為【黑色】

            tai_gi_cell.color = (255, 255, 255)       # 將底色設為【白色】

            jin_kang_piau_im_cell.font.color = (255, 0, 0)    # 將文字顏色設為【紅色】
            jin_kang_piau_im_cell.font.size = 24

            # 依據【漢字】儲存格讀取之資料，進行處理作業
            if han_ji_cell.value == 'φ':
                EOF = True
                print(f"({row}, {xw.utils.col_name(col)}) = 《文章終止》")
                break
            elif han_ji_cell.value == '\n':
                print(f"({row}, {xw.utils.col_name(col)}) = 《換行》")
                break
            elif han_ji_cell.value == None or han_ji_cell.value == "":
                print(f"({row}, {xw.utils.col_name(col)}) = 《空格》")
                Empty_Cells_Total += 1
                if Empty_Cells_Total >= 2:
                    EOF = True
                    break
                else:
                    continue
            else:
                # 若不為【標點符號】，則以【漢字】處理
                if is_punctuation(han_ji_cell.value):
                    print(f"({row}, {xw.utils.col_name(col)}) = {han_ji_cell.value}：標點符號不處理")
                    continue
                else:
                    # 取得【漢字注音】表中的【漢字】儲存格內容
                    han_ji = han_ji_cell.value
                    # 取得【漢字注音】表中的【台語音標】儲存格內容
                    tai_gi_cell = han_ji_piau_im_sheet.range((row - 1, col))
                    tai_gi_im_piau = tai_gi_cell.value or ""
                    # 取得【漢字注音】表中的【漢字標音】儲存格
                    han_ji_piau_im_cell = han_ji_piau_im_sheet.range((row + 1, col))
                    han_ji_piau_im = han_ji_piau_im_cell.value or ""
                    # 取得【漢字注音】表中的【人工標音】儲存格內容
                    jin_kang_piau_im_cell = han_ji_piau_im_sheet.range((row - 2, col))
                    jin_kang_piau_im = strip_cell(jin_kang_piau_im_cell.value)

                    # ---------------------------------------------------------
                    # 確認【漢字】有【人工標音】時之處理作業
                    # ---------------------------------------------------------
                    if jin_kang_piau_im == None:        # 【漢字】沒用【人工標音】
                        # ---------------------------------------------------------
                        # 重置【漢字】儲存格的底色和文字顏色
                        # ---------------------------------------------------------
                        if han_ji_cell.color == (0, 255, 200) and jin_kang_piau_im_cell.value == tai_gi_cell.value:
                            jin_kang_piau_im_cell.value = ""
                            han_ji_cell.color = (255, 255, 255)       # 將底色設為【白色】
                            han_ji_cell.font.color = (0, 0, 0)    # 將文字顏色設為【黑色】
                    else:                               # 【漢字】以【人工標音】更正【程式自動標音】
                        # 在【漢字注音】工作表，為有【人工標音】之【漢字】儲存格做醒目標記
                        # jin_kang_piau_im_cell.value = ''
                        han_ji_cell.color = (255, 255, 0)       # 將底色設為【黄色】
                        han_ji_cell.font.color = (255, 0, 0)    # 將文字顏色設為【紅色】

                        if '〔' in jin_kang_piau_im and '〕' in jin_kang_piau_im:
                            # 將人工輸入的〔台語音標〕轉換成【方音符號】
                            im_piau = jin_kang_piau_im.split('〔')[1].split('〕')[0]
                            tai_gi_im_piau = im_piau
                            # 依使用者指定之【標音方法】，將【台語音標】轉換成其所需之【漢字標音】
                            han_ji_piau_im = tlpa_tng_han_ji_piau_im(
                                piau_im=piau_im,
                                piau_im_huat=piau_im_huat,
                                tai_gi_im_piau=tai_gi_im_piau
                            )
                        elif '【' in jin_kang_piau_im and '】' in jin_kang_piau_im:
                            # 將人工輸入的【方音符號】轉換成【台語音標】
                            han_ji_piau_im = jin_kang_piau_im.split('【')[1].split('】')[0]
                            siann, un, tiau = split_hong_im_hu_ho(han_ji_piau_im)
                            # 依使用者指定之【標音方法】，將【台語音標】轉換成其所需之【漢字標音】
                            han_ji_piau_im = piau_im.han_ji_piau_im_tng_huan(
                                piau_im=piau_im,    # 注音法物件
                                piau_im_huat=han_ji_piau_im_huat,
                                siann_bu=siann,
                                un_bu=un,
                                tiau_ho=tiau
                            )
                        else:
                            # 將人工輸入的【台語音標】，解構為【聲母】、【韻母】、【聲調】
                            tai_gi_im_piau = jin_kang_piau_im
                            # 依指定之【標音方法】，將【台語音標】轉換成其所需之【漢字標音】
                            han_ji_piau_im = tlpa_tng_han_ji_piau_im(
                                piau_im=piau_im,
                                piau_im_huat=piau_im_huat,
                                tai_gi_im_piau=tai_gi_im_piau
                            )

                        # 遇【漢字】具【人工標音】，於【人工標音字庫】工作表登錄一筆紀錄
                        jin_kang_piau_im_ji_khoo.add_or_update_entry(
                            han_ji=han_ji,
                            tai_gi_im_piau=tai_gi_im_piau,
                            kenn_ziann_im_piau='N/A',
                            coordinates=(row, col)
                        )
                        print(f"({row}, {xw.utils.col_name(col)}) = {han_ji_cell.value}：將【人工標音】{jin_kang_piau_im} 登錄至【人工標音字庫】工作表")

                        # 若【漢字】之標音有【人工標音】，則將【人工標音】填入【標音字庫】工作表之【校正音標】
                        # if jin_kang_piau_im and jin_kang_piau_im != tai_gi_im_piau:
                        if jin_kang_piau_im:
                            # 依據【漢字】的【人工標音】，更新【標音字庫】之【校正音標】欄位資料（新增或更新）
                            piau_im_ji_khoo.update_kau_ziang_im_piau(
                                han_ji=han_ji,
                                tai_gi_im_piau=tai_gi_im_piau,
                                kenn_ziann_im_piau=jin_kang_piau_im,
                                coordinates=(row, col)
                            )
                            print(f"({row}, {xw.utils.col_name(col)}) = {han_ji_cell.value}：將【人工標音】{jin_kang_piau_im} 填入【標音字庫】工作表之【校正音標】")

            # 每欄結束前處理作業
            msg_tail = f"：《{status}》" if status else f"：不處理"
            print(f"({row}, {xw.utils.col_name(col)}) = {han_ji}【{tai_gi_im_piau}】/【{han_ji_piau_im}】{msg_tail}")

        # 每列結束前處理作業
        print('\n')
        line += 1
        if EOF or line > TOTAL_LINES:
            # 將【人工標音字庫】及【標音字庫】工作表內容更新
            jin_kang_piau_im_ji_khoo.write_to_excel_sheet(wb, jin_kang_piau_im_sheet_name)
            piau_im_ji_khoo.write_to_excel_sheet(wb, piau_im_sheet_name)
            break

    #-------------------------------------------------------------------------------------
    # 作業結束前處理
    #-------------------------------------------------------------------------------------
    # 更新【標音字庫】工作表內容
    write_ji_khoo_dict_to_sheet(wb=wb, sheet_name='標音字庫', ji_khoo_dict=piau_im_ji_khoo)
    # 更新【人工標音字庫】工作表內容
    write_ji_khoo_dict_to_sheet(wb=wb, sheet_name=sheet_name, ji_khoo_dict=jin_kang_piau_im_ji_khoo)
    han_ji_piau_im_sheet.activate()
    han_ji_piau_im_sheet.range('A1').select()
    return EXIT_CODE_SUCCESS


def update_by_piau_im_ji_khoo(wb, sheet_name: str, piau_im: PiauIm, piau_im_huat: str):
    """
    依【標音字庫】中的【校正音標】欄位進行更新，並將【校正音標】覆蓋至原【台語音標】。
    """
    try:
        # 確保工作表存在
        han_ji_piau_im_sheet_name = '漢字注音'
        ensure_sheet_exists(wb, han_ji_piau_im_sheet_name)
        han_ji_piau_im_sheet = wb.sheets[han_ji_piau_im_sheet_name]

        # 依據【標音字庫】工作表，建置【字庫】物件
        piau_im_sheet_name='標音字庫'
        piau_im_ji_khoo_dict = JiKhooDict.create_ji_khoo_dict_from_sheet(wb, piau_im_sheet_name)
    except Exception as e:
        raise ValueError(f"無法找到或建立工作表 '{sheet_name}'：{e}")

    # 遍歷【標音字庫】工作表中的每個【漢字】
    for han_ji, (total_count, tai_gi_im_piau, kau_ziann_im_piau, coordinates) in piau_im_ji_khoo_dict.items():
        # 若【標音字庫】工作表的【校正音標】欄位為空，或為“N/A”，則略過
        if kau_ziann_im_piau == 'N/A' or kau_ziann_im_piau == '':
            row_no, col_no = coordinates[0]
            print(f"（{row_no}, {xw.utils.col_name(col_no)}）= {han_ji}【{tai_gi_im_piau}】/ 【{kau_ziann_im_piau}】：無需校正【台語音標】，略過！")
            continue
        # 遍歷每個座標
        for row, col in coordinates:
            original_total_count = total_count
            # 將漢字和台語音標寫入指定座標
            han_ji_piau_im_sheet.activate()
            han_ji_piau_im_sheet.range((row, col)).select()
            # 取得【漢字注音】表中的【漢字】儲存格內容
            han_ji_cell = han_ji_piau_im_sheet.range((row, col))
            # 使用【校正音標】，更新【漢字注音】工作表中的【台語音標】儲存格內容
            tai_gi_im_piau_cell = han_ji_piau_im_sheet.range((row - 1, col))
            tai_gi_im_piau = kau_ziann_im_piau
            # 取得【漢字注音】表中的【漢字標音】儲存格
            han_ji_piau_im_cell = han_ji_piau_im_sheet.range((row + 1, col))
            # han_ji_piau_im = han_ji_piau_im_cell.value or ""

            # 將【台語音標】寫入【漢字注音】表的【台語音標】儲存格
            # han_ji_piau_im_sheet.range((row-1, col)).value = tai_gi_im_piau
            tai_gi_im_piau_cell.value = kau_ziann_im_piau
            # 依指定之【標音方法】，將【台語音標】轉換成其所需之【漢字標音】
            han_ji_piau_im_cell.value = tlpa_tng_han_ji_piau_im(
                piau_im=piau_im,
                piau_im_huat=piau_im_huat,
                tai_gi_im_piau=tai_gi_im_piau
            )
            # 重置【漢字】儲存格的底色和文字顏色
            han_ji_cell.color = (0, 255, 255)       # 將底色設為【黄色】
            han_ji_cell.font.color = (255, 0, 0)    # 將文字顏色設為【紅色】
            # 減少剩餘更新次數，並同步回缺字表
            total_count -= 1
            # # 每寫入一次，total_count 減 1
            # ji_khoo_dict[han_ji][0] = total_count

            # 顯示更新訊息
            print(f"({row}, {xw.utils.col_name(col)}) = {han_ji}：【{tai_gi_im_piau}】/【{kau_ziann_im_piau}】"
                f"（原有：{original_total_count} 字；尚有 {total_count} 字待補上）")

    # 作業結束前處理
    # write_ji_khoo_dict_to_sheet(wb=wb, sheet_name=sheet_name, ji_khoo_dict=ji_khoo_dict)
    piau_im_ji_khoo_dict.write_to_excel_sheet(wb=wb, sheet_name=piau_im_sheet_name)
    han_ji_piau_im_sheet.range('A1').select()
    return EXIT_CODE_SUCCESS


def process(wb):
    """
    更新【漢字注音】表中【台語音標】儲存格的內容，依據【標音字庫】中的【校正音標】欄位進行更新，並將【校正音標】覆蓋至原【台語音標】。
    """
    try:
        # 連接【河洛話】資料庫，並建立 piau_im 物件
        han_ji_khoo_field = '漢字庫'
        han_ji_khoo_name = get_value_by_name(wb=wb, name=han_ji_khoo_field) # 取得【漢字庫】名稱：河洛話、廣韻
        piau_im = PiauIm(han_ji_khoo=han_ji_khoo_name)
        piau_im_huat = get_value_by_name(wb=wb, name='標音方法')    # 指定【台語音標】轉換成【漢字標音】的方法
        # 取得工作表
        han_ji_piau_im_sheet = wb.sheets['漢字注音']
        han_ji_piau_im_sheet.activate()
    except Exception as e:
        raise ValueError(f"找不到【漢字注音】工作表 ！'錯誤警示'：{e}")

    #-------------------------------------------------------------------------
    # 根據【缺字表】工作表更新【漢字注音】工作表中缺【台語音標】的【漢字】
    #-------------------------------------------------------------------------
    sheet_name = '缺字表'
    print('\n\n')
    print("======================================================================")
    print(f"使用【{sheet_name}】工作表中的【校正音標】，更正【台語音標】儲存格：")
    print("======================================================================")
    update_by_khuat_ji_piau(wb=wb,
                            sheet_name=sheet_name,
                            piau_im=piau_im,
                            piau_im_huat=piau_im_huat)
    print("\n使用【缺字表】之【台語音標】更新【台語音標】作業已完成！")
    #-------------------------------------------------------------------------
    # 根據【漢字注音】工作表之【人工標音】儲存格內容更新【台語音標】儲存格
    #-------------------------------------------------------------------------
    sheet_name = '人工標音字庫'
    print('\n\n')
    print("================================================================================")
    print(f"使用【漢字注音】工作表中的【人工標音】儲存格內容，更新【台語音標】：")
    print("================================================================================")
    update_by_jin_kang_piau_im(wb=wb,
                               sheet_name='人工標音字庫',
                               piau_im=piau_im,
                               piau_im_huat=piau_im_huat)
    print("\n使用【漢字注音】之【人工標音】更新【台語音標】作業已完成！")
    #-------------------------------------------------------------------------
    # 根據【標音字庫】工作表更新【漢字注音】工作表中的【台語音標】
    #-------------------------------------------------------------------------
    sheet_name = '標音字庫'
    print('\n\n')
    print("================================================================================")
    print(f"使用【{sheet_name}】工作表中的【校正音標】，更新【漢字注音】工作表中的【台語音標】：")
    print("================================================================================")
    update_by_piau_im_ji_khoo(wb=wb,
                              sheet_name=sheet_name,
                              piau_im=piau_im,
                              piau_im_huat=piau_im_huat)
    print("\n使用【標音字庫】之【校正音標】更新【台語音標】作業已完成！")
    #-------------------------------------------------------------------------
    # 作業結束前處理
    #-------------------------------------------------------------------------
    han_ji_piau_im_sheet.range('A1').select()
    print('\n\n')
    print("================================================================================")
    print("【漢字注音】表的【台語音標】更新作業已完成")
    print("================================================================================")

    logging_process_step(f"完成【作業程序】：更新漢字標音並同步【標音字庫】內容...")
    return EXIT_CODE_SUCCESS

# =========================================================================
# 程式主要作業流程
# =========================================================================
def main():
    # =========================================================================
    # (1) 取得專案根目錄。
    # =========================================================================
    current_file_path = Path(__file__).resolve()
    project_root = current_file_path.parent
    logging_process_step(f"專案根目錄為: {project_root}")

    # =========================================================================
    # (2) 若無指定輸入檔案，則獲取當前作用中的 Excel 檔案並另存新檔。
    # =========================================================================
    wb = None
    # 使用已打開且處於作用中的 Excel 工作簿
    try:
        # 嘗試獲取當前作用中的 Excel 工作簿
        wb = xw.apps.active.books.active
    except Exception as e:
        logging_process_step(f"發生錯誤: {e}")
        logging.error(f"無法找到作用中的 Excel 工作簿: {e}", exc_info=True)
        return EXIT_CODE_NO_FILE

    if not wb:
        logging_process_step("無法作業，因未無任何 Excel 檔案己開啟。")
        return EXIT_CODE_NO_FILE

    try:
        # =========================================================================
        # (3) 執行【處理作業】
        # =========================================================================
        result_code = process(wb)
        if result_code != EXIT_CODE_SUCCESS:
            logging_process_step("處理作業失敗，過程中出錯！")
            return result_code

    except Exception as e:
        print(f"執行過程中發生未知錯誤: {e}")
        logging.error(f"執行過程中發生未知錯誤: {e}", exc_info=True)
        return EXIT_CODE_UNKNOWN_ERROR

    finally:
        if wb:
            wb.save()
            # 是否關閉 Excel 視窗可根據需求決定
            # xw.apps.active.quit()  # 確保 Excel 被釋放資源，避免開啟殘留
            logging.info("釋放 Excel 資源，處理完成。")

    # 結束作業
    logging.info("作業成功完成！")
    return EXIT_CODE_SUCCESS


if __name__ == "__main__":
    exit_code = main()
    if exit_code == EXIT_CODE_SUCCESS:
        print("作業正常結束！")
    else:
        print(f"作業異常終止，錯誤代碼為: {exit_code}")
    sys.exit(exit_code)
