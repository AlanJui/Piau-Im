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
from a330_以作用儲存格之人工標音更新標音字庫 import check_and_update_pronunciation
from mod_excel_access import ensure_sheet_exists, get_value_by_name, strip_cell
from mod_file_access import save_as_new_file
from mod_字庫 import JiKhooDict  # 漢字字庫物件
from mod_帶調符音標 import (
    cing_bo_iong_ji_bu,
    is_han_ji,
    kam_si_u_tiau_hu,
    tng_im_piau,
    tng_tiau_ho,
)
from mod_標音 import PiauIm  # 漢字標音物件
from mod_標音 import convert_tl_with_tiau_hu_to_tlpa  # 去除台語音標的聲調符號
from mod_標音 import is_punctuation  # 是否為標點符號
from mod_標音 import split_hong_im_hu_ho  # 分解漢字標音
from mod_標音 import tlpa_tng_han_ji_piau_im  # 台語音標轉台語音標

# =========================================================================
# 常數定義
# =========================================================================
# 定義 Exit Code
EXIT_CODE_SUCCESS = 0  # 成功
EXIT_CODE_NO_FILE = 1  # 無法找到檔案
EXIT_CODE_INVALID_INPUT = 2  # 輸入錯誤
EXIT_CODE_SAVE_FAILURE = 3  # 儲存失敗
EXIT_CODE_PROCESS_FAILURE = 10  # 過程失敗
EXIT_CODE_UNKNOWN_ERROR = 99  # 未知錯誤

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
from mod_logging import (
    init_logging,
    logging_exc_error,
    logging_exception,
    logging_process_step,
    logging_warning,
)

init_logging()

# =========================================================================
# 本程式主要處理作業程序
# =========================================================================

#---------------------------------------------------------------------------
# 此函式之功用，功能類似 a310*.py 之 update_khuat_ji_piau()　函式，但其作法有所不同。
# 但此函式則是將【缺字表】工作表之 table 資料讀入，再以 dict item() 方式，逐筆讀取資料進行更新。
#---------------------------------------------------------------------------
def update_khuat_ji_piau_by_jin_kang_piau_im(wb):
    """
    將字典中的所有漢字資料寫入 Excel 的「漢字注音」工作表。

    :param wb: Excel 活頁簿物件。
    :param sheet_name: 工作表名稱（例如「漢字注音」）。
    """
    # 取得本函式所需之【選項】參數
    try:
        han_ji_khoo = wb.names["漢字庫"].refers_to_range.value
        piau_im_huat = wb.names["標音方法"].refers_to_range.value
    except Exception as e:
        logging_exc_error("找不到作業所需之選項設定", e)
        return EXIT_CODE_INVALID_INPUT

    piau_im = PiauIm(han_ji_khoo=han_ji_khoo)

    # 依【缺字表】工作表內容建立【字庫字典】
    khuat_ji_piau_sheet_name = '缺字表'
    try:
        # khuat_ji_piau_sheet = wb.sheets[khuat_ji_piau_sheet_name]
        khuat_ji_khoo_dict = JiKhooDict.create_ji_khoo_dict_from_sheet(
            wb=wb, sheet_name=khuat_ji_piau_sheet_name)
    except Exception as e:
        logging_exc_error("無法為【缺字表】建【缺字表字典表格】", e)
        return EXIT_CODE_PROCESS_FAILURE

    # 確保【漢字漢音】工作表存在
    piau_im_ji_khoo_sheet_name = '漢字注音'
    try:
        ensure_sheet_exists(wb, piau_im_ji_khoo_sheet_name)
        han_ji_piau_im_sheet = wb.sheets[piau_im_ji_khoo_sheet_name]
    except Exception as e:
        raise ValueError(f"無法找到或建立【漢字注音】工作表：{e}")

    # 建置【標音字庫】工作表之字典表格
    piau_im_sheet_name = '標音字庫'
    try:
        piau_im_ji_khoo_dict = JiKhooDict.create_ji_khoo_dict_from_sheet(
            wb=wb, sheet_name=piau_im_sheet_name)
    except Exception as e:
        raise ValueError(f"無法找到或建立工作表：{e}")

    # 依【缺字表】工作表所建立之【字典表格】，遍歷【表格】每個字典查不到【音標】之【漢字】
    try:
        for han_ji, entry in khuat_ji_khoo_dict.items():
            tai_gi_im_piau = entry["tai_gi_im_piau"]
            kenn_ziann_im_piau = entry["kenn_ziann_im_piau"]
            coordinates = entry["coordinates"]

            # 遍歷【座標】欄位中每個座標，依【座標】所指向【漢字注音】工作表之儲存格，讀取【漢字】之【人工標音】
            for row, col in coordinates:
                # 取得【漢字注音】表中的【漢字】儲存格物件
                han_ji_cell = han_ji_piau_im_sheet.range((row, col))
                # 取得【漢字注音】表中的【人工標音】儲存格內容
                jin_kang_piau_im_cell = han_ji_piau_im_sheet.range((row - 2, col))

                # 取得【漢字注音】表中的【台語音標】儲存格內容
                tai_gi_cell = han_ji_piau_im_sheet.range((row - 1, col))
                han_ji_piau_im_cell = han_ji_piau_im_sheet.range((row + 1, col))

                # 如果【人工標音】為【帶調符音標】，則需確保轉換為【帶調號TLPA音標】
                jin_kang_piau_im = strip_cell(jin_kang_piau_im_cell.value)
                if not jin_kang_piau_im:
                    continue
                if tai_gi_im_piau == 'N/A' or tai_gi_im_piau == '':
                    continue
                elif kenn_ziann_im_piau == 'N/A' or kenn_ziann_im_piau == '':
                    # 若【缺字表】表格中【校正音標】欄位值為空，則略過
                    continue
                # 若取得之【人工標音】，為【帶調符音標】時，則需轉換為【帶調號TLPA音標】
                if kam_si_u_tiau_hu(jin_kang_piau_im):
                    jin_kang_im_piau = cing_bo_iong_ji_bu(jin_kang_piau_im_cell.value)
                    # 轉換成【帶調符TLPA音標】
                    tlpa_im_piau_u_tiau_hu = tng_im_piau(jin_kang_im_piau)
                    # 轉換成【帶調號TLPA音標】，並轉成小寫
                    tlpa_im_piau = tng_tiau_ho(tlpa_im_piau_u_tiau_hu).lower()
                else:
                    # tlpa_im_piau = jin_kang_piau_im_cell.value
                    tlpa_im_piau = jin_kang_piau_im

                # 依【人工標音】轉換【漢字標音】
                han_ji_piau_im = tlpa_tng_han_ji_piau_im(
                    piau_im=piau_im,
                    piau_im_huat=piau_im_huat,
                    tai_gi_im_piau=tlpa_im_piau
                )

                # 回填【缺字表】表格【校正音標】欄位
                tai_gi_im_piau = tlpa_im_piau
                kenn_ziann_im_piau = jin_kang_piau_im

                # 更新【漢字注音】工作表中【台語音標】、【漢字標音】儲存格內容
                tai_gi_cell.value = tai_gi_im_piau
                han_ji_piau_im_cell.value = han_ji_piau_im

                # ----- 新增程式邏輯：更新【標音字庫】 -----
                # Step 1: 在【標音字庫】搜尋該筆【漢字】+【台語音標】
                existing_entries = piau_im_ji_khoo_dict.ji_khoo_dict.get(han_ji, [])

                # 標記是否找到
                entry_found = False

                for existing_entry in existing_entries:
                    # Step 2: 若找到，移除該筆資料內的座標
                    if (row, col) in existing_entry["coordinates"]:
                        existing_entry["coordinates"].remove((row, col))
                    entry_found = True
                    break  # 找到即可離開迴圈

                # Step 3: 將此筆資料（校正音標為 'N/A'）於【標音字庫】底端新增
                piau_im_ji_khoo_dict.add_entry(
                    han_ji=han_ji,
                    tai_gi_im_piau=tai_gi_im_piau,
                    kenn_ziann_im_piau="N/A",  # 預設值
                    coordinates=(row, col)
                )

                # 重置【漢字】儲存格的底色和文字顏色
                han_ji_cell.color = (255, 255, 0)       # 將底色設為【黄色】
                han_ji_cell.font.color = (255, 0, 0)    # 將文字顏色設為【紅色】

                # 顯示更新訊息
                msg = f"{han_ji}：【{tai_gi_im_piau}】/【{kenn_ziann_im_piau}】<-- 【{jin_kang_im_piau}】"
                print(f"({row}, {col}) = {msg}")

    except Exception as e:
        logging_exception(msg=f"處理【漢字】補【台語音標】作業異常！", error=e)
        raise

    #-----------------------------------------------------------------------------------------
    # 作業結束前處理
    #-----------------------------------------------------------------------------------------
    try:
        # 將【缺字表】字典保存之資料，回填【缺字表】工作表
        khuat_ji_khoo_dict.write_to_excel_sheet(wb=wb, sheet_name=khuat_ji_piau_sheet_name)
        piau_im_ji_khoo_dict.write_to_excel_sheet(wb=wb, sheet_name=piau_im_sheet_name)
    except Exception as e:
        logging_exception(msg=f"將【字典】存放之資料，更新工作表作業異常！", error=e)
        raise
    # 顯示【漢字注音】工作表
    han_ji_piau_im_sheet.activate()
    han_ji_piau_im_sheet.range('A1').select()

    return EXIT_CODE_SUCCESS


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


def jin_kang_piau_im_cu_han_ji_piau_im(wb, jin_kang_piau_im: str, piau_im: PiauIm, piau_im_huat: str):
    """人工標音取【台語音標】"""

    if '〔' in jin_kang_piau_im and '〕' in jin_kang_piau_im:
        # 將人工輸入的〔台語音標〕轉換成【方音符號】
        im_piau = jin_kang_piau_im.split('〔')[1].split('〕')[0]
        tai_gi_im_piau = convert_tl_with_tiau_hu_to_tlpa(im_piau)
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
        tai_gi_im_piau = piau_im.hong_im_tng_tai_gi_im_piau(
            siann=siann,
            un=un,
            tiau=tiau)['台語音標']
    else:
        # 將人工輸入的【台語音標】，解構為【聲母】、【韻母】、【聲調】
        tai_gi_im_piau = convert_tl_with_tiau_hu_to_tlpa(jin_kang_piau_im)
        # 依指定之【標音方法】，將【台語音標】轉換成其所需之【漢字標音】
        han_ji_piau_im = tlpa_tng_han_ji_piau_im(
            piau_im=piau_im,
            piau_im_huat=piau_im_huat,
            tai_gi_im_piau=tai_gi_im_piau
        )

    return tai_gi_im_piau, han_ji_piau_im


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

    try:
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
                msg = ""
                if han_ji_cell.value == 'φ':
                    EOF = True
                    msg = f"《文章終止》"
                    break
                elif han_ji_cell.value == '\n':
                    msg = f"《換行》"
                    break
                # 若不為【標點符號】，則以【漢字】處理
                elif not is_han_ji(han_ji_cell.value):
                    str_value = str(han_ji_cell.value).strip()
                    # ✅ 若為全形／半形標點符號
                    if is_punctuation(str_value):
                        msg = f"{han_ji_cell.value}【標點符號】"
                    elif isinstance(str_value, float) and han_ji_cell.value.is_integer():
                        han_ji_cell.value = str(int(han_ji_cell.value))
                        msg = f"{han_ji_cell.value}【英/數半形字元】"
                    elif str_value == None or str_value == "":  # 若儲存格內無值
                        if Empty_Cells_Total == 0:
                            Empty_Cells_Total += 1
                        elif Empty_Cells_Total == 1:
                            EOF = True
                        msg = "【空白】"    # 表【儲存格】未填入任何字/符，不同於【空白】字元
                else:
                    # 取得【漢字注音】表中的【漢字】儲存格內容
                    han_ji = han_ji_cell.value
                    msg = f"{han_ji}"
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
                            msg += f"無【人工標音】"
                    else:                               # 【漢字】以【人工標音】更正【程式自動標音】
                        # 在【漢字注音】工作表，為有【人工標音】之【漢字】儲存格做醒目標記
                        han_ji_cell.color = (255, 255, 0)       # 將底色設為【黄色】
                        han_ji_cell.font.color = (255, 0, 0)    # 將文字顏色設為【紅色】

                        # 依據【人工標音】取【台語音標】及【漢字標音】
                        kenn_ziann_im_piau, han_ji_piau_im = jin_kang_piau_im_cu_han_ji_piau_im(
                            wb=wb,
                            jin_kang_piau_im=jin_kang_piau_im,
                            piau_im=piau_im,
                            piau_im_huat=piau_im_huat
                        )

                        # 依據【漢字】的【人工標音】，更新【漢字注音】工作表之【台語音標】及轉換取而得之【漢字標音】
                        tai_gi_cell.value = kenn_ziann_im_piau
                        han_ji_piau_im_cell.value = han_ji_piau_im

                        # 遇【漢字】具【人工標音】，於【人工標音字庫】工作表登錄一筆紀錄
                        jin_kang_piau_im_ji_khoo.add_or_update_entry(
                            han_ji=han_ji,
                            tai_gi_im_piau=tai_gi_im_piau,
                            kenn_ziann_im_piau=kenn_ziann_im_piau,
                            coordinates=(row, col)
                        )

                        # 若【漢字】之標音有【人工標音】，則將【人工標音】填入【標音字庫】工作表之【校正音標】
                        # if jin_kang_piau_im and jin_kang_piau_im != tai_gi_im_piau:
                        if jin_kang_piau_im:
                            # 依據【漢字】的【人工標音】，更新【標音字庫】之【校正音標】欄位資料（新增或更新）
                            piau_im_ji_khoo.update_kau_ziang_im_piau(
                                han_ji=han_ji,
                                tai_gi_im_piau=tai_gi_im_piau,
                                kenn_ziann_im_piau=kenn_ziann_im_piau,
                                coordinates=(row, col)
                            )

                            # 每欄結束前處理作業
                            msg += f"：[{tai_gi_im_piau}]/[{han_ji_piau_im}]【人工標音】"
                print(f"【{xw.utils.col_name(col)}{row}】({row}, {col}) = {msg}")

            # 每列結束前處理作業
            print('\n')
            line += 1
            if EOF or line > TOTAL_LINES: break
    except Exception as e:
        logging_exception(msg=f"處理【人工標音】作業異常！", error=e)
        raise

    #-------------------------------------------------------------------------------------
    # 作業結束前處理
    #-------------------------------------------------------------------------------------
    try:
        # 更新【標音字庫】工作表內容
        piau_im_ji_khoo.write_to_excel_sheet(wb, piau_im_sheet_name)
        # 更新【人工標音字庫】工作表內容
        jin_kang_piau_im_ji_khoo.write_to_excel_sheet(wb, jin_kang_piau_im_sheet_name)
    except Exception as e:
        logging_exception(msg=f"將【字典】存放之資料，更新工作表作業異常！", error=e)
        raise
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

    try:
        # 遍歷【標音字庫】工作表中的每個【漢字】
        # for han_ji, entries in piau_im_ji_khoo_dict.items():
        for han_ji, entries in piau_im_ji_khoo_dict.ji_khoo_dict.items():
            if not isinstance(entries, list):
                continue
            for entry in entries:
                if not isinstance(entry, dict):
                    continue
                tai_gi_im_piau = entry.get("tai_gi_im_piau", "")
                kau_ziann_im_piau = entry.get("kenn_ziann_im_piau", "")
                coordinates = entry.get("coordinates", [])

                # 若校正音標為空或 N/A，略過
                if (not kau_ziann_im_piau or kau_ziann_im_piau == 'N/A') or \
                    (not tai_gi_im_piau or tai_gi_im_piau == 'N/A'):
                    row_no, col_no = coordinates[0]
                    msg = f"{han_ji} [{tai_gi_im_piau}] / [{kau_ziann_im_piau}]【略過】"
                    print(f"({row_no}, {col_no}) = {msg}")
                    continue

                for row, col in coordinates:
                    # 將【台語音標】寫入指定之【座標】
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
                    # tai_gi_im_piau_cell.value = kau_ziann_im_piau
                    tai_gi_im_piau_cell.value = tai_gi_im_piau
                    # 依指定之【標音方法】，將【台語音標】轉換成其所需之【漢字標音】
                    han_ji_piau_im_cell.value = tlpa_tng_han_ji_piau_im(
                        piau_im=piau_im,
                        piau_im_huat=piau_im_huat,
                        tai_gi_im_piau=tai_gi_im_piau
                    )
                    # 重置【漢字】儲存格的底色和文字顏色
                    han_ji_cell.color = None            # 將底色設為【無填滿】
                    han_ji_cell.font.color = (0, 0, 0)  # 將文字顏色設為【自動】（黑色）
                    # # 減少剩餘更新次數，並同步回缺字表
                    # counter -= 1
                    # ji_khoo_dict[han_ji][0] = total_count

                    # 顯示更新訊息
                    # msg = f"{han_ji} [{tai_gi_im_piau}] / [{kau_ziann_im_piau}]（原有：{total_count} 字；尚有 {counter} 字待補上）"
                    msg = f"{han_ji} [{tai_gi_im_piau}] / [{kau_ziann_im_piau}]（自【標音字庫】回填）"
                    print(f"({row}, {col}) = {msg}")

    except Exception as e:
        logging_exception(msg=f"使用【標音字庫】之【校正音標】，改正【漢字注音】之【台語音標】作業異常！", error=e)
        raise

    #-------------------------------------------------------------------------------------
    # 作業結束前處理
    #-------------------------------------------------------------------------------------
    try:
        # write_ji_khoo_dict_to_sheet(wb=wb, sheet_name=sheet_name, ji_khoo_dict=ji_khoo_dict)
        piau_im_ji_khoo_dict.write_to_excel_sheet(wb=wb, sheet_name=piau_im_sheet_name)
    except Exception as e:
        logging_exception(msg=f"將【字典】存放之資料，更新工作表作業異常！", error=e)
        raise

    han_ji_piau_im_sheet.range('A1').select()
    return EXIT_CODE_SUCCESS


def process(wb):
    """
    更新【漢字注音】表中【台語音標】儲存格的內容，依據【標音字庫】中的【校正音標】欄位進行更新，並將【校正音標】覆蓋至原【台語音標】。
    """
    logging_process_step("<----------- 作業開始！---------->")
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
        logging_exc_error(msg=f"找不到【漢字注音】工作表 ！", error=e)
        return EXIT_CODE_PROCESS_FAILURE
    logging_process_step(f"已完成作業所需之初始化設定！")

    #-------------------------------------------------------------------------
    # 將【缺字表】工作表，已填入【台語音標】之資料，登錄至【標音字庫】工作表
    # 使用【缺字表】工作表中的【校正音標】，更正【漢字注音】工作表中之【台語音標】、【漢字標音】；
    # 並依【缺字表】工作表中的【台語音標】儲存格內容，更新【標音字庫】工作表中之【台語音標】及【校正音標】欄位
    #-------------------------------------------------------------------------
    try:
        sheet_name = '缺字表'
        print('\n\n')
        print("======================================================================")
        print(f"使用【{sheet_name}】工作表中的【校正音標】，更正【台語音標】儲存格：")
        print("======================================================================")
        # 將【缺字表】工作表中的【台語音標】儲存格內容，更新至【標音字庫】工作表中之【台語音標】及【校正音標】欄位
        # update_khuat_ji_piau(wb=wb)
        # 依據【缺字表】工作表紀錄，並參考【漢字注音】工作表在【人工標音】欄位的內容，更新【缺字表】工作表中的【校正音標】及【台語音標】欄位
        # 即使用者為【漢字】補入查找不到的【台語音標】時，若是在【缺字表】工作表中之【校正音標】直接填寫
        # 則應執行 a310*.py 程式；但使用者若是在【漢字注音】工作表中之【人工標音】欄位填寫，則應執行 a320*.py 程式
        # a300*.py 之本程式
        update_khuat_ji_piau_by_jin_kang_piau_im(wb=wb)
    except Exception as e:
        logging_exc_error(msg=f"處理【缺字表】作業異常！", error=e)
        return EXIT_CODE_PROCESS_FAILURE
    logging_process_step(f"【漢字庫】缺漏之【台語音標】，己填入【標音字庫】之【校正音標】！")
    #-------------------------------------------------------------------------
    # 將【漢字注音】工作表，【漢字】填入【人工標音】內容，登錄至【人工標音字庫】及【標音字庫】工作表
    #-------------------------------------------------------------------------
    try:
        sheet_name = '人工標音字庫'
        print('\n\n')
        print("================================================================================")
        print(f"使用【漢字注音】工作表中的【人工標音】儲存格內容，更新【台語音標】：")
        print("================================================================================")
        update_by_jin_kang_piau_im(wb=wb,
                                sheet_name='人工標音字庫',
                                piau_im=piau_im,
                                piau_im_huat=piau_im_huat)
    except Exception as e:
        logging_exc_error(msg=f"處理【漢字】之【人工標音】作業異常！", error=e)
        return EXIT_CODE_PROCESS_FAILURE
    logging_process_step(f"以【人工標音】更正之【台語音標】，已填入【標音字庫】之【校正音標】！")
    #-------------------------------------------------------------------------
    # 根據【標音字庫】工作表，更新【漢字注音】工作表中的【台語音標】及【漢字標音】欄位
    #-------------------------------------------------------------------------
    try:
        sheet_name = '標音字庫'
        print('\n\n')
        print("================================================================================")
        print(f"使用【{sheet_name}】工作表中的【校正音標】，更新【漢字注音】工作表中的【台語音標】：")
        print("================================================================================")
        update_by_piau_im_ji_khoo(wb=wb,
                                sheet_name=sheet_name,
                                piau_im=piau_im,
                                piau_im_huat=piau_im_huat)
    except Exception as e:
        logging_exc_error(msg=f"處理以【標音字庫】更新【漢字注音】工作表之作業，發生執行異常！", error=e)
        return EXIT_CODE_PROCESS_FAILURE
    logging_process_step(f"以【標音字庫】之【校正音標】，更新【漢字注音】工作表之【台語音標】！")
    #--------------------------------------------------------------------------
    # 結束作業
    #--------------------------------------------------------------------------
    han_ji_piau_im_sheet.range('A1').select()
    logging_process_step("<----------- 作業結束！---------->")

    #--------------------------------------------------------------------------
    # 依【漢字】之【台語音標】，轉換成【漢字標音】
    #--------------------------------------------------------------------------
    # han_ji_piau_im(wb)

    return EXIT_CODE_SUCCESS

# =========================================================================
# 程式主要作業流程
# =========================================================================
def main():
    # =========================================================================
    # (0) 程式初始化
    # =========================================================================
    # 取得專案根目錄。
    current_file_path = Path(__file__).resolve()
    project_root = current_file_path.parent
    # 取得程式名稱
    # program_file_name = current_file_path.name
    program_name = current_file_path.stem

    # =========================================================================
    # (1) 開始執行程式
    # =========================================================================
    logging_process_step(f"《========== 程式開始執行：{program_name} ==========》")
    logging_process_step(f"專案根目錄為: {project_root}")

    # =========================================================================
    # (2) 設定【作用中活頁簿】：偵測及獲取 Excel 已開啟之活頁簿檔案。
    # =========================================================================
    wb = None
    # 取得【作用中活頁簿】
    try:
        wb = xw.apps.active.books.active    # 取得 Excel 作用中的活頁簿檔案
    except Exception as e:
        print(f"發生錯誤: {e}")
        logging.error(f"無法找到作用中的 Excel 工作簿: {e}", exc_info=True)
        return EXIT_CODE_NO_FILE

    # 若無法取得【作用中活頁簿】，則因無法繼續作業，故返回【作業異常終止代碼】結束。
    if not wb:
        return EXIT_CODE_NO_FILE

    # =========================================================================
    # (3) 執行【處理作業】
    # =========================================================================
    try:
        result_code = process(wb)
        if result_code != EXIT_CODE_SUCCESS:
            msg = f"程式異常終止：{program_name}"
            logging_exc_error(msg=msg, error=e)
            return EXIT_CODE_PROCESS_FAILURE

    except Exception as e:
        msg = f"程式異常終止：{program_name}"
        logging_exc_error(msg=msg, error=e)
        return EXIT_CODE_UNKNOWN_ERROR

    finally:
        #--------------------------------------------------------------------------
        # 儲存檔案
        #--------------------------------------------------------------------------
        try:
            # 要求畫面回到【漢字注音】工作表
            wb.sheets['漢字注音'].activate()
            # 儲存檔案
            file_path = save_as_new_file(wb=wb)
            if not file_path:
                logging_exc_error(msg="儲存檔案失敗！", error=e)
                return EXIT_CODE_SAVE_FAILURE    # 作業異當終止：無法儲存檔案
            else:
                logging_process_step(f"儲存檔案至路徑：{file_path}")
        except Exception as e:
            logging_exc_error(msg="儲存檔案失敗！", error=e)
            return EXIT_CODE_SAVE_FAILURE    # 作業異當終止：無法儲存檔案

        # if wb:
        #     xw.apps.active.quit()  # 確保 Excel 被釋放資源，避免開啟殘留

    # =========================================================================
    # 結束程式
    # =========================================================================
    logging_process_step(f"《========== 程式終止執行：{program_name} ==========》")
    return EXIT_CODE_SUCCESS    # 作業正常結束


if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)

