# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import logging
import os

# import re
# import sqlite3
import sys
from datetime import datetime
from pathlib import Path

# 載入第三方套件
import xlwings as xw
from dotenv import load_dotenv

# 載入自訂模組/函式
# from a320_人工標音更正漢字自動標音 import jin_kang_piau_im_cu_han_ji_piau_im
from mod_excel_access import (
    convert_to_excel_address,
    excel_address_to_row_col,
    get_active_cell,
    get_active_cell_address,
    get_active_cell_info,
    get_line_no_by_row,
    get_row_by_line_no,
    get_row_col_from_coordinate,
    get_value_by_name,
)
from mod_字庫 import JiKhooDict  # 漢字字庫物件
from mod_標音 import (
    PiauIm,
    convert_tl_with_tiau_hu_to_tlpa,
    split_hong_im_hu_ho,
    tlpa_tng_han_ji_piau_im,
)

# from mod_標音 import convert_tl_with_tiau_hu_to_tlpa  # 去除台語音標的聲調符號
# from mod_標音 import is_punctuation  # 是否為標點符號
# from mod_標音 import split_hong_im_hu_ho  # 分解漢字標音
# from mod_標音 import tlpa_tng_han_ji_piau_im  # 漢字標音物件

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
# 作業程序
# =========================================================================
def jin_kang_piau_im_cu_han_ji_piau_im(wb, jin_kang_piau_im: str, piau_im: PiauIm, piau_im_huat: str):
    """
    取人工標音【台語音標】
    """

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


# =============================================================================
# 作業主流程
# =============================================================================

def process(wb, source_sheet_name='漢字注音', target_sheet_name='人工標音字庫'):
    """
    作業流程：
    1. 取得當前 Excel 作用儲存格 (漢字、座標)
    2. 計算【人工標音】位置與值
    3. 查詢【標音字庫】確認該座標是否已登錄
    4. 若【標正音標】為 'N/A'，則更新為【人工標音】
    """

    try:
        #----------------------------------------------------------------------
        # 作業前置處理
        #----------------------------------------------------------------------
        # 建置 PiauIm 物件，供作漢字拼音轉換作業
        piau_im_huat = get_value_by_name(wb=wb, name='標音方法')    # 指定【台語音標】轉換成【漢字標音】的方法
        han_ji_khoo_name = get_value_by_name(wb=wb, name='漢字庫')
        piau_im = PiauIm(han_ji_khoo=han_ji_khoo_name)            # 指定漢字自動查找使用的【漢字庫】

        # 建置【標音字庫】工作表之【查詢資料表】
        piau_im_sheet_name = '標音字庫'
        piau_im_ji_khoo = JiKhooDict.create_ji_khoo_dict_from_sheet(
            wb=wb,
            sheet_name=piau_im_sheet_name)

        # 建置【人工標音字庫】工作表之【查詢資料表】
        jin_kang_piau_im_sheet_name=target_sheet_name
        jin_kang_piau_im_ji_khoo = JiKhooDict.create_ji_khoo_dict_from_sheet(
            wb=wb,
            sheet_name=jin_kang_piau_im_sheet_name)

        # 指定【漢字注音】工作表為【作用工作表】
        sheet = wb.sheets[source_sheet_name]
        sheet.activate()

        #----------------------------------------------------------------------
        # 取得【作用儲存格】
        #----------------------------------------------------------------------
        source_sheet = wb.sheets[source_sheet_name]
        active_cell_address = get_active_cell_address()
        row, col = excel_address_to_row_col(active_cell_address)
        current_line_no = get_line_no_by_row(current_row_no=row)  # 計算行號
        jin_kang_piau_im_row, tai_gi_im_piau_row, han_ji_row, han_ji_piau_im_row = get_row_by_line_no(current_line_no)

        han_ji = source_sheet.range((han_ji_row, col)).value
        jin_kang_piau_im = source_sheet.range((jin_kang_piau_im_row, col)).value
        tai_gi_im_piau = source_sheet.range((tai_gi_im_piau_row, col)).value
        han_ji_piau_im = source_sheet.range((han_ji_piau_im_row, col)).value
        han_ji_position = (han_ji_row, col)
        han_ji_cell = source_sheet.range((han_ji_row, col))

        print(f"📌 作用儲存格：{active_cell_address} ==> 座標：{han_ji_position}")
        print(f"📌 漢字：{han_ji}")
        print(f"📌 人工標音：{jin_kang_piau_im}，台語音標：{tai_gi_im_piau}，漢字標音：{han_ji_piau_im}")

        #----------------------------------------------------------------------
        # 自【漢字注音】工作表之【作用儲存格】取得【人工標音】
        #----------------------------------------------------------------------
        tai_gi_im_piau, han_ji_piau_im = jin_kang_piau_im_cu_han_ji_piau_im(
            wb=wb,
            jin_kang_piau_im=jin_kang_piau_im,
            piau_im=piau_im,
            piau_im_huat=piau_im_huat)

        # 將【台語音標】和【漢字標音】寫入儲存格
        han_ji_cell.offset(-1, 0).value = tai_gi_im_piau      # 台語音標
        han_ji_cell.offset(+1, 0).value = han_ji_piau_im      # 漢字標音
        msg = f"{han_ji}： [{jin_kang_piau_im}] / [{tai_gi_im_piau}] /【{han_ji_piau_im}】"
        print(f"✅ 已更新儲存格：{active_cell_address}，內容為：{msg}")

        # 【標音字庫】添加或更新【漢字】及【台語音標】資料
        jin_kang_piau_im_ji_khoo.add_entry(
            han_ji=han_ji,
            tai_gi_im_piau=tai_gi_im_piau,
            hau_ziann_im_piau=jin_kang_piau_im,
            coordinates=(row, col)
        )

        #----------------------------------------------------------------------
        # 作業結束前處理
        #----------------------------------------------------------------------
        # 將【標音字庫】、【人工標音字庫】，寫入 Excel 工作表
        piau_im_ji_khoo.write_to_excel_sheet(wb=wb, sheet_name=piau_im_sheet_name)
        jin_kang_piau_im_ji_khoo.write_to_excel_sheet(wb=wb, sheet_name=jin_kang_piau_im_sheet_name)

        logging_process_step("已完成【台語音標】和【漢字標音】標注工作。")
        return EXIT_CODE_SUCCESS
    except Exception as e:
        # 你可以在這裡加上紀錄或處理，例如:
        logging.exception(f"自動為【漢字】查找【台語音標】作業，發生例外！\n{e}")
        # 再次拋出異常，讓外層函式能捕捉
        raise


# =============================================================================
# 程式主流程
# =============================================================================
def main():
    # =========================================================================
    # 開始作業
    # =========================================================================
    logging.info("作業開始")

    # =========================================================================
    # (1) 取得專案根目錄。
    # =========================================================================
    current_file_path = Path(__file__).resolve()
    project_root = current_file_path.parent
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
            logging_process_step("作業異常終止！")
            return result_code

    except Exception as e:
        print(f"作業過程發生未知的異常錯誤: {e}")
        logging.error(f"作業過程發生未知的異常錯誤: {e}", exc_info=True)
        return EXIT_CODE_UNKNOWN_ERROR

    finally:
        if wb:
            # xw.apps.active.quit()  # 確保 Excel 被釋放資源，避免開啟殘留
            logging.info("處理作業結束！")

    # =========================================================================
    # 結束作業
    # =========================================================================
    return EXIT_CODE_SUCCESS


def ut01(wb):
    # 作業流程：獲取當前作用中的 Excel 儲存格
    sheet_name, cell_address = get_active_cell(wb)
    print(f"✅ 目前作用中的儲存格：{sheet_name} 工作表 -> {cell_address}")

    # 將 Excel 儲存格地址轉換為 (row, col) 格式
    row, col = excel_address_to_row_col(cell_address)
    print(f"📌 Excel 位址 {cell_address} 轉換為 (row, col): ({row}, {col})")

    return EXIT_CODE_SUCCESS


if __name__ == "__main__":
    exit_code = main()
    if exit_code == EXIT_CODE_SUCCESS:
        print("程式正常完成！")
    else:
        print(f"程式異常終止，錯誤代碼為: {exit_code}")
    sys.exit(exit_code)
