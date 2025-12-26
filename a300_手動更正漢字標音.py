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

from a301_以作用儲存格之人工標音更新標音字庫 import check_and_update_pronunciation

# 載入自訂模組/函式
from a310_缺字表修正後續作業 import process as update_khuat_ji_piau_by_jin_kang_piau_im
from a320_人工標音更正漢字自動標音 import process as update_by_jin_kang_piau_im
from a330_以標音字庫更新漢字注音工作表 import process as update_by_piau_im_ji_khoo
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
def process(wb):
    """
    更新【漢字注音】表中【台語音標】儲存格的內容，依據【標音字庫】中的【校正音標】欄位進行更新，並將【校正音標】覆蓋至原【台語音標】。
    """
    logging_process_step("<=========== 開始處理流程作業！==========>")
    try:
        # 取得工作表
        target_sheet_name = '漢字注音'
        ensure_sheet_exists(wb, target_sheet_name)
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
        logging_process_step(f"以【{sheet_name}】工作表之【校正音標】，更新【{target_sheet_name}】工作表之【台語音標】與【漢字標音】！")
        print('\n\n')
        print("=" * 100)
        print(f"使用【{sheet_name}】工作表的【校正音標】欄位，更新【{target_sheet_name}】工作表之【台語音標】、【漢字標音】：")
        print("=" * 100)
        # 將【缺字表】工作表中的【台語音標】儲存格內容，更新至【標音字庫】工作表中之【台語音標】及【校正音標】欄位
        # update_khuat_ji_piau(wb=wb)
        # 依據【缺字表】工作表紀錄，並參考【漢字注音】工作表在【人工標音】欄位的內容，更新【缺字表】工作表中的【校正音標】及【台語音標】欄位
        # 即使用者為【漢字】補入查找不到的【台語音標】時，若是在【缺字表】工作表中之【校正音標】直接填寫
        # 則應執行 a310*.py 程式；但使用者若是在【漢字注音】工作表中之【人工標音】欄位填寫，則應執行 a320*.py 程式
        # a300*.py 之本程式
        update_khuat_ji_piau_by_jin_kang_piau_im(wb=wb, sheet_name=sheet_name)
    except Exception as e:
        logging_exc_error(msg=f"處理【缺字表】作業異常！", error=e)
        return EXIT_CODE_PROCESS_FAILURE
    #-------------------------------------------------------------------------
    # 將【漢字注音】工作表，【漢字】填入【人工標音】內容，登錄至【人工標音字庫】及【標音字庫】工作表
    #-------------------------------------------------------------------------
    try:
        sheet_name = '人工標音字庫'
        logging_process_step(f"以【{sheet_name}】工作表之【校正音標】，更新【{target_sheet_name}】工作表之【台語音標】與【漢字標音】！")
        print('\n\n')
        print("=" * 100)
        print(f"使用【{sheet_name}】工作表的【校正音標】欄位，更新【{target_sheet_name}】工作表之【台語音標】、【漢字標音】：")
        print("=" * 100)
        update_by_jin_kang_piau_im(wb=wb, sheet_name=sheet_name )
    except Exception as e:
        logging_exc_error(msg=f"處理【漢字】之【人工標音】作業異常！", error=e)
        return EXIT_CODE_PROCESS_FAILURE
    #-------------------------------------------------------------------------
    # 根據【標音字庫】工作表，更新【漢字注音】工作表中的【台語音標】及【漢字標音】欄位
    #-------------------------------------------------------------------------
    try:
        sheet_name = '標音字庫'
        logging_process_step(f"以【{sheet_name}】工作表之【校正音標】，更新【{target_sheet_name}】工作表之【台語音標】與【漢字標音】！")
        print('\n\n')
        print("=" * 100)
        print(f"使用【{sheet_name}】工作表中的【校正音標】，更新【漢字注音】工作表中的【台語音標】：")
        print("=" * 100)
        update_by_piau_im_ji_khoo(wb, sheet_name=sheet_name)
    except Exception as e:
        logging_exc_error(msg=f"處理以【標音字庫】更新【漢字注音】工作表之作業，發生執行異常！", error=e)
        return EXIT_CODE_PROCESS_FAILURE
    #--------------------------------------------------------------------------
    # 結束作業
    #--------------------------------------------------------------------------
    han_ji_piau_im_sheet.activate()
    logging_process_step("<=========== 完成處理流程作業！==========>")

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

