# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import logging
import os
import sqlite3
import sys
from pathlib import Path

# 載入第三方套件
import xlwings as xw
from dotenv import load_dotenv

# 載入自訂模組
from mod_excel_access import ensure_sheet_exists, get_value_by_name
from mod_file_access import save_as_new_file
from mod_字庫 import JiKhooDict  # 漢字字庫物件
from mod_帶調符音標 import is_han_ji
from mod_標音 import PiauIm  # 漢字標音物件
from mod_標音 import ca_ji_kiat_ko_tng_piau_im  # 查字典得台語音標及漢字標音
from mod_標音 import is_punctuation  # 是否為標點符號
from mod_標音 import split_tai_gi_im_piau  # 分解台語音標

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
)

init_logging()


# =========================================================================
# 程式區域函式
# =========================================================================

def han_ji_piau_im(wb, sheet_name: str = '漢字注音'):
    """
    # 將【漢字注音】表中的【漢字】，依【台語音標】轉換成【漢字標音】
    """
    han_ji_khoo_name = wb.names['漢字庫'].refers_to_range.value # 取得【漢字庫】名稱：河洛話、廣韻
    ue_im_lui_piat = wb.names['語音類型'].refers_to_range.value # 取得【漢字庫】名稱：河洛話、廣韻
    db_name = 'Ho_Lok_Ue.db' if han_ji_khoo_name == '河洛話' else 'Kong_Un.db'

    try:
        # 確保工作表存在
        ensure_sheet_exists(wb, sheet_name)
        han_ji_piau_im_sheet = wb.sheets[sheet_name]

        han_ji_khoo = wb.names['漢字庫'].refers_to_range.value
        han_ji_piau_im_huat = wb.names['標音方法'].refers_to_range.value
        piau_im = PiauIm(han_ji_khoo)

    except Exception as e:
        raise ValueError(f"無法找到或建立工作表 '{sheet_name}'：{e}")

    try:
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
            han_ji_piau_im_sheet.activate()
            han_ji_piau_im_sheet.range((row, 1)).select()

            # 逐欄取出漢字處理
            for col in range(start_col, end_col):
                # 取得【漢字注音】表中的【漢字】儲存格內容
                # jin_kang_piau_im_cell = han_ji_piau_im_sheet.range((row - 2, col))
                tai_gi_cell = han_ji_piau_im_sheet.range((row - 1, col))
                han_ji_cell = han_ji_piau_im_sheet.range((row, col))
                han_ji_piau_im_cell = han_ji_piau_im_sheet.range((row + 1, col))

                # 依據【漢字】儲存格讀取之資料，進行處理作業
                if han_ji_cell.value == 'φ':
                    EOF = True
                    msg = f"《文章終止》"
                    break
                elif han_ji_cell.value == '\n':
                    msg = f"《換行》"
                    break
                elif not is_han_ji(han_ji_cell.value):
                    # 若儲存格為：非【漢字】，有可能為：全形/半形【標點符號】，或半形字元
                    msg = f"{han_ji_cell.value}"
                else:
                    # ---------------------------------------------------------
                    # 確認【漢字】有【台語標音】時之處理作業
                    # ---------------------------------------------------------
                    if tai_gi_cell.value:        # 【漢字】沒用【人工標音】
                        siann_bu, un_bu, tiau_ho = split_tai_gi_im_piau(tai_gi_cell.value)
                        han_ji_piau_im = piau_im.han_ji_piau_im_tng_huan(
                            piau_im_huat=han_ji_piau_im_huat,
                            siann_bu=siann_bu,
                            un_bu=un_bu,
                            tiau_ho=tiau_ho,
                        )
                        tlpa_im_piau = f"{siann_bu}{un_bu}{tiau_ho}"
                        han_ji_piau_im_cell.value = han_ji_piau_im

                        msg = f"{han_ji_cell.value} [{tlpa_im_piau}] / [{han_ji_piau_im}]"

                # 每欄結束前處理作業
                print(f"({row}, {col}) = {msg}")

            # 每列結束前處理作業
            print(f"({row}, {col}) = {msg}")
            row += ROWS_PER_LINE
            col = start_col
            line += 1
            if EOF or line > TOTAL_LINES: break
    except Exception as e:
        logging_exception(msg=f"處理【人工標音】作業異常！", error=e)
        raise

    #-------------------------------------------------------------------------------------
    # 作業結束前處理
    #-------------------------------------------------------------------------------------
    han_ji_piau_im_sheet.activate()
    han_ji_piau_im_sheet.range('A1').select()

    print("------------------------------------------------------")
    msg = f'依【漢字】之【台語音標】轉換【漢字音標音】作業己完成！'
    logging_process_step(msg)
    logging_process_step(f'【語音類型】：{ue_im_lui_piat}')
    logging_process_step(f'【漢字庫】：{db_name}')

    return EXIT_CODE_SUCCESS


def process(wb):
    logging_process_step("<----------- 作業開始！---------->")
    # ---------------------------------------------------------------------
    # 重設【漢字】儲存格文字及底色格式
    # ---------------------------------------------------------------------
    # reset_han_ji_cells(wb=wb)

    try:
        han_ji_piau_im(wb=wb, sheet_name='漢字注音')
    except Exception as e:
        logging_exc_error(msg="轉換成【漢字標音】時發生錯誤！", error=e)
        return EXIT_CODE_PROCESS_FAILURE

    #--------------------------------------------------------------------------
    # 結束作業
    #--------------------------------------------------------------------------
    # 要求畫面回到【漢字注音】工作表
    wb.sheets['漢字注音'].activate()
    # 作業正常結束
    logging_process_step("<----------- 作業結束！---------->")
    return EXIT_CODE_SUCCESS

# =============================================================================
# 程式主流程
# =============================================================================
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
        status_code = process(wb)
        if status_code != EXIT_CODE_SUCCESS:
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