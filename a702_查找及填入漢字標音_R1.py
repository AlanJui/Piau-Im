# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import logging
import os
import sys
from pathlib import Path

# 載入第三方套件
import xlwings as xw
from dotenv import load_dotenv

# 載入自訂模組
from mod_file_access import get_han_ji_khoo, get_sound_type, save_as_new_file
from p702_Ca_Han_Ji_Thak_Im import ca_han_ji_thak_im

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
EXIT_CODE_NO_FILE = 1  # 無法找到檔案
EXIT_CODE_INVALID_INPUT = 2  # 輸入錯誤
EXIT_CODE_PROCESS_FAILURE = 3  # 過程失敗
EXIT_CODE_UNKNOWN_ERROR = 99  # 未知錯誤

# =========================================================================
# 作業程序
# =========================================================================
def process(wb):
    # ------------------------------------------------------------------------------
    # 指定【作業中工作表】
    # ------------------------------------------------------------------------------
    sheet = wb.sheets["漢字注音"]  # 選擇工作表
    sheet.activate()  # 將「漢字注音」工作表設為作用中工作表
    sheet.range("A1").select()  # 將 A1 儲存格設為作用儲存格

    # ------------------------------------------------------------------------------
    # 為漢字查找讀音，漢字上方填：【台語音標】；漢字下方填使用者指定之【漢字標音】
    # ------------------------------------------------------------------------------
    type = get_sound_type(wb)  # 取得【語音類型】，判別使用【白話音】或【文讀音】何者。
    han_ji_khoo = get_han_ji_khoo(wb)
    if han_ji_khoo == "河洛話" and type == "白話音":
        ca_han_ji_thak_im(
            wb=wb,
            sheet_name="漢字注音",
            cell="V3",
            ue_im_lui_piat=type,
            han_ji_khoo="河洛話",
            db_name="Ho_Lok_Ue.db",
            module_name="mod_河洛話",
            function_name="han_ji_ca_piau_im",
        )
    elif han_ji_khoo == "河洛話" and type == "文讀音":
        ca_han_ji_thak_im(
            wb=wb,
            sheet_name="漢字注音",
            cell="V3",
            ue_im_lui_piat=type,
            han_ji_khoo="河洛話",
            db_name="Ho_Lok_Ue.db",
            module_name="mod_河洛話",
            function_name="han_ji_ca_piau_im",
        )
    elif han_ji_khoo == "廣韻":
        ca_han_ji_thak_im(
            wb=wb,
            sheet_name="漢字注音",
            cell="V3",
            ue_im_lui_piat="文讀音",
            han_ji_khoo="廣韻",
            db_name="Kong_Un.db",
            module_name="mod_廣韻",
            function_name="han_ji_ca_piau_im",
        )
    else:
        msg = "無法執行漢字標音作業，請確認【env】工作表【語音類型】及【漢字庫】欄位的設定是否正確！"
        print(msg)
        logging.error(msg)
        return EXIT_CODE_INVALID_INPUT

    #--------------------------------------------------------------------------
    # 作業結尾處理
    #--------------------------------------------------------------------------
    save_as_new_file(wb=wb)
    logging.info("己存檔至路徑：{file_path}")
    return EXIT_CODE_SUCCESS


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
            logging.info("a702_查找及填入漢字標音.py 程式已執行完畢！")

    # =========================================================================
    # 結束作業
    # =========================================================================
    logging.info("作業完成！")
    return EXIT_CODE_SUCCESS


if __name__ == "__main__":
    exit_code = main()
    if exit_code == EXIT_CODE_SUCCESS:
        print("程式正常完成！")
    else:
        print(f"程式異常終止，錯誤代碼為: {exit_code}")
    sys.exit(exit_code)
