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
from p701_Clear_Cells import clear_han_ji_kap_piau_im
from p702_Ca_Han_Ji_Thak_Im import ca_han_ji_thak_im
from p709_reset_han_ji_cells import reset_han_ji_cells
from p710_thiam_han_ji import fill_hanji_in_cells

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
# 定義 Exit Code
# =========================================================================
EXIT_CODE_SUCCESS = 0  # 成功
EXIT_CODE_NO_FILE = 1  # 無法找到檔案
EXIT_CODE_INVALID_INPUT = 2  # 輸入錯誤
EXIT_CODE_PROCESS_FAILURE = 3  # 過程失敗
EXIT_CODE_UNKNOWN_ERROR = 99  # 未知錯誤


# =========================================================================
# 作業程序
# =========================================================================
def process(wb):
    # ---------------------------------------------------------------------
    # 取得【待注音漢字】總字數
    # ---------------------------------------------------------------------
    cell_value = wb.sheets['漢字注音'].range('V3').value

    if cell_value is None:
        print("【待注音漢字】儲存格未填入文字，作業無法繼續。")
        logging.warning("【待注音漢字】儲存格為空")
        return EXIT_CODE_INVALID_INPUT

    value_length = len(cell_value.strip())
    print(f"【待注音漢字】總字數為: {value_length}")
    logging.info(f"【待注音漢字】總字數為: {value_length}")

    # ---------------------------------------------------------------------
    # 執行儲存格重設作業
    # ---------------------------------------------------------------------
    print("清除儲存格內容...")
    clear_han_ji_kap_piau_im(wb)
    logging.info("儲存格內容清除完畢")

    print("重設儲存格之格式...")
    reset_han_ji_cells(wb)
    logging.info("儲存格格式重設完畢")

    print("待注音漢字填入【漢字注音】工作表...")
    fill_hanji_in_cells(wb)
    logging.info("待注音漢字已填入【漢字注音】工作表")

    # ---------------------------------------------------------------------
    # 為漢字查找標音
    # ---------------------------------------------------------------------
    type = get_sound_type(wb)
    han_ji_khoo = get_han_ji_khoo(wb)

    if han_ji_khoo in ["河洛話", "廣韻"]:
        db_name = DB_HO_LOK_UE if han_ji_khoo == "河洛話" else DB_KONG_UN
        module_name = 'mod_河洛話' if han_ji_khoo == "河洛話" else 'mod_廣韻'
        ue_im_lui_piat = type if han_ji_khoo == "白話音" else "文讀音"

        # 查找漢字標音
        logging.info(f"開始【漢字標音作業】 - {han_ji_khoo}: {type}")
        ca_han_ji_thak_im(wb, sheet_name='漢字注音', cell='V3',
                          ue_im_lui_piat=ue_im_lui_piat,
                          han_ji_khoo=han_ji_khoo, db_name=db_name,
                          module_name=module_name,
                          function_name='han_ji_ca_piau_im')
        logging.info(f"完成【漢字標音作業】 - {han_ji_khoo}: {type}")
    else:
        print("無法執行【漢字標音作業】，請確認設定是否正確！")
        logging.warning("無法執行【漢字標音作業】，需檢查【env】工作表之設定。")
        return EXIT_CODE_PROCESS_FAILURE

    # ---------------------------------------------------------------------
    # 作業結尾處理
    # ---------------------------------------------------------------------
    file_path = save_as_new_file(wb=wb)
    logging.info("己存檔至路徑：{file_path}")
    return EXIT_CODE_SUCCESS


# =============================================================================
# 程式主流程
# =============================================================================
def main():
    logging.info("作業開始")

    # =========================================================================
    # (1) 取得專案根目錄
    # =========================================================================
    current_file_path = Path(__file__).resolve()
    project_root = current_file_path.parent
    print(f"專案根目錄為: {project_root}")
    logging.info(f"專案根目錄為: {project_root}")

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
            logging.info("a701_作業中活頁簿填入漢字.py 程式已執行完畢！")

    return EXIT_CODE_SUCCESS


if __name__ == "__main__":
    exit_code = main()
    if exit_code == EXIT_CODE_SUCCESS:
        print("程式正常完成！")
    else:
        print(f"程式異常終止，錯誤代碼為: {exit_code}")
    sys.exit(exit_code)
