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

# =========================================================================
# 定義 Exit Code
# =========================================================================
EXIT_CODE_SUCCESS = 0  # 成功
EXIT_CODE_NO_FILE = 1  # 無法找到檔案
EXIT_CODE_INVALID_INPUT = 2  # 輸入錯誤
EXIT_CODE_PROCESS_FAILURE = 3  # 過程失敗
EXIT_CODE_UNKNOWN_ERROR = 99  # 未知錯誤


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
    # (2) 嘗試獲取當前作用中的 Excel 工作簿
    # =========================================================================
    wb = None
    try:
        wb = xw.apps.active.books.active
    except Exception as e:
        logging.error(f"無法找到作用中的 Excel 工作簿: {e}", exc_info=True)
        print("無法找到作用中的 Excel 工作簿")
        return EXIT_CODE_NO_FILE

    if not wb:
        print("無法作業，原因可能為：(1) 未指定輸入檔案；(2) 未找到作用中的 Excel 工作簿！")
        logging.error("無法作業，未指定輸入檔案或 Excel 無效。")
        return EXIT_CODE_NO_FILE

    try:
        # =========================================================================
        # (3) 讀取儲存格資料
        # =========================================================================
        cell_value = wb.sheets['漢字注音'].range('V3').value

        if cell_value is None:
            print("【待注音漢字】儲存格未填入文字，作業無法繼續。")
            logging.warning("【待注音漢字】儲存格為空")
            return EXIT_CODE_INVALID_INPUT

        value_length = len(cell_value.strip())
        print(f"【待注音漢字】總字數為: {value_length}")
        logging.info(f"【待注音漢字】總字數為: {value_length}")

        # =========================================================================
        # (4) 執行 Excel 作業
        # =========================================================================
        print("正在清除儲存格...")
        clear_han_ji_kap_piau_im(wb)
        logging.info("儲存格清除完畢")

        print("重設格式...")
        reset_han_ji_cells(wb)
        logging.info("儲存格格式重設完畢")

        print("填入待注音漢字...")
        fill_hanji_in_cells(wb)
        logging.info("填入待注音漢字完成")

        # =========================================================================
        # (5) 自動查找讀音
        # =========================================================================
        type = get_sound_type(wb)
        han_ji_khoo = get_han_ji_khoo(wb)

        if han_ji_khoo in ["河洛話", "廣韻"]:
            db_name = DB_HO_LOK_UE if han_ji_khoo == "河洛話" else DB_KONG_UN
            module_name = 'mod_河洛話' if han_ji_khoo == "河洛話" else 'mod_廣韻'
            hue_im = type if han_ji_khoo == "河洛話" else "文讀音"

            logging.info(f"開始標音作業 - {han_ji_khoo}: {type}")
            ca_han_ji_thak_im(wb, sheet_name='漢字注音', cell='V3', hue_im=hue_im,
                              han_ji_khoo=han_ji_khoo, db_name=db_name,
                              module_name=module_name, function_name='han_ji_ca_piau_im')
            logging.info(f"標音作業完成 - {han_ji_khoo}: {type}")
        else:
            print("無法執行漢字標音作業，請確認設定是否正確！")
            logging.warning("無法執行標音作業，檢查【env】設定。")
            return EXIT_CODE_PROCESS_FAILURE

        # =========================================================================
        # (6) 儲存檔案
        # =========================================================================
        print("儲存檔案...")
        save_as_new_file(wb=wb)
        logging.info("檔案已成功儲存")

    except Exception as e:
        print(f"執行過程中發生未知錯誤: {e}")
        logging.error(f"執行過程中發生未知錯誤: {e}", exc_info=True)
        return EXIT_CODE_UNKNOWN_ERROR

    finally:
        if wb:
            # 是否關閉 Excel 視窗可根據需求決定
            # xw.apps.active.quit()  # 確保 Excel 被釋放資源，避免開啟殘留
            logging.error(f"作業正常完成！")

    return EXIT_CODE_SUCCESS


if __name__ == "__main__":
    exit_code = main()
    if exit_code == EXIT_CODE_SUCCESS:
        print("作業成功完成！")
    else:
        print(f"程式結束，代碼: {exit_code}")
    sys.exit(exit_code)
