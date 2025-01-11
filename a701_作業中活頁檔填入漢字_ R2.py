import os
import sys
from pathlib import Path

import xlwings as xw

from mod_file_access import get_han_ji_khoo, get_sound_type, save_as_new_file
from p701_Clear_Cells import clear_han_ji_kap_piau_im
from p702_Ca_Han_Ji_Thak_Im import ca_han_ji_thak_im
from p709_reset_han_ji_cells import reset_han_ji_cells
from p710_thiam_han_ji import fill_hanji_in_cells

EXIT_CODE_SUCCESS = 0  # 成功
EXIT_CODE_NO_FILE = 1  # 無法找到檔案
EXIT_CODE_INVALID_INPUT = 2  # 輸入錯誤
EXIT_CODE_PROCESS_FAILURE = 3  # 過程失敗
EXIT_CODE_UNKNOWN_ERROR = 99  # 未知錯誤

def main():
    # =========================================================================
    # (1) 取得專案根目錄
    # =========================================================================
    current_file_path = Path(__file__).resolve()
    project_root = current_file_path.parent
    print(f"專案根目錄為: {project_root}")

    # =========================================================================
    # (2) 若無指定輸入檔案，則獲取當前作用中的 Excel 檔案並另存新檔
    # =========================================================================
    try:
        wb = xw.apps.active.books.active
    except Exception as e:
        print(f"發生錯誤: {e}")
        print("無法找到作用中的 Excel 工作簿")
        return EXIT_CODE_NO_FILE

    if not wb:
        print("無法作業，原因可能為：(1) 未指定輸入檔案；(2) 未找到作用中的 Excel 工作簿！")
        return EXIT_CODE_NO_FILE

    try:
        cell_value = wb.sheets['漢字注音'].range('V3').value

        # 判斷是否為 None 並執行處理
        if cell_value is None:
            print("【待注音漢字】儲存格未填入文字，作業無法繼續。")
            return EXIT_CODE_INVALID_INPUT
        else:
            value_length = len(cell_value.strip())
            print(f"【待注音漢字】總字數為: {value_length}")

        # 將儲存格已填入之漢字及標音清除
        clear_han_ji_kap_piau_im(wb)

        # 將待注音的【漢字儲存格】重設格式
        reset_han_ji_cells(wb)

        # 將待注音的漢字填入
        fill_hanji_in_cells(wb)

        # A731: 自動為漢字查找讀音
        type = get_sound_type(wb)
        han_ji_khoo = get_han_ji_khoo(wb)
        if han_ji_khoo == "河洛話" and type in ["白話音", "文讀音"]:
            ca_han_ji_thak_im(wb, sheet_name='漢字注音', cell='V3', ue_im_lui_piat=type, han_ji_khoo="河洛話",
                              db_name='Ho_Lok_Ue.db', module_name='mod_河洛話', function_name='han_ji_ca_piau_im')
        elif han_ji_khoo == "廣韻":
            ca_han_ji_thak_im(wb, sheet_name='漢字注音', cell='V3', ue_im_lui_piat="文讀音", han_ji_khoo="廣韻",
                              db_name='Kong_Un.db', module_name='mod_廣韻', function_name='han_ji_ca_piau_im')
        else:
            print("無法執行漢字標音作業，請確認【env】工作表【語音類型】及【漢字庫】欄位的設定是否正確！")
            return EXIT_CODE_PROCESS_FAILURE

        # 將檔案存放路徑設為【專案根目錄】之下
        save_as_new_file(wb=wb)

    except Exception as e:
        print(f"執行過程中發生未知錯誤: {e}")
        return EXIT_CODE_UNKNOWN_ERROR

    return  EXIT_CODE_SUCCESS

if __name__ == "__main__":
    exit_code = main()
    if exit_code == 0:
        print("作業成功完成！")
    else:
        print(f"程式結束，代碼: {exit_code}")
    sys.exit(exit_code)
