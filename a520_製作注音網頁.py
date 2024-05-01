#================================================================
# 《予我廣韻標音》
# 使用《廣韻》作為漢字標讀音之依據。
#================================================================
import getopt
import os
import sqlite3
import sys

import xlwings as xw

import settings
from mod_file_access import (
    close_excel_file,
    get_cmd_input,
    open_excel_file,
    save_to_a_working_copy,
    write_to_excel_file,
)
from mod_廣韻 import init_sing_bu_dict, init_un_bu_dict
from p500_Import_Source_Sheet import San_Sing_Han_Ji_Tsh_Im_Piau
from p501_Kong_Un_Cha_Ji_Tian import Kong_Un_Piau_Im
from p502_TLPA_Cu_Im import Iong_TLPA_Cu_Im


def main():
    # =========================================================="
    # 資料庫",
    # =========================================================="
    # 自 .env 檔案取得資料庫名稱
    DATABASE = settings.get_database_path()
    conn = sqlite3.connect(DATABASE)
    db_cursor = conn.cursor()
    print(f"DATABASE = {DATABASE}")

    # =========================================================================
    # (1) 取得需要注音的「檔案名稱」及其「目錄路徑」。
    # =========================================================================
    # 自命令列取得 Input 檔案名稱
    opts = get_cmd_input()
    CONVERT_FILE_NAME = opts["input"]
    print(f"CONVERT_FILE_NAME = {CONVERT_FILE_NAME}")

    # 指定提供來源的【檔案】
    # wb = open_excel_file(CONVERT_FILE_NAME)
    save_to_a_working_copy(CONVERT_FILE_NAME)
    wb = open_excel_file("working.xlsx")

    # =========================================================================
    # (2) 建置「漢字注音表」
    # 將存放在「工作表1」的「漢字」文章，製成「漢字注音表」以便填入注音。
    # =========================================================================
    # San_Sing_Han_Ji_Tsh_Im_Piau(CONVERT_FILE_NAME)

    # =========================================================================
    # (3) 在字典查注音，填入漢字注音表。
    # =========================================================================
    # Kong_Un_Piau_Im(CONVERT_FILE_NAME, db_cursor)

    # =========================================================================
    # (4) 將已注音之「漢字注音表」，製作成 HTML 格式之「注音／拼音／標音」網頁。
    # =========================================================================

    # 設定聲母及韻母之注音對照表
    try:
        sing_bu_dict = init_sing_bu_dict(db_cursor)
        un_bu_dict = init_un_bu_dict(db_cursor)
    except Exception as e:
        print(e)
    Iong_TLPA_Cu_Im(wb, sing_bu_dict, un_bu_dict)

    # ==========================================================
    # 檢查「缺字表」狀態
    # ==========================================================
    # 指定來源工作表
    source_sheet = wb.sheets["缺字表"]
    # 取得工作表內總列數
    end_row_no = (
        source_sheet.range("A" + str(source_sheet.cells.last_cell.row)).end("up").row
    )
    if end_row_no > 1:
        print(f"總計字典查不到注音的漢字共：{end_row_no}個。")

    # =========================================================================
    # (5) 依據《文章標題》另存新檔。
    # =========================================================================
    # write_to_excel_file(wb)
    close_excel_file(wb)
    
    # ==========================================================
    # 關閉資料庫
    # ==========================================================
    conn.close()

if __name__ == "__main__":
    main()