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
from mod_廣韻 import init_sing_bu_dict, init_un_bu_dict
from p500_Import_Source_Sheet import San_Sing_Han_Ji_Tsh_Im_Piau
from p501_Kong_Un_Cha_Ji_Tian import Kong_Un_Piau_Im
from p502_TLPA_Cu_Im import Iong_TLPA_Cu_Im

# 專案全域常數
# from config_dev_env import DATABASE
DATABASE = "Kong_Un_V2.db"


def get_cmd_input(gargv):
    arg_input = ""
    arg_output = ""
    arg_user = ""
    arg_help = "{0} -i <input> -u <user> -o <output>".format(argv[0])

    try:
        opts, args = getopt.getopt(  # pyright: ignore
            argv[1:], "hi:u:o:", ["help", "input=", "user=", "output="]
        )
    except Exception as e:
        print(e)
        print(arg_help)
        sys.exit(2)

    for opt, arg in opts:
        if opt in ("-h", "--help"):
            print(arg_help)  # print the help message
            sys.exit(2)
        elif opt in ("-i", "--input"):
            arg_input = arg
        elif opt in ("-u", "--user"):
            arg_user = arg
        elif opt in ("-o", "--output"):
            arg_output = arg

    print("input:", arg_input)
    print("user:", arg_user)
    print("output:", arg_output)

    return {
        "input": arg_input,
        "user": arg_user,
        "output": arg_output,
    }

def main():
    # =========================================================="
    # 資料庫",
    # =========================================================="
    conn = sqlite3.connect(DATABASE)
    db_cursor = conn.cursor()

    # =========================================================================
    # (1) 取得需要注音的「檔案名稱」及其「目錄路徑」。
    # =========================================================================
    # 取得 Input 檔案名稱
    file_path = settings.get_input_file_path()
    if not file_path:
        print("未設定 .env 檔案")
        # sys.exit(2)
        opts = get_cmd_input(sys.argv)
        if opts["input"] != "":
            CONVERT_FILE_NAME = opts["input"]
        else:
            CONVERT_FILE_NAME = "Piau-Tsu-Im"
    else:
        CONVERT_FILE_NAME = file_path
    print(f"CONVERT_FILE_NAME = {CONVERT_FILE_NAME}")

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
    Iong_TLPA_Cu_Im(CONVERT_FILE_NAME, sing_bu_dict, un_bu_dict)

    # ==========================================================
    # 檢查「缺字表」狀態
    # ==========================================================
    # 指定來源工作表
    source_sheet = xw.Book(CONVERT_FILE_NAME).sheets["缺字表"]
    # 取得工作表內總列數
    end_row_no = (
        source_sheet.range("A" + str(source_sheet.cells.last_cell.row)).end("up").row
    )
    if end_row_no > 1:
        print(f"總計字典查不到注音的漢字共：{end_row_no}個。")

    # =========================================================================
    # (5) 依據《文章標題》另存新檔。
    # =========================================================================
    wb = xw.Book(CONVERT_FILE_NAME)
    setting_sheet = wb.sheets["env"]
    new_file_name = str(
        setting_sheet.range("C4").value
    ).strip()
    new_file_path = os.path.join(
        ".\\output", 
        f"【河洛話注音】{new_file_name}" + ".xlsx")

    # 儲存新建立的工作簿
    wb.save(new_file_path)
    
    # ==========================================================
    # 關閉資料庫
    # ==========================================================
    conn.close()

if __name__ == "__main__":
    main()