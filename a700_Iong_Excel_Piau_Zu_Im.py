import getopt
import math
import os
import sys

import xlwings as xw

import settings
from a720_Thiam_Zu_Im import thiam_zu_im
from p730_Tng_Sing_Bang_Iah import tng_sing_bang_iah


def get_input_and_output_options(argv):
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


if __name__ == "__main__":
    # =========================================================================
    # (1) 取得需要注音的「檔案名稱」及其「目錄路徑」。
    # =========================================================================
    # 取得 Input 檔案名稱
    file_path = settings.get_tai_gi_zu_im_bun_path()
    if not file_path:
        print("未設定 .env 檔案")
        # sys.exit(2)
        opts = get_input_and_output_options(sys.argv)
        if opts["input"] != "":
            CONVERT_FILE_NAME = opts["input"]
        else:
            CONVERT_FILE_NAME = "Tai_Gi_Zu_Im_Bun.xlsx"
    else:
        CONVERT_FILE_NAME = file_path
    print(f"CONVERT_FILE_NAME = {CONVERT_FILE_NAME}")

    # =========================================================================
    # (2) 分析已輸入的【台語音標】及【台語注音符號】，將之各別填入漢字之上、下方。
    #     - 上方：台語音標
    #     - 下方：台語注音符號
    # =========================================================================
    thiam_zu_im(CONVERT_FILE_NAME)

    # =========================================================================
    # (4) 將已注音之「漢字注音表」，製作成 HTML 格式之「注音／拼音／標音」網頁。
    # =========================================================================
    tng_sing_bang_iah(CONVERT_FILE_NAME)

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

    # 保存 Excel 檔案
    wb.close()
