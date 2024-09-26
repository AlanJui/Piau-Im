import getopt
import os
import sys

import xlwings as xw

import settings


def get_input_and_output_options(argv):
    arg_input = ""
    arg_output = ""
    arg_help = "{0} -i <input> -o <output>".format(argv[0])

    try:
        opts, args = getopt.getopt(  # pyright: ignore
            argv[1:], "hi:o:", ["help", "input=", "output="]
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
        elif opt in ("-o", "--output"):
            arg_output = arg

    print("input:", arg_input)
    print("output:", arg_output)

    return {
        "input": arg_input,
        "output": arg_output,
    }


if __name__ == "__main__":
    # =========================================================================
    # (1) 取得需要注音的「檔案名稱」及其「目錄路徑」。
    # =========================================================================
    # 取得 Input 檔案名稱
    opts = get_input_and_output_options(sys.argv)
    if opts["input"] != "":
        CONVERT_FILE_NAME = opts["input"].replace(" ", "")
    else:
        CONVERT_FILE_NAME = "Tai_Gi_Zu_Im_Bun.xlsx"
    print(f"CONVERT_FILE_NAME = {CONVERT_FILE_NAME}")

    # 打開 Excel 檔案
    wb = xw.Book(CONVERT_FILE_NAME)

    # =========================================================================
    # (2) 依據《文章標題》另存新檔。
    # =========================================================================
    setting_sheet = wb.sheets["env"]
    new_file_name = str(
        setting_sheet.range("C4").value
    ).strip()
    
    # 設定檔案輸出路徑，存於專案根目錄下的 output2 資料夾
    output_path = wb.names['OUTPUT_PATH'].refers_to_range.value 
    new_file_path = os.path.join(
        ".\\{0}".format(output_path), 
        f"【河洛話注音】{new_file_name}" + ".xlsx")

    # 儲存新建立的工作簿
    wb.save(new_file_path)

    # 保存 Excel 檔案
    wb.close()
