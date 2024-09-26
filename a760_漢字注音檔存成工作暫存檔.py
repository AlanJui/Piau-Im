import getopt
import os
import sys
from pathlib import Path

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
        CONVERT_FILE_NAME = ""
    print(f"CONVERT_FILE_NAME = {CONVERT_FILE_NAME}")

    # 獲取當前檔案的路徑
    current_file_path = Path(__file__).resolve()

    # 專案根目錄
    project_root = current_file_path.parent

    print(f"專案根目錄為: {project_root}")

    # 打開 Excel 檔案
    source_path = "output2"
    old_file_path = os.path.join(
        project_root,
        "{0}".format(source_path), 
        f"{CONVERT_FILE_NAME}")
    wb = xw.Book(f"{old_file_path}")

    # =========================================================================
    # (2) 依據《文章標題》另存新檔。
    # =========================================================================
    setting_sheet = wb.sheets["env"]
    new_file_name = str(
        setting_sheet.range("C4").value
    ).strip()
    
    # 將檔案存放路徑設為：【專案根目錄】之下
    output_file = opts["output"].replace(" ", "")
    new_file_path = os.path.join(
        project_root,
        f"{output_file}")

    # 儲存新建立的工作簿
    wb.save(new_file_path)

    # 保存 Excel 檔案
    wb.close()
