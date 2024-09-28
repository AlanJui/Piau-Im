#==============================================================================
# 說明：
#
# 本程式可將當前打開的 Excel 檔案另存為指定的檔名。
#
#  1. wb = xw.apps.active.books.active：這行代碼使用 xlwings 獲取當前作用中的 Excel 工作簿。
#  如果沒有已打開的工作簿，它會提示並退出程式。
#
#  2. output_file：通過命令行參數獲取輸出的檔名，若沒有指定，則使用預設的 Tai_Gi_Zu_Im_Bun.xlsx。
# project_root：將新檔案儲存到專案根目錄。
#
# 【執行範例】：
# 在CMD 執行以下命令：
# python your_script.py -o Tai_Gi_Zu_Im_Bun.xlsx
#
#  3. 如果沒有指定參數 -o，預設之另存新檔名稱為 Tai_Gi_Zu_Im_Bun.xlsx。
#==============================================================================
import getopt
import os
import sys
from pathlib import Path

import xlwings as xw


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
    # (1) 取得專案根目錄。
    # =========================================================================
    # 獲取當前檔案的路徑
    current_file_path = Path(__file__).resolve()

    # 專案根目錄
    project_root = current_file_path.parent

    print(f"專案根目錄為: {project_root}")

    # =========================================================================
    # (2) 取得輸入及輸出的檔案名稱。
    # =========================================================================
    # 取得新的檔案名稱
    opts = get_input_and_output_options(sys.argv)
    output_file = opts["output"].replace(" ", "")
    if output_file == "":
        output_file = "Tai_Gi_Zu_Im_Bun.xlsx"  # 預設檔名

    # =========================================================================
    # (3) 若無指定輸入檔案，則獲取當前作用中的 Excel 檔案並另存新檔。
    # =========================================================================
    wb = None
    input_file = opts["input"].replace(" ", "")
    if input_file:
        # 打開 Excel 檔案
        source_path = "output2"
        old_file_path = os.path.join(
            project_root,
            "{0}".format(source_path), 
            f"{input_file}")
        wb = xw.Book(f"{old_file_path}")
    else:
        # 使用已打開且處於作用中的 Excel 工作簿
        try:
            # 嘗試獲取當前作用中的 Excel 工作簿
            wb = xw.apps.active.books.active
        except Exception as e:
            print(f"發生錯誤: {e}")
            print("無法找到作用中的 Excel 工作簿")
            sys.exit(2)

    if not wb:
        print("無法執行，可能原因：(1) 未指定輸入檔案；(2) 未找到作用中的 Excel 工作簿")
        sys.exit(2)

    # 將檔案存放路徑設為【專案根目錄】之下
    new_file_path = os.path.join(
        project_root,
        f"{output_file}"
    )

    # 儲存新建立的工作簿
    print(f"正在將檔案另存為: {new_file_path}")
    wb.save(new_file_path)

    # 關閉工作簿
    wb.close()

    print(f"檔案已成功存為 {new_file_path}")
