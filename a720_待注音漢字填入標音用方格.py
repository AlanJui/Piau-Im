# 填漢字等標音：將整段的文字拆解，個別填入儲存格，以便後續人工手動填入台語音標、注音符號。
import getopt
import sys

import xlwings as xw

import settings
from p701_Clear_Cells import clear_hanji_in_cells
from p710_thiam_han_ji import fill_hanji_in_cells


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

    # 打開 Excel 檔案
    wb = xw.Book(CONVERT_FILE_NAME)

    # 顯示「已輸入之拼音字母及注音符號」 
    named_range = wb.names['顯示注音輸入']  # 選擇名為 "顯示注音輸入" 的命名範圍# 選擇名為 "顯示注音輸入" 的命名範圍
    named_range.refers_to_range.value = True

    # =========================================================================
    # (2) 將漢字填入
    #     - 上方：台語音標
    #     - 下方：台語注音符號
    # =========================================================================
    fill_hanji_in_cells(wb)     # 將漢字逐個填入各儲存格

    # =========================================================================
    # (3) 依據《文章標題》另存新檔。
    # =========================================================================
    # wb = xw.Book(CONVERT_FILE_NAME)
    # setting_sheet = wb.sheets["env"]
    # new_file_name = str(
    #     setting_sheet.range("C4").value
    # ).strip()
    # new_file_path = os.path.join(
    #     ".\\output", 
    #     f"【河洛話注音】{new_file_name}" + ".xlsx")

    # # 儲存新建立的工作簿
    # wb.save(new_file_path)

    # # 保存 Excel 檔案
    # wb.close()