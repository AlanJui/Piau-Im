import getopt
import sys

import xlwings as xw

import settings
from p702_Ca_Han_Ji_Thak_Im import ca_han_ji_thak_im


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

    # =========================================================================
    # (2) 分析已輸入的【台語音標】及【台語注音符號】，將之各別填入漢字之上、下方。
    #     - 上方：台語音標
    #     - 下方：台語注音符號
    # =========================================================================
    ca_han_ji_thak_im(wb, '漢字注音', 'V3')
