# 填漢字等標音：將整段的文字拆解，個別填入儲存格，以便後續人工手動填入台語音標、注音符號。
import getopt
import sys

import settings
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

    # =========================================================================
    # (2) 將漢字填入
    #     - 上方：台語音標
    #     - 下方：台語注音符號
    # =========================================================================
    fill_hanji_in_cells(CONVERT_FILE_NAME)