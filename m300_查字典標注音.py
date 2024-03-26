import getopt
import sys

import settings
# from p100_tsa_ji_tian import main_run as tsa_ji_tian_tshue_tsu_im
from p100_tsa_nga_siok_thong_ji_tian import main_run as tsa_ji_tian_tshue_tsu_im


def myfunc(argv):
    arg_input = ""
    arg_output = ""
    arg_user = ""
    arg_help = "{0} -i <input> -u <user> -o <output>".format(argv[0])

    try:
        # pyright: ignore[reportOptionalMemberAccess]
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
    # ===========================================================================
    # (1) 取得需要注音的「檔案名稱」及其「目錄路徑」。
    # ===========================================================================
    # 取得 Input 檔案名稱
    file_path = settings.get_input_file_path()
    if not file_path:
        print("未設定 .env 檔案")
        # sys.exit(2)
        opts = myfunc(sys.argv)
        if opts["input"] != "":
            CONVERT_FILE_NAME = opts["input"]
        else:
            CONVERT_FILE_NAME = "Piau-Tsu-Im"
    else:
        CONVERT_FILE_NAME = file_path
    print(f"CONVERT_FILE_NAME = {CONVERT_FILE_NAME}")

    # ===========================================================================
    # (2) 在字典查注音，填入漢字注音表。
    # ===========================================================================
    tsa_ji_tian_tshue_tsu_im(CONVERT_FILE_NAME)
