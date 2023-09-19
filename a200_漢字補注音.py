import getopt
import os
import sys
import xlwings as xw

from dotenv import dotenv_values

# from hun_siann_un_tiau import main_run as ping_im_hun_siann_un_tiau
from p110_khuat_ji_poo_tsu_im import main_run as poo_tsu_im
from p210_hoo_goa_chu_im_all import main_run as hoo_goa_chu_im_all


def myfunc(argv):
    arg_input = ""
    arg_output = ""
    arg_user = ""
    arg_help = "{0} -i <input> -u <user> -o <output>".format(argv[0])

    try:
        opts, args = getopt.getopt(
            argv[1:], "hi:u:o:", ["help", "input=", "user=", "output="]
        )
    except getopt.GetoptError:
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
    # 取得 Input 檔案名稱
    config = dotenv_values(".env")
    dir_path = str(config["INPUT_FILE_PATH"])
    file_name = str(config["FILE_NAME"])
    file_path = os.path.join(dir_path, file_name)
    if not file_path:
        print("未設定 config.env 檔案")
        # sys.exit(2)
        opts = myfunc(sys.argv)
        if opts["input"] != "":
            CONVERT_FILE_NAME = opts["input"]
        else:
            CONVERT_FILE_NAME = "Piau-Tsu-Im.xlsx"
    else:
        CONVERT_FILE_NAME = file_path
    print(f"CONVERT_FILE_NAME = {CONVERT_FILE_NAME}")

    # 將「漢字注音表」的「台羅注音」拆解出聲母、韻母及聲調
    poo_tsu_im(CONVERT_FILE_NAME)

    # 將已填入注音之「漢字注音表」，製作成 HTML 格式的各式「注音／拼音／標音」。
    hoo_goa_chu_im_all(CONVERT_FILE_NAME)

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
