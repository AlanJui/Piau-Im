# 填漢字等標音：將整段的文字拆解，個別填入儲存格，以便後續人工手動填入台語音標、注音符號。
import getopt
import math
import sys

import xlwings as xw

import settings


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


def clear_hanji_in_cells(wb, sheet_name='漢字注音', cell='V3'):
    # 選擇指定的工作表
    sheet = wb.sheets[sheet_name]

    # 取得 V3 儲存格的字串
    v3_value = sheet.range(cell).value

    # 確認 V3 不為空
    if v3_value:
        # 計算字串的總長度
        total_length = len(v3_value)

        # 每列最多處理 15 個字元，計算總共需要多少列
        chars_per_row = 15
        total_rows_needed = math.ceil(total_length / chars_per_row)  # 無條件進位

        # 迴圈清空所有漢字的上下方儲存格 (羅馬拼音和台語注音符號)
        row = 5
        for i in range(total_rows_needed+1):
            for col in range(4, 19):  # 【D欄=4】到【R欄=18】
                # 清空漢字儲存格 (Row)
                sheet.range((row, col)).value = None
                # 清空上方的台語拼音儲存格 (Row-1)
                sheet.range((row - 1, col)).value = None
                # 清空下方的台語注音儲存格 (Row+1)
                sheet.range((row + 1, col)).value = None
                # 清空填入注音的儲存格 (Row-2)
                sheet.range((row - 2, col)).value = None

            # 每處理 15 個字元後，換到下一行
            row += 4



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
    # (2) 清除原先已填入的漢字
    # =========================================================================
    clear_hanji_in_cells(wb, '漢字注音', 'V3')
