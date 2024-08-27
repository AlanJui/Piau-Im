import getopt
import math
import os
import sys

import xlwings as xw

import settings
from p700_cu_zu_im import cu_zu_im


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
    file_path = settings.get_input_file_path()
    if not file_path:
        print("未設定 .env 檔案")
        # sys.exit(2)
        opts = get_input_and_output_options(sys.argv)
        if opts["input"] != "":
            CONVERT_FILE_NAME = opts["input"]
        else:
            CONVERT_FILE_NAME = "Piau-Tsu-Im"
    else:
        CONVERT_FILE_NAME = file_path
    print(f"CONVERT_FILE_NAME = {CONVERT_FILE_NAME}")

    # =========================================================================
    # (2) 分析已輸入的【台語音標】及【台語注音符號】，將之各別填入漢字之上、下方。
    #     - 上方：台語音標
    #     - 下方：台語注音符號
    # =========================================================================

    # 打開一個新的或現有的 Excel 檔案
    wb = xw.Book('Tai_Gi_Zu_Im_Bun.xlsx')  # 替換成你的 Excel 檔案名稱

    # 選擇工作表
    # sheet = wb.sheets[0]  # 選擇第一個工作表
    sheet = wb.sheets['漢字注音']

    # 取得 V3 儲存格的字串
    v3_value = sheet.range('V3').value
    
    # 確認 V3 不為空
    if v3_value:
        # 計算字串的總長度
        total_length = len(v3_value)

        # 每列最多處理 15 個字元，計算總共需要多少列
        chars_per_row = 15
        total_rows_needed = math.ceil(total_length / chars_per_row)  # 無條件進位

        # 逐行處理資料，從 Row 4 開始，每列處理 15 個字元
        row = 3
        for i in range(total_rows_needed):
            for col in range(4, 19):  # D列到R列, D=4, R=18
                cell_value = sheet.range((row, col)).value  # 取得 D4, E4, ..., R4 的內容

                # 確認內容不為空
                if cell_value:
                    # 分割字串來提取羅馬拼音和台語注音
                    romaji = cell_value.split('〔')[1].split('〕')[0]  # 取得〔羅馬拼音〕
                    zhuyin = cell_value.split('【')[1].split('】')[0]  # 取得【台語注音】

                    # 將羅馬拼音填入當前 row + 1 的儲存格
                    sheet.range((row + 1, col)).value = romaji

                    # 將台語注音填入當前 row + 3 的儲存格
                    sheet.range((row + 3, col)).value = zhuyin

            # 每處理 15 個字元後，換到下一行
            row += 4
            
    print("已完成【台語音標】和【台語注音符號】標註工作。")

    # =========================================================================
    # (4) 將已注音之「漢字注音表」，製作成 HTML 格式之「注音／拼音／標音」網頁。
    # =========================================================================
    # hoo_gua_tsu_im(CONVERT_FILE_NAME)

    # =========================================================================
    # (5) 依據《文章標題》另存新檔。
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

    # 保存 Excel 檔案
    wb.save('Tai_Gi_Zu_Im_Bun.xlsx')
    wb.close()
