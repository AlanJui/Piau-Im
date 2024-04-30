import argparse
import os
import os.path
import time

import xlwings as xw


def get_cmd_input():
    parser = argparse.ArgumentParser(description='Process some files.')
    parser.add_argument('-i', '--input', default='Piau_Tsu_Im', help='Input file name')
    parser.add_argument('-o', '--output', default='', help='Output file name')
    args = parser.parse_args()

    return {
        "input": args.input,
        "output": args.output,
    }


# def open_excel_file(main_file_name):
#     # excel_file_name = main_file_name + ".xlsx"
#     excel_file_name = main_file_name
#     current_path = os.getcwd()
#     file_path = os.path.join(current_path, "output", excel_file_name)
#     return xw.Book(file_path)
def open_excel_file(main_file_name):
    # 檢查檔案名稱是否已包含副檔名
    file_name, file_extension = os.path.splitext(main_file_name)
    if not file_extension:
        # 如果沒有副檔名，添加 .xlsx
        excel_file_name = file_name + '.xlsx'
    else:
        excel_file_name = main_file_name

    current_path = os.getcwd()
    file_path = os.path.join(current_path, "output", excel_file_name)
    return xw.Book(file_path)


def write_to_excel_file(excel_workbook):
    # 自工作表「env」取得新檔案名稱
    setting_sheet = excel_workbook.sheets["env"]
    new_file_name = str(
        setting_sheet.range("C4").value
    ).strip()
    current_path = os.getcwd()
    new_file_path = os.path.join(
        current_path,
        "output", 
        f"【河洛話注音】{new_file_name}" + ".xlsx")
    print(f"儲存輸出，置於檔案：{new_file_path}")

    # 儲存新建立的工作簿
    excel_workbook.save(new_file_path)

    # 等待一段時間讓 save 完成
    time.sleep(10)

    # 關閉工作簿
    excel_workbook.close()

# -----------------------------------------------------
# 將「字串」轉換成「串列（Characters List）」
# Python code to convert string to list character-wise
def convert_string_to_chars_list(string):
    list1 = []
    list1[:0] = string
    return list1

# -----------------------------------------------------
# 要生成超連結的目錄
# directory = 'output'
# extenstion = 'xlsx'
# exculude_list = ['Piau-Tsu-Im.xlsx', 'env.xlsx', 'env_osX.xlsx']

def create_file_list(directory, extension, exculude_list):
    # 建立檔案清單
    file_list = []

    # 遍歷目錄下的檔案
    for filename in os.listdir(directory):
        # 排除 index.html 和 _template.html 檔案
        if filename not in exculude_list:
            if filename.endswith(extension):
                file_list.append(filename)

    return file_list