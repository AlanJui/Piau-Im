import argparse
import os

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


def open_excel_file(main_file_name):
    excel_file_name = main_file_name + ".xlsx"
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

# -----------------------------------------------------
# 將「字串」轉換成「串列（Characters List）」
# Python code to convert string to list character-wise
def convert_string_to_chars_list(string):
    list1 = []
    list1[:0] = string
    return list1