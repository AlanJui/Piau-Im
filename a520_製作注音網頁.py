import os
import sqlite3
import sys

import xlwings as xw

import settings
from mod_file_access import (
    San_Sing_Han_Ji_Zu_Im_Piau,
    close_excel_file,
    get_cmd_input,
    open_excel_file,
    save_to_a_working_copy,
    write_to_excel_file,
)
from mod_廣韻 import init_sing_bu_dict, init_un_bu_dict
from p501_Kong_Un_Cha_Ji_Tian import Kong_Un_Piau_Im
from p502_TLPA_Cu_Im import Iong_TLPA_Cu_Im


def initialize_dicts(db_cursor):
    """初始化聲母和韻母對照表字典"""
    try:
        sing_bu_dict = init_sing_bu_dict(db_cursor)
        un_bu_dict = init_un_bu_dict(db_cursor)
        print("字典初始化完成。")
    except Exception as e:
        print(f"字典初始化失敗：{e}")
        sys.exit(1)
    return sing_bu_dict, un_bu_dict


def create_annotation_file(wb, db_cursor):
    """建立注音表並進行查詢填寫"""
    # 建立漢字注音表
    San_Sing_Han_Ji_Zu_Im_Piau(wb.name)
    # # 查詢注音並填寫表格
    # Kong_Un_Piau_Im(wb.name, db_cursor)


def export_to_html(wb):
    """將已注音的漢字注音表導出為 HTML 格式"""
    # 這裡可以使用已經填寫的漢字注音表進行轉換
    print("將注音表轉換為 HTML 格式的功能可以在這裡實現。")


def main():
    #-------------------------------------------------------------------------
    # 使用已打開且處於作用中的 Excel 工作簿
    #-------------------------------------------------------------------------
    # 取得專案根目錄。
    try:
        wb = xw.apps.active.books.active
    except Exception as e:
        print(f"發生錯誤: {e}")
        print("無法找到作用中的 Excel 工作簿")
        sys.exit(2)

    # 獲取活頁簿的完整檔案路徑
    file_path = wb.fullname
    print(f"完整檔案路徑: {file_path}")

    # 獲取活頁簿的檔案名稱（不包括路徑）
    file_name = wb.name
    print(f"檔案名稱: {file_name}")

    # 資料庫連接
    DATABASE = settings.get_database_path()
    with sqlite3.connect(DATABASE) as conn:
        db_cursor = conn.cursor()
        print(f"DATABASE = {DATABASE}")

        # # 取得命令列參數和檔案路徑
        # opts = get_cmd_input()
        # CONVERT_FILE_NAME = opts["input"]
        # dir_path = opts["dir_path"]
        # print(f"處理檔案: {CONVERT_FILE_NAME}")

        # # 開啟指定的 Excel 檔案
        # wb = open_excel_file(dir_path, CONVERT_FILE_NAME)
        # if wb is None:
        #     print("無法開啟檔案，終止程式執行。")
        #     sys.exit()

        # 初始化字典
        sing_bu_dict, un_bu_dict = initialize_dicts(db_cursor)

        # 創建漢字注音表並查詢注音
        create_annotation_file(wb, db_cursor)

        # 注音轉換處理
        Iong_TLPA_Zu_Im(wb, sing_bu_dict, un_bu_dict)

        # 檢查缺字表狀態
        source_sheet = wb.sheets["缺字表"]
        end_row_no = source_sheet.range("A" + str(source_sheet.cells.last_cell.row)).end("up").row
        if end_row_no > 1:
            print(f"總計字典查不到注音的漢字共：{end_row_no}個。")

        # 儲存 Excel 檔案並關閉
        write_to_excel_file(wb)
        close_excel_file(wb)

if __name__ == "__main__":
    main()
