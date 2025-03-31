# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import logging
import os
import re
import sqlite3
import sys
from datetime import datetime
from pathlib import Path

# 載入第三方套件
import xlwings as xw
from dotenv import load_dotenv

# 載入自訂模組/函式
from mod_excel_access import (
    check_and_update_pronunciation,
    create_dict_by_sheet,
    ensure_sheet_exists,
    get_active_cell_info,
    get_ji_khoo,
    get_value_by_name,
    maintain_ji_khoo,
)

# =========================================================================
# 載入環境變數
# =========================================================================
load_dotenv()

# 預設檔案名稱從環境變數讀取
DB_HO_LOK_UE = os.getenv('DB_HO_LOK_UE', 'Ho_Lok_Ue.db')
DB_KONG_UN = os.getenv('DB_KONG_UN', 'Kong_Un.db')

# =========================================================================
# 設定日誌
# =========================================================================
logging.basicConfig(
    filename='process_log.txt',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def logging_process_step(msg):
    print(msg)
    logging.info(msg)

# =========================================================================
# 常數定義
# =========================================================================
# 定義 Exit Code
EXIT_CODE_SUCCESS = 0  # 成功
EXIT_CODE_FAILURE = 1  # 失敗
EXIT_CODE_NO_FILE = 1  # 無法找到檔案
EXIT_CODE_INVALID_INPUT = 2  # 輸入錯誤
EXIT_CODE_PROCESS_FAILURE = 3  # 過程失敗
EXIT_CODE_UNKNOWN_ERROR = 99  # 未知錯誤

# =========================================================================
# 作業程序
# =========================================================================
def check_han_ji_in_excel(wb, han_ji, excel_cell):
    """
    在【標音字庫】工作表內查詢【漢字】與【Excel座標】是否存在。

    :param wb: Excel 活頁簿物件
    :param han_ji: 要查找的漢字 (str)
    :param excel_cell: 要查找的 Excel 座標 (如 "D9")
    :return: Boolean 值 (True: 找到, False: 未找到)
    """
    sheet_name = "標音字庫"  # Excel 工作表名稱
    try:
        sheet = wb.sheets[sheet_name]
    except Exception:
        print(f"⚠️ 無法找到工作表: {sheet_name}")
        return False

    # 讀取資料範圍
    data = sheet.range("A2").expand("table").value  # 讀取所有資料

    # 確保資料為 2D 列表
    if not isinstance(data[0], list):
        data = [data]

    for row in data:
        row_han_ji = row[0]  # A 欄: 漢字
        coordinates = row[4]  # E 欄: 座標 (可能是 "(9, 4); (25, 9)" 這類格式)

        if row_han_ji == han_ji and coordinates:
            # 將座標解析成一個 set
            coord_list = coordinates.split("; ")
            parsed_coords = {convert_to_excel_address(coord) for coord in coord_list}

            # 檢查 Excel 座標是否在列表內
            if excel_cell in parsed_coords:
                return True

    return False


def convert_to_excel_address(coord_str):
    """
    轉換 `(row, col)` 格式為 Excel 座標 (如 `(9, 4)` 轉換為 "D9")

    :param coord_str: 例如 "(9, 4)"
    :return: Excel 座標字串，例如 "D9"
    """
    coord_str = coord_str.strip("()")  # 去除括號
    try:
        row, col = map(int, coord_str.split(", "))
        return f"{chr(64 + col)}{row}"  # 轉換成 Excel 座標
    except ValueError:
        return ""  # 避免解析錯誤


def excel_address_to_row_col(cell_address):
    """
    將 Excel 儲存格地址 (如 'D9') 轉換為 (row, col) 格式。

    :param cell_address: Excel 儲存格地址 (如 'D9', 'AA15')
    :return: (row, col) 元組，例如 (9, 4)
    """
    match = re.match(r"([A-Z]+)(\d+)", cell_address)  # 用 regex 拆分字母(列) 和 數字(行)

    if not match:
        raise ValueError(f"無效的 Excel 儲存格地址: {cell_address}")

    col_letters, row_number = match.groups()

    # 將 Excel 字母列轉換成數字，例如 A -> 1, B -> 2, ..., Z -> 26, AA -> 27
    col_number = 0
    for letter in col_letters:
        col_number = col_number * 26 + (ord(letter) - ord("A") + 1)

    return int(row_number), col_number


def get_active_cell(wb):
    """
    獲取目前作用中的 Excel 儲存格 (Active Cell)

    :param wb: Excel 活頁簿物件 (xlwings.Book)
    :return: (工作表名稱, 儲存格地址)，如 ("漢字注音", "D9")
    """
    active_cell = wb.app.selection  # 獲取目前作用中的儲存格
    sheet_name = active_cell.sheet.name  # 獲取所在的工作表名稱
    cell_address = active_cell.address.replace("$", "")  # 取得 Excel 格式地址 (去掉 "$")

    return sheet_name, cell_address


def ut01():
    han_ji = "傀"  # 要查找的漢字
    excel_cell = "D9"  # 要查找的 Excel 座標

    exists = check_han_ji_in_excel(wb, han_ji, excel_cell)
    if exists:
        print(f"✅ 漢字 '{han_ji}' 存在於 {excel_cell}")
    else:
        print(f"❌ 找不到漢字 '{han_ji}' 在 {excel_cell}")

    return EXIT_CODE_SUCCESS


def ut02():
    # 作業流程：獲取當前作用中的 Excel 儲存格
    sheet_name, cell_address = get_active_cell(wb)
    print(f"✅ 目前作用中的儲存格：{sheet_name} 工作表 -> {cell_address}")

    # 將 Excel 儲存格地址轉換為 (row, col) 格式
    row, col = excel_address_to_row_col(cell_address)
    print(f"📌 Excel 位址 {cell_address} 轉換為 (row, col): ({row}, {col})")

    return EXIT_CODE_SUCCESS


# =============================================================================
# 作業主流程
# =============================================================================
def process(wb):
    """
    作業流程：
    1. 取得當前 Excel 作用儲存格 (漢字、座標)
    2. 計算【人工標音】位置與值
    3. 查詢【標音字庫】確認該座標是否已登錄
    4. 若【標正音標】為 'N/A'，則更新為【人工標音】
    """
    # 取得當前 Excel 作用儲存格資訊
    sheet_name, han_ji, active_cell, artificial_pronounce, position = get_active_cell_info(wb)

    print(f"📌 作用儲存格：{active_cell}，位於【{sheet_name}】工作表")
    print(f"📌 漢字：{han_ji}，漢字儲存格座標：{active_cell}")
    print(f"📌 人工標音：{artificial_pronounce}，人工標音儲存格座標：{position}")

    # 執行檢查與更新
    success = check_and_update_pronunciation(wb, han_ji, active_cell, artificial_pronounce)

    return EXIT_CODE_SUCCESS if success else EXIT_CODE_FAILURE


# =============================================================================
# 程式主流程
# =============================================================================
def main():
    # =========================================================================
    # 開始作業
    # =========================================================================
    logging.info("作業開始")

    # =========================================================================
    # (1) 取得專案根目錄。
    # =========================================================================
    current_file_path = Path(__file__).resolve()
    project_root = current_file_path.parent
    logging_process_step(f"專案根目錄為: {project_root}")

    # =========================================================================
    # (2) 設定【作用中活頁簿】：偵測及獲取 Excel 已開啟之活頁簿檔案。
    # =========================================================================
    wb = None
    # 取得【作用中活頁簿】
    try:
        wb = xw.apps.active.books.active    # 取得 Excel 作用中的活頁簿檔案
    except Exception as e:
        print(f"發生錯誤: {e}")
        logging.error(f"無法找到作用中的 Excel 工作簿: {e}", exc_info=True)
        return EXIT_CODE_NO_FILE

    # 若無法取得【作用中活頁簿】，則因無法繼續作業，故返回【作業異常終止代碼】結束。
    if not wb:
        return EXIT_CODE_NO_FILE

    # =========================================================================
    # (3) 執行【處理作業】
    # =========================================================================
    try:
        result_code = process(wb)
        if result_code != EXIT_CODE_SUCCESS:
            logging_process_step("作業異常終止！")
            return result_code

    except Exception as e:
        print(f"作業過程發生未知的異常錯誤: {e}")
        logging.error(f"作業過程發生未知的異常錯誤: {e}", exc_info=True)
        return EXIT_CODE_UNKNOWN_ERROR

    finally:
        if wb:
            # xw.apps.active.quit()  # 確保 Excel 被釋放資源，避免開啟殘留
            logging.info("a702_查找及填入漢字標音.py 程式已執行完畢！")

    # =========================================================================
    # 結束作業
    # =========================================================================
    logging.info("作業完成！")
    return EXIT_CODE_SUCCESS


if __name__ == "__main__":
    exit_code = main()
    if exit_code == EXIT_CODE_SUCCESS:
        print("程式正常完成！")
    else:
        print(f"程式異常終止，錯誤代碼為: {exit_code}")
    sys.exit(exit_code)
