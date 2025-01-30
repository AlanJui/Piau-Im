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

import xlwings as xw
from dotenv import load_dotenv

# 載入自訂模組/函式
from mod_excel_access import ensure_sheet_exists

# =========================================================================
# 載入環境變數
# =========================================================================
load_dotenv()

DB_HO_LOK_UE = os.getenv('DB_HO_LOK_UE', 'Ho_Lok_Ue.db')

# =========================================================================
# 設定日誌
# =========================================================================
logging.basicConfig(
    filename='process_log.txt',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# =========================================================================
# 常數定義
# =========================================================================
EXIT_CODE_SUCCESS = 0
EXIT_CODE_FAILURE = 1
EXIT_CODE_INVALID_INPUT = 2
EXIT_CODE_PROCESS_FAILURE = 3
EXIT_CODE_UNKNOWN_ERROR = 99

# =========================================================================
# 功能 1：使用【人工標音】更新【標音字庫】的校正音標
# =========================================================================
def update_pronunciation_in_excel(wb):
    """
    更新【標音字庫】工作表中的【校正音標】（D 欄）
    - 依據 【人工標音】(row-2, col) 更新 (row, col) 的【校正音標】

    :param wb: Excel 活頁簿物件
    :return: EXIT_CODE_SUCCESS or EXIT_CODE_FAILURE
    """
    sheet_name = "標音字庫"
    active_cell = wb.app.selection  # 取得目前作用儲存格
    cell_address = active_cell.address.replace("$", "")

    row, col = excel_address_to_row_col(cell_address)
    han_ji = active_cell.value

    # 計算人工標音儲存格位置
    artificial_row = row - 2
    artificial_pronounce = wb.sheets[sheet_name].cells(artificial_row, col).value

    # 檢查標音字庫是否有此漢字，並更新校正音標
    sheet = wb.sheets[sheet_name]
    data = sheet.range("A2").expand("table").value

    if not isinstance(data[0], list):
        data = [data]

    for idx, row_data in enumerate(data):
        row_han_ji = row_data[0]
        correction_pronounce_cell = sheet.range(f"D{idx+2}")
        coordinates = row_data[4]

        if row_han_ji == han_ji and coordinates:
            if convert_to_excel_address(str((row, col))) in coordinates:
                if correction_pronounce_cell.value == "N/A":
                    correction_pronounce_cell.value = artificial_pronounce
                    print(f"✅ 更新成功: {han_ji} ({row}, {col}) -> {artificial_pronounce}")
                    return EXIT_CODE_SUCCESS

    print(f"❌ 未找到匹配的資料或不符合更新條件: {han_ji} ({row}, {col})")
    return EXIT_CODE_FAILURE


# =========================================================================
# 功能 2：使用【標音字庫】更新【Ho_Lok_Ue.db】資料庫
# =========================================================================
def update_database_from_excel(wb):
    """
    使用【標音字庫】工作表的資料更新 SQLite 資料庫。
    - 【漢字】 -> `漢字庫`.`漢字`
    - 【校正音標】 -> `漢字庫`.`台羅音標`
    - 略過 `N/A` 的校正音標

    :param wb: Excel 活頁簿物件
    :return: EXIT_CODE_SUCCESS or EXIT_CODE_FAILURE
    """
    sheet_name = "標音字庫"
    sheet = wb.sheets[sheet_name]
    data = sheet.range("A2").expand("table").value

    if not isinstance(data[0], list):
        data = [data]

    conn = sqlite3.connect(DB_HO_LOK_UE)
    cursor = conn.cursor()

    try:
        for row_data in data:
            han_ji = row_data[0]
            tai_lo_pinyin = row_data[3]  # D 欄 (校正音標)

            if han_ji and tai_lo_pinyin and tai_lo_pinyin != "N/A":
                cursor.execute("""
                    INSERT INTO 漢字庫 (漢字, 台羅音標, 常用度, 更新時間)
                    VALUES (?, ?, ?, CURRENT_TIMESTAMP)
                    ON CONFLICT(漢字) DO UPDATE SET 台羅音標=excluded.台羅音標, 更新時間=CURRENT_TIMESTAMP
                """, (han_ji, tai_lo_pinyin, 0.8))  # 常用度固定為 0.8

        conn.commit()
        print("✅ 資料庫更新完成！")
        return EXIT_CODE_SUCCESS

    except Exception as e:
        print(f"❌ 資料庫更新失敗: {e}")
        return EXIT_CODE_FAILURE

    finally:
        conn.close()


# =========================================================================
# 解析 Excel 位置轉換函數
# =========================================================================
def excel_address_to_row_col(cell_address):
    match = re.match(r"([A-Z]+)(\d+)", cell_address)
    if not match:
        raise ValueError(f"無效的 Excel 儲存格地址: {cell_address}")

    col_letters, row_number = match.groups()
    col_number = sum((ord(letter) - ord("A") + 1) * (26 ** i) for i, letter in enumerate(reversed(col_letters)))

    return int(row_number), col_number


# =========================================================================
# 主程式執行
# =========================================================================
def main():
    if len(sys.argv) > 1:
        mode = sys.argv[1]  # 取得 Command Line 參數
    else:
        mode = "1"  # 預設執行 功能 1

    wb = xw.apps.active.books.active  # 取得當前 Excel 活頁簿

    if mode == "1":
        return update_pronunciation_in_excel(wb)
    elif mode == "2":
        return update_database_from_excel(wb)
    else:
        print("❌ 錯誤：請輸入有效模式 (1 或 2)")
        return EXIT_CODE_INVALID_INPUT


if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)
