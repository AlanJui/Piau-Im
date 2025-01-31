# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import logging
import os
import re
import sqlite3
import sys
from datetime import datetime

import xlwings as xw
from dotenv import load_dotenv

# 載入自訂模組/函式
from mod_excel_access import (
    convert_to_excel_address,
    ensure_sheet_exists,
    excel_address_to_row_col,
)

# =========================================================================
# 載入環境變數
# =========================================================================
load_dotenv()

DB_HO_LOK_UE = os.getenv('DB_HO_LOK_Ue', 'Ho_Lok_Ue.db')

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
# 台羅拼音 → 台語音標（TL → TLPA）轉換函數
# =========================================================================
def convert_tl_to_tlpa(im_piau):
    """
    轉換台羅拼音（TL）為台語音標（TLPA）。

    :param im_piau: 台羅拼音 (如 "tsua7")
    :return: 台語音標 (如 "zua7")
    """
    if not im_piau:
        return ""

    # 先替換較長的 "tsh" → "c"，避免 "ts" 被誤轉換
    im_piau = re.sub(r'\btsh', 'c', im_piau)  # tsh → c
    im_piau = re.sub(r'\bts', 'z', im_piau)   # ts → z

    return im_piau


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
# 功能 2：使用【標音字庫】更新【Ho_Lok_Ue.db】資料庫（含拼音轉換）
# =========================================================================
def update_database_from_excel(wb):
    """
    使用【標音字庫】工作表的資料更新 SQLite 資料庫（轉換台羅拼音 → 台語音標）。

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
        for idx, row_data in enumerate(data, start=2):  # Excel A2 起始，Python Index 2
            han_ji = row_data[0]  # A 欄
            tai_lo_im_piau = row_data[3]  # D 欄 (校正音標)

            if not han_ji or not tai_lo_im_piau or tai_lo_im_piau == "N/A":
                continue  # 跳過無效資料

            # **轉換台羅拼音（TL）→ 台語音標（TLPA）**
            tlpa_im_piau = convert_tl_to_tlpa(tai_lo_im_piau)

            # **在 INSERT 之前，顯示 Console 訊息**
            print(f"📌 寫入資料庫: 漢字='{han_ji}', 台羅拼音='{tai_lo_im_piau}', 轉換後 TLPA='{tlpa_im_piau}', Excel 第 {idx} 列")

            cursor.execute("""
                INSERT INTO 漢字庫 (漢字, 台羅音標, 常用度, 更新時間)
                VALUES (?, ?, ?, CURRENT_TIMESTAMP)
                ON CONFLICT(漢字, 台羅音標) DO UPDATE
                SET 更新時間=CURRENT_TIMESTAMP;
            """, (han_ji, tlpa_im_piau, 0.8))  # 常用度固定為 0.8

        conn.commit()
        print("✅ 資料庫更新完成！")
        return EXIT_CODE_SUCCESS

    except Exception as e:
        print(f"❌ 資料庫更新失敗: {e}")
        return EXIT_CODE_FAILURE

    finally:
        conn.close()


# =========================================================================
# 功能 3：將【漢字庫】資料表匯出到 Excel 的【漢字庫】工作表
# =========================================================================
def export_database_to_excel(wb):
    """
    將 `漢字庫` 資料表的資料寫入 Excel 的【漢字庫】工作表。

    :param wb: Excel 活頁簿物件
    :return: EXIT_CODE_SUCCESS or EXIT_CODE_FAILURE
    """
    sheet_name = "漢字庫"
    ensure_sheet_exists(wb, sheet_name)
    sheet = wb.sheets[sheet_name]

    conn = sqlite3.connect(DB_HO_LOK_UE)
    cursor = conn.cursor()

    try:
        # 讀取資料庫內容
        cursor.execute("SELECT 識別號, 漢字, 台羅音標, 常用度, 更新時間 FROM 漢字庫;")
        rows = cursor.fetchall()

        # 清空舊內容
        sheet.clear()

        # 寫入標題列
        sheet.range("A1").value = ["識別號", "漢字", "台羅音標", "常用度", "更新時間"]

        # 寫入資料
        sheet.range("A2").value = rows

        print("✅ 資料成功匯出至 Excel！")
        return EXIT_CODE_SUCCESS

    except Exception as e:
        print(f"❌ 匯出資料失敗: {e}")
        return EXIT_CODE_FAILURE

    finally:
        conn.close()


# =========================================================================
# 主程式執行
# =========================================================================
def main():
    if len(sys.argv) > 1:
        mode = sys.argv[1]
    else:
        mode = "1"

    wb = xw.apps.active.books.active

    if mode == "1":
        return update_pronunciation_in_excel(wb)
    elif mode == "2":
        return update_database_from_excel(wb)
    elif mode == "3":
        return export_database_to_excel(wb)
    else:
        print("❌ 錯誤：請輸入有效模式 (1, 2, 3)")
        return EXIT_CODE_INVALID_INPUT


if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)
