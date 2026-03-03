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

DB_HO_LOK_UE = os.getenv("DB_HO_LOK_UE", "Ho_Lok_Ue.db")

# =========================================================================
# 設定日誌
# =========================================================================
logging.basicConfig(
    filename="process_log.txt",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
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
# 作業程序
# =========================================================================
def get_active_cell_info(wb):
    """
    取得目前 Excel 作用儲存格的資訊：
    - 作用儲存格的位置 (row, col)
    - 取得【漢字】的值
    - 計算【人工標音】儲存格位置，並取得【人工標音】值

    :param wb: Excel 活頁簿物件
    :return: (sheet_name, han_ji, (row, col), artificial_pronounce, (artificial_row, col))
    """
    active_cell = wb.app.selection  # 取得目前作用中的儲存格
    sheet_name = active_cell.sheet.name  # 取得所在的工作表名稱
    cell_address = active_cell.address.replace(
        "$", ""
    )  # 取得 Excel 格式地址 (去掉 "$")

    row, col = excel_address_to_row_col(cell_address)  # 轉換為 (row, col)

    # 取得【漢字】 (作用儲存格的值)
    han_ji = active_cell.value

    # 計算【人工標音】位置 (row-2, col) 並取得其值
    artificial_row = row - 2
    artificial_cell = wb.sheets[sheet_name].cells(artificial_row, col)
    artificial_pronounce = artificial_cell.value  # 取得人工標音的值

    return sheet_name, han_ji, (row, col), artificial_pronounce, (artificial_row, col)


# def excel_address_to_row_col(cell_address):
#     """
#     將 Excel 儲存格地址 (如 'D9') 轉換為 (row, col) 格式。

#     :param cell_address: Excel 儲存格地址 (如 'D9', 'AA15')
#     :return: (row, col) 元組，例如 (9, 4)
#     """
#     match = re.match(r"([A-Z]+)(\d+)", cell_address)  # 用 regex 拆分字母(列) 和 數字(行)

#     if not match:
#         raise ValueError(f"無效的 Excel 儲存格地址: {cell_address}")

#     col_letters, row_number = match.groups()

#     # 將 Excel 字母列轉換成數字，例如 A -> 1, B -> 2, ..., Z -> 26, AA -> 27
#     col_number = 0
#     for letter in col_letters:
#         col_number = col_number * 26 + (ord(letter) - ord("A") + 1)

#     return int(row_number), col_number


def check_and_update_pronunciation(wb, han_ji, position, artificial_pronounce):
    """
    查詢【標音字庫】工作表，確認是否有該【漢字】與【座標】，
    且【校正音標】是否為 'N/A'，若符合則更新為【人工標音】。

    :param wb: Excel 活頁簿物件
    :param han_ji: 查詢的漢字
    :param position: (row, col) 該漢字的座標
    :param artificial_pronounce: 需要更新的【人工標音】
    :return: 是否更新成功 (True/False)
    """
    sheet_name = "標音字庫"

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

    for idx, row in enumerate(data):
        row_han_ji = row[0]  # A 欄: 漢字
        correction_pronounce_cell = sheet.range(f"D{idx+2}")  # D 欄: 校正音標
        coordinates = row[4]  # E 欄: 座標 (可能是 "(9, 4); (25, 9)" 這類格式)

        if row_han_ji == han_ji and coordinates:
            # 將座標解析成一個 set
            coord_list = coordinates.split("; ")
            parsed_coords = {convert_to_excel_address(coord) for coord in coord_list}

            # 確認該座標是否存在於【標音字庫】中
            if convert_to_excel_address(str(position)) in parsed_coords:
                # 檢查標正音標是否為 'N/A'
                if correction_pronounce_cell.value == "N/A":
                    # 更新【校正音標】為【人工標音】
                    correction_pronounce_cell.value = artificial_pronounce
                    print(
                        f"✅ 更新成功: {han_ji} ({position}) -> {artificial_pronounce}"
                    )
                    return True

    print(f"❌ 未找到匹配的資料或不符合更新條件: {han_ji} ({position})")
    return False


# =========================================================================
# 台羅拼音 → 台語音標（TL → TLPA）轉換函數
# =========================================================================
def convert_tl_to_tlpa(im_piau):
    """
    轉換台羅拼音（TL）為台語音標（TLPA）。
    """
    if not im_piau:
        return ""
    im_piau = re.sub(r"\btsh", "c", im_piau)  # tsh → c
    im_piau = re.sub(r"\bts", "z", im_piau)  # ts → z
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
                    print(
                        f"✅ 更新成功: {han_ji} ({row}, {col}) -> {artificial_pronounce}"
                    )
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
            print(
                f"📌 寫入資料庫: 漢字='{han_ji}', 台羅拼音='{tai_lo_im_piau}', 轉換後 TLPA='{tlpa_im_piau}', Excel 第 {idx} 列"
            )

            cursor.execute(
                """
                INSERT INTO 漢字庫 (漢字, 台羅音標, 常用度, 更新時間)
                VALUES (?, ?, ?, CURRENT_TIMESTAMP)
                ON CONFLICT(漢字, 台羅音標) DO UPDATE
                SET 更新時間=CURRENT_TIMESTAMP;
            """,
                (han_ji, tlpa_im_piau, 0.8),
            )  # 常用度固定為 0.8

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
        # cursor.execute("SELECT 識別號, 漢字, 台羅音標, 常用度, 摘要說明, 更新時間 FROM 漢字庫;")
        cursor.execute(
            "SELECT 識別號, 漢字, 台羅音標, 常用度, 摘要說明, 更新時間 FROM 漢字庫R1;"
        )
        rows = cursor.fetchall()

        # 清空舊內容
        sheet.clear()

        # 寫入標題列
        sheet.range("A1").value = [
            "識別號",
            "漢字",
            "台羅音標",
            "常用度",
            "摘要說明" "更新時間",
        ]

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
# 功能 4：重建 `漢字庫` 資料表（補上 `摘要說明` 欄位）
# =========================================================================
def rebuild_database_from_excel(wb):
    """
    依據 Excel【漢字庫】工作表，重建 `漢字庫` 資料表（包含 `摘要說明` 欄位）。

    :param wb: Excel 活頁簿物件
    :return: EXIT_CODE_SUCCESS or EXIT_CODE_FAILURE
    """
    sheet_name = "漢字庫"
    ensure_sheet_exists(wb, sheet_name)
    sheet = wb.sheets[sheet_name]

    conn = sqlite3.connect(DB_HO_LOK_UE)
    cursor = conn.cursor()

    try:
        # **1️⃣ 刪除現有 `漢字庫` 資料表**
        cursor.execute("DROP TABLE IF EXISTS 漢字庫")

        # **2️⃣ 重新建立 `漢字庫` 資料表**
        cursor.execute(
            """
        CREATE TABLE 漢字庫 (
            識別號 INTEGER PRIMARY KEY AUTOINCREMENT,
            漢字 TEXT NOT NULL,
            台羅音標 TEXT NOT NULL,
            常用度 REAL DEFAULT 0.1,
            摘要說明 TEXT DEFAULT 'NA',
            更新時間 TEXT DEFAULT (DATETIME('now', 'localtime')) NOT NULL
        );
        """
        )

        # **3️⃣ 讀取 Excel `漢字庫` 工作表**
        data = sheet.range("A2").expand("table").value
        if not isinstance(data[0], list):
            data = [data]

        # **4️⃣ 新增資料**
        for idx, row_data in enumerate(data, start=2):
            han_ji = row_data[1]  # B 欄
            tai_lo_im_piau = row_data[2]  # C 欄
            siong_iong_too = (
                row_data[3] if isinstance(row_data[3], (int, float)) else 0.1
            )  # D 欄
            summary = (
                row_data[4] if isinstance(row_data[4], str) else "NA"
            )  # E 欄（摘要）
            updated_time = (
                row_data[5]
                if isinstance(row_data[5], str)
                else datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            )

            # **Console Debug 訊息**
            print(
                f"📌 正在處理第 {idx-1} 筆資料 (Excel 第 {idx} 列): 漢字='{han_ji}', 台羅音標='{tai_lo_im_piau}', 更新時間='{updated_time}'"
            )

            # **確保 `漢字` 和 `台羅音標` 務必要有資料**
            if not han_ji or not tai_lo_im_piau:
                print(f"⚠️ 跳過無效資料: Excel 第 {idx} 列：缺【漢字】或【台羅音標】")
                # **將錯誤記錄寫入 `error_log.txt`**
                with open("error_log.txt", "a", encoding="utf-8") as log_file:
                    log_file.write(f"❌ 無效資料（Excel 第 {idx} 列）: {row_data}\n")
                continue  # 跳過無效資料

            # **檢查 `台羅音標` 是否為有效字串**
            if (
                not han_ji
                or not isinstance(tai_lo_im_piau, str)
                or not tai_lo_im_piau.strip()
            ):
                print(f"⚠️ 跳過無效資料: Excel 第 {idx} 列 (台羅音標格式錯誤)")
                # **將錯誤記錄寫入 `error_log.txt`**
                with open("error_log.txt", "a", encoding="utf-8") as log_file:
                    log_file.write(f"❌ 無效資料（Excel 第 {idx} 列）: {row_data}\n")
                continue  # **跳過此筆錯誤資料**

            # 轉換台羅拼音（TL）→ 台語音標（TLPA）
            # tlpa_pinyin = convert_tl_to_tlpa(tai_lo_im_piau)

            cursor.execute(
                """
                INSERT INTO 漢字庫 (漢字, 台羅音標, 常用度, 摘要說明, 更新時間)
                VALUES (?, ?, ?, ?, ?);
            """,
                (han_ji, tai_lo_im_piau, siong_iong_too, summary, updated_time),
            )

        # **5️⃣ 建立 `UNIQUE INDEX` 確保無重複**
        cursor.execute(
            "CREATE UNIQUE INDEX idx_漢字_台羅音標 ON 漢字庫 (漢字, 台羅音標);"
        )

        conn.commit()
        print("✅ `漢字庫` 資料表已成功重建！")
        return EXIT_CODE_SUCCESS

    except Exception as e:
        print(f"❌ 重建 `漢字庫` 失敗: {e}")
        return EXIT_CODE_FAILURE

    finally:
        conn.close()


# =========================================================================
# 功能 5：匯出成 RIME 輸入法字典
# =========================================================================
def export_to_rime_dict():
    """
    將 `漢字庫` 資料表轉換成 RIME 輸入法字典格式（YAML）。
    """
    conn = sqlite3.connect(DB_HO_LOK_UE)
    cursor = conn.cursor()

    try:
        cursor.execute("SELECT 漢字, 台羅音標, 常用度, 摘要說明, 更新時間 FROM 漢字庫;")
        rows = cursor.fetchall()

        dict_filename = "tl_ji_khoo_peh_ue.dict.yaml"
        with open(dict_filename, "w", encoding="utf-8") as file:
            # 寫入 RIME 字典檔頭
            file.write("# Rime dictionary\n")
            file.write("# encoding: utf-8\n")
            file.write("#\n# 河洛白話音\n#\n")
            file.write("---\n")
            file.write("name: tl_ji_khoo_peh_ue\n")
            file.write('version: "v0.1.0.0"\n')
            file.write("sort: by_weight\n")
            file.write("use_preset_vocabulary: false\n")
            file.write("columns:\n")
            file.write("  - text    # 漢字\n")
            file.write("  - code    # 台灣音標（TLPA）拼音\n")
            file.write("  - weight  # 常用度（優先顯示度）\n")
            file.write("  - stem    # 用法舉例\n")
            file.write("  - create  # 建立時間\n")
            file.write("import_tables:\n")
            file.write("  - tl_ji_khoo_kah_kut_bun\n")
            file.write("...\n")

            # **寫入字典內容**
            for han_ji, tai_lo_pinyin, weight, summary, create_time in rows:
                file.write(
                    f"{han_ji}\t{tai_lo_pinyin}\t{weight}\t{summary}\t{create_time}\n"
                )

        print(f"✅ RIME 字典 `{dict_filename}` 匯出完成！")
        return EXIT_CODE_SUCCESS
    except Exception as e:
        print(f"❌ 匯出 RIME 字典失敗: {e}")
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
        mode = "3"

    if mode == "5":
        return export_to_rime_dict()

    wb = xw.apps.active.books.active

    if mode == "1":
        return update_pronunciation_in_excel(wb)
    elif mode == "2":
        return update_database_from_excel(wb)
    elif mode == "3":
        return export_database_to_excel(wb)
    elif mode == "4":
        return rebuild_database_from_excel(wb)
    else:
        print("❌ 錯誤：請輸入有效模式 (1, 2, 3, 4, 5)")
        return EXIT_CODE_INVALID_INPUT


if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)
