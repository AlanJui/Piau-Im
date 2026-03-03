"""
a500_漢字庫工具.py v0.2

變更紀錄：
- v0.2 2026-03-03:
    - 新增功能：功能 5：匯出成 RIME 輸入法字典，需衣【漢字】及【台語音標】兩欄排序，以令同一【漢字】的不同【台語音標】之資料列，能夠相鄰排列；
    - 新增功能：功能 6：將資料庫之【漢字庫】資料表，匯出至 Excel 的工作表。
"""

# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import os
import re
import sqlite3
import sys
from datetime import datetime
from pathlib import Path

import xlwings as xw
from dotenv import load_dotenv

# 載入自訂模組/函式
from mod_excel_access import (
    convert_to_excel_address,
    ensure_sheet_exists,
    excel_address_to_row_col,
    get_value_by_name,
)
from mod_帶調符音標 import tng_im_piau, tng_tiau_ho
from mod_標音 import convert_tlpa_to_tl, split_tai_gi_im_piau

# =========================================================================
# 常數定義
# =========================================================================
# 定義 Exit Code
EXIT_CODE_SUCCESS = 0  # 成功
EXIT_CODE_NO_FILE = 1  # 無法找到檔案
EXIT_CODE_INVALID_INPUT = 2  # 輸入錯誤
EXIT_CODE_SAVE_FAILURE = 3  # 儲存失敗
EXIT_CODE_PROCESS_FAILURE = 10  # 過程失敗
EXIT_CODE_UNKNOWN_ERROR = 99  # 未知錯誤

# =========================================================================
# 載入環境變數
# =========================================================================
load_dotenv()

# 預設檔案名稱從環境變數讀取
DB_HO_LOK_UE = os.getenv("DB_HO_LOK_UE", "Ho_Lok_Ue.db")
DB_KONG_UN = os.getenv("DB_KONG_UN", "Kong_Un.db")

# =========================================================================
# 設定日誌
# =========================================================================
from mod_logging import init_logging, logging_exc_error, logging_process_step

init_logging()


# =========================================================================
# 程式區域函式
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
        # correction_pronounce_cell = sheet.range(f"C{idx+2}")  # C 欄: 校正音標
        correction_pronounce_cell = sheet.range(f"B{idx+2}")  # B 欄: 台語音標
        coordinates = row[3]  # D 欄: 座標 (可能是 "(9, 4); (25, 9)" 這類格式)

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


def insert_or_update_to_db(
    db_path, table_name: str, han_ji: str, tai_gi_im_piau: str, piau_im_huat: str
):
    """
    將【漢字】與【台語音標】插入或更新至資料庫。

    :param db_path: 資料庫檔案路徑。
    :param table_name: 資料表名稱。
    :param han_ji: 漢字。
    :param tai_gi_im_piau: 台語音標。
    """
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    # 確保資料表存在
    cursor.execute(
        f"""
    CREATE TABLE IF NOT EXISTS {table_name} (
        識別號 INTEGER NOT NULL UNIQUE PRIMARY KEY AUTOINCREMENT,
        漢字 TEXT,
        台羅音標 TEXT,
        常用度 REAL,
        摘要說明 TEXT,
        建立時間 TEXT NOT NULL DEFAULT (DATETIME('now', 'localtime')),
        更新時間 TEXT NOT NULL DEFAULT (DATETIME('now', 'localtime'))
    );
    """
    )

    # 檢查是否已存在該漢字
    cursor.execute(f"SELECT 識別號 FROM {table_name} WHERE 漢字 = ?", (han_ji,))
    row = cursor.fetchone()

    siong_iong_too = 0.8 if piau_im_huat == "文讀音" else 0.6
    tai_lo_im_piau = convert_tlpa_to_tl(tai_gi_im_piau)
    if row:
        # 更新資料
        cursor.execute(
            f"""
        UPDATE {table_name}
        SET 台羅音標 = ?, 更新時間 = ?
        WHERE 識別號 = ?;
        """,
            (tai_lo_im_piau, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), row[0]),
        )
    else:
        # 若語音類型為：【文讀音】，設定【常用度】欄位值為 0.8
        cursor.execute(
            f"""
        INSERT INTO {table_name} (漢字, 台羅音標, 常用度, 摘要說明)
        VALUES (?, ?, ?, NULL);
        """,
            (han_ji, tai_gi_im_piau, siong_iong_too),
        )

    conn.commit()
    conn.close()


# =========================================================================
# 功能 1：使用【人工標音】更新【標音字庫】的校正音標
# =========================================================================
def khuat_ji_piau_poo_im_piau(wb):
    """
    讀取 Excel 的【缺字表】工作表，並將資料回填至 SQLite 資料庫。

    :param excel_path: Excel 檔案路徑。
    :param sheet_name: Excel 工作表名稱。
    :param db_path: 資料庫檔案路徑。
    :param table_name: 資料表名稱。
    """
    sheet_name = "缺字表"
    sheet = wb.sheets[sheet_name]
    piau_im_huat = get_value_by_name(wb=wb, name="語音類型")
    db_path = "Ho_Lok_Ue.db"  # 替換為你的資料庫檔案路徑
    table_name = "漢字庫"  # 替換為你的資料表名稱
    hue_im = wb.names["語音類型"].refers_to_range.value
    siong_iong_too = 0.8 if hue_im == "文讀音" else 0.6  # 根據語音類型設定常用度

    # 讀取資料表範圍
    data = sheet.range("A2").expand("table").value

    # 確保資料為 2D 列表
    if not isinstance(data[0], list):
        data = [data]

    idx = 0
    for row in data:
        han_ji = row[0]  # 漢字
        # zong_siau = row[1] # 總數
        tai_gi_im_piau = row[1]  # 台語音標
        hau_ziann_im_piau = row[2]  # 台語音標
        zo_piau = row[3]  # (儲存格位置)座標

        if han_ji and (tai_gi_im_piau != "N/A" or hau_ziann_im_piau != "N/A"):
            # 將 Excel 工作表存放的【台語音標（TLPA）】，改成資料庫保存的【台羅拼音（TL）】
            tlpa_im_piau = tng_im_piau(
                tai_gi_im_piau
            )  # 將【音標】使用之【拼音字母】轉換成【TLPA拼音字母】；【音標調符】仍保持
            tlpa_im_piau_cleanned = tng_tiau_ho(
                tlpa_im_piau
            ).lower()  # 將【音標調符】轉換成【數值調號】
            tl_im_piau = convert_tlpa_to_tl(tlpa_im_piau_cleanned)

            insert_or_update_to_db(
                db_path, table_name, han_ji, tl_im_piau, piau_im_huat
            )
            print(
                f"📌 {idx+1}. 【{han_ji}】：台羅音標：【{tl_im_piau}】、校正音標：【{hau_ziann_im_piau}】、台語音標=【{tai_gi_im_piau}】、座標：{zo_piau}"
            )
            idx += 1

    logging_process_step(
        f"【缺字表】中的資料已成功回填至資料庫： {db_path} 的【{table_name}】資料表中。"
    )
    return EXIT_CODE_SUCCESS


# =========================================================================
# 功能 2：使用【標音字庫】更新【Ho_Lok_Ue.db】資料庫（含拼音轉換）
# =========================================================================
def update_database_from_excel(wb):
    """
    使用【標音字庫】工作表的資料更新 SQLite 資料庫（轉換台羅拼音 → 台語音標）。

    - 如果資料庫中已存在相同的【漢字】和【台羅音標】，則更新【常用度】欄位。
    - 如果資料庫中不存在相同的【漢字】和【台羅音標】，則新增一筆資料。

    :param wb: Excel 活頁簿物件
    :return: EXIT_CODE_SUCCESS or EXIT_CODE_FAILURE
    """
    sheet_name = "標音字庫"
    sheet = wb.sheets[sheet_name]
    data = sheet.range("A2").expand("table").value
    hue_im = wb.names["語音類型"].refers_to_range.value
    siong_iong_too = 0.8 if hue_im == "文讀音" else 0.6  # 根據語音類型設定常用度

    if not isinstance(data[0], list):
        data = [data]

    conn = sqlite3.connect(DB_HO_LOK_UE)
    cursor = conn.cursor()

    try:
        for idx, row_data in enumerate(data, start=2):  # Excel A2 起始，Python Index 2
            han_ji = row_data[0]  # A 欄 (漢字)
            # hau_ziann_im_piau = row_data[2]  # C 欄 (校正音標)
            tai_gi_im_piau = row_data[1]  # B 欄 (台語音標)

            # 跳過無效資料
            if not han_ji or not tai_gi_im_piau or tai_gi_im_piau == "N/A":
                continue

            # 將 Excel 工作表存放的【台語音標（TLPA）】，改成資料庫保存的【台羅拼音（TL）】
            tai_lo_im_piau = convert_tlpa_to_tl(tai_gi_im_piau)

            # 檢查【台語音標】與【校正音標】兩欄位，若【校正音標】非空白，則以此欄位之【音標】輸入資料庫

            # 檢查資料庫中是否已存在相同的【漢字】和【台羅音標】
            cursor.execute(
                """
                SELECT 1 FROM 漢字庫
                WHERE 漢字 = ? AND 台羅音標 = ?
            """,
                (han_ji, tai_lo_im_piau),
            )
            exists = cursor.fetchone()

            if exists:
                # 如果存在，更新【常用度】欄位
                cursor.execute(
                    """
                    UPDATE 漢字庫
                    SET 常用度 = ?, 更新時間 = CURRENT_TIMESTAMP
                    WHERE 漢字 = ? AND 台羅音標 = ?
                """,
                    (siong_iong_too, han_ji, tai_lo_im_piau),
                )
                print(
                    f"🔄 更新資料庫: 漢字='{han_ji}'，台語音標='{tai_gi_im_piau}' ==> 台羅音標='{tai_lo_im_piau}', 常用度={siong_iong_too}, Excel 第 {idx} 列"
                )
            else:
                # 如果不存在，新增一筆資料
                cursor.execute(
                    """
                    INSERT INTO 漢字庫 (漢字, 台羅音標, 常用度, 更新時間)
                    VALUES (?, ?, ?, CURRENT_TIMESTAMP)
                """,
                    (han_ji, tai_lo_im_piau, siong_iong_too),
                )
                print(
                    f"📌 新增資料庫: 漢字='{han_ji}'，台語音標='{tai_gi_im_piau}' ==> 台羅音標='{tai_lo_im_piau}', 常用度={siong_iong_too}, Excel 第 {idx} 列"
                )

        conn.commit()
        print("✅ 資料庫更新完成！")

    except Exception as e:
        print(f"❌ 資料庫更新失敗: {e}")
        return EXIT_CODE_PROCESS_FAILURE

    finally:
        conn.close()

    logging_process_step(f"【標音字庫】中的【漢字】與【台語音標】已成功回填資料庫！")
    return EXIT_CODE_SUCCESS


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

    except Exception as e:
        print(f"❌ 匯出資料失敗: {e}")
        return EXIT_CODE_PROCESS_FAILURE

    finally:
        conn.close()

    logging_process_step(
        f"已將資料庫之【漢字庫】資料表，匯出至 Excel 作用中活頁簿檔的【漢字庫】工作表！"
    )
    return EXIT_CODE_SUCCESS


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
            tai_gi_im_piau = row_data[2]  # C 欄
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
                f"📌 正在處理第 {idx-1} 筆資料 (Excel 第 {idx} 列): 漢字='{han_ji}', 台語音標='{tai_gi_im_piau}', 更新時間='{updated_time}'"
            )

            # **確保 `漢字` 和 `台羅音標` 務必要有資料**
            if not han_ji or not tai_gi_im_piau:
                print(f"⚠️ 跳過無效資料: Excel 第 {idx} 列：缺【漢字】或【台羅音標】")
                # **將錯誤記錄寫入 `error_log.txt`**
                with open("error_log.txt", "a", encoding="utf-8") as log_file:
                    log_file.write(f"❌ 無效資料（Excel 第 {idx} 列）: {row_data}\n")
                continue  # 跳過無效資料

            # **檢查 `台羅音標` 是否為有效字串**
            if (
                not han_ji
                or not isinstance(tai_gi_im_piau, str)
                or not tai_gi_im_piau.strip()
            ):
                print(f"⚠️ 跳過無效資料: Excel 第 {idx} 列 (台羅音標格式錯誤)")
                # **將錯誤記錄寫入 `error_log.txt`**
                with open("error_log.txt", "a", encoding="utf-8") as log_file:
                    log_file.write(f"❌ 無效資料（Excel 第 {idx} 列）: {row_data}\n")
                continue  # **跳過此筆錯誤資料**

            # 將 Excel 工作表存放的【台語音標（TLPA）】，改成資料庫保存的【台羅拼音（TL）】
            tai_lo_im_piau = convert_tlpa_to_tl(tai_gi_im_piau)

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

    except Exception as e:
        print(f"❌ 重建 `漢字庫` 失敗: {e}")
        return EXIT_CODE_PROCESS_FAILURE

    finally:
        conn.close()

    logging_process_step(
        f"自【作用中活頁簿】檔之【漢字庫】工作表，匯入資料進資料庫之【漢字庫】資料表！"
    )
    return EXIT_CODE_SUCCESS


# =========================================================================
# 功能 5：匯出成 RIME 輸入法字典
# =========================================================================
def export_to_rime_dict():
    """
    將 `漢字庫` 資料表轉換成 RIME 輸入法字典格式（YAML）。
    """
    conn = sqlite3.connect(DB_HO_LOK_UE)
    cursor = conn.cursor()
    # dict_filename = "tl_ji_khoo_peh_ue.dict.yaml"
    dict_filename = "tl_han_ji_khoo.dict.yaml"

    try:
        cursor.execute(
            "SELECT 漢字, 台羅音標, 常用度, 摘要說明, 更新時間 FROM 漢字庫 "
            "ORDER BY 漢字 ASC, 台羅音標 ASC;"
        )
        rows = cursor.fetchall()

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
            # file.write("  - tl_ji_khoo_kah_kut_bun\n")
            file.write(f"  - {dict_filename.replace('.dict.yaml', '')}\n")
            file.write("...\n")

            # **寫入字典內容**
            for han_ji, tai_lo_pinyin, weight, summary, create_time in rows:
                file.write(
                    f"{han_ji}\t{tai_lo_pinyin}\t{weight}\t{summary}\t{create_time}\n"
                )

        print(f"✅ RIME 字典 `{dict_filename}` 匯出完成！")
    except Exception as e:
        print(f"❌ 匯出 RIME 字典失敗: {e}")
        return EXIT_CODE_PROCESS_FAILURE
    finally:
        conn.close()

    logging_process_step(f"已將資料庫之【漢字庫】資料表，匯出並製成【中州韻字典檔】！")
    return EXIT_CODE_SUCCESS


# =========================================================================
# 功能 6：匯出 SQLite 漢字庫至 Excel 的【SQLite漢字庫】工作表
# =========================================================================
def export_sqlite_db_to_excel(wb):
    """
    將 漢字庫 資料表的資料寫入 Excel 的【SQLite漢字庫】工作表。

    【註】：如果【SQLite漢字庫】工作表已存在，則將原有之工作表更名，補上（BAK）；
    然後再建立新工表。

    A欄：識別號
    B欄：漢字
    C欄：台羅音標
    D欄：常用度
    E欄：聲母
    F欄：韻母
    G欄：調號

    :return: EXIT_CODE_SUCCESS or EXIT_CODE_FAILURE
    """
    sheet_name = "SQLite漢字庫"
    bak_sheet_name = f"{sheet_name}(BAK)"

    sheet_names = [sheet.name for sheet in wb.sheets]
    if sheet_name in sheet_names:
        if bak_sheet_name in sheet_names:
            wb.sheets[bak_sheet_name].delete()
        wb.sheets[sheet_name].name = bak_sheet_name

    sheet = wb.sheets.add(sheet_name)

    conn = sqlite3.connect(DB_HO_LOK_UE)
    cursor = conn.cursor()

    try:
        # 讀取資料庫內容
        cursor.execute(
            "SELECT 識別號, 漢字, 台羅音標, 常用度 FROM 漢字庫 ORDER BY 漢字 ASC, 台羅音標 ASC;"
        )
        rows = cursor.fetchall()

        # 寫入標題列
        sheet.range("A1").value = [
            "識別號",
            "漢字",
            "台語音標",
            "常用度",
            "聲母",
            "韻母",
            "調號",
        ]

        out_rows = []
        for row in rows:
            it_id, han_ji, tl_im_piau, siong_iong = row

            # 將資料庫中的【台羅音標】先轉換成【台語音標】
            tlpa_im_piau = convert_tl_to_tlpa(tl_im_piau)

            # 將台語音標拆解成聲母、韻母、調號
            # 注意：這裡使用從 mod_標音 引入的 split_tai_gi_im_piau 處理單一字串
            siann_bu, un_bu, tiau_ho = split_tai_gi_im_piau(tlpa_im_piau)

            out_rows.append(
                [it_id, han_ji, tlpa_im_piau, siong_iong, siann_bu, un_bu, tiau_ho]
            )

        # 寫入資料（使用範圍賦值加速）
        if out_rows:
            sheet.range("A2").value = out_rows

        # 自動調整欄寬
        sheet.autofit("c")

        print("✅ 資料成功匯出至 Excel 的【SQLite漢字庫】工作表！")

    except Exception as e:
        print(f"❌ 匯出資料失敗: {e}")
        return EXIT_CODE_PROCESS_FAILURE

    finally:
        conn.close()

    logging_process_step(
        f"已將資料庫之【漢字庫】資料表，匯出至 Excel 作用中活頁簿檔的【SQLite漢字庫】工作表！"
    )
    return EXIT_CODE_SUCCESS


# =========================================================================
# 主程式執行
# =========================================================================
def main():
    # =========================================================================
    # (0) 程式初始化
    # =========================================================================
    # 取得專案根目錄。
    current_file_path = Path(__file__).resolve()
    project_root = current_file_path.parent
    # 取得程式名稱
    # program_file_name = current_file_path.name
    program_name = current_file_path.stem

    # =========================================================================
    # 程式初始化
    # =========================================================================
    logging_process_step(f"《========== 程式開始執行：{program_name} ==========》")
    logging_process_step(f"專案根目錄為: {project_root}")

    # =========================================================================
    # 開始執行程式
    # =========================================================================
    if len(sys.argv) > 1:
        mode = sys.argv[1]
    else:
        mode = "3"

    if mode == "5":
        return export_to_rime_dict()

    wb = xw.apps.active.books.active

    if mode == "1":
        return khuat_ji_piau_poo_im_piau(wb)
    elif mode == "2":
        return update_database_from_excel(wb)
    elif mode == "3":
        return export_database_to_excel(wb)
    elif mode == "4":
        return rebuild_database_from_excel(wb)
    elif mode == "6":
        return export_sqlite_db_to_excel(wb)
    else:
        print("❌ 錯誤：請輸入有效模式 (1, 2, 3, 4, 5, 6)")
        return EXIT_CODE_INVALID_INPUT


if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)
