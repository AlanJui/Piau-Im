# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import logging
import os
import sqlite3
import sys
from datetime import datetime
from pathlib import Path

# 載入第三方套件
import xlwings as xw
from dotenv import load_dotenv

# 載入自訂模組/函式
from mod_excel_access import ensure_sheet_exists, get_value_by_name

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
EXIT_CODE_NO_FILE = 1
EXIT_CODE_INVALID_INPUT = 2
EXIT_CODE_PROCESS_FAILURE = 3
EXIT_CODE_UNKNOWN_ERROR = 99

# =========================================================================
# 台羅拼音 → 台語音標（TL → TLPA）轉換函數
# =========================================================================
def convert_tl_to_tlpa(pinyin):
    """
    轉換台羅拼音（TL）為台語音標（TLPA）。

    :param pinyin: 台羅拼音 (如 "tsua7")
    :return: 台語音標 (如 "zua7")
    """
    if not pinyin:
        return ""

    pinyin = pinyin.strip().lower()

    # 替換較長的 "tsh" → "c"，避免 "ts" 被誤轉換
    pinyin = pinyin.replace("tsh", "c")  # tsh → c
    pinyin = pinyin.replace("ts", "z")   # ts → z

    return pinyin


# =========================================================================
# 更新 `漢字庫` 資料表
# =========================================================================
def insert_or_update_to_db(db_path, han_ji: str, tai_lo_pinyin: str, piau_im_huat: str):
    """
    插入或更新 `漢字庫` 資料表。

    :param db_path: 資料庫檔案路徑。
    :param han_ji: 漢字。
    :param tai_lo_pinyin: 台羅拼音（TL）。
    :param piau_im_huat: 音讀類型（文讀音 or 白話音）。
    """
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    # 確保 `漢字庫` 資料表存在
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS 漢字庫 (
        識別號 INTEGER PRIMARY KEY AUTOINCREMENT,
        漢字 TEXT NOT NULL,
        台羅音標 TEXT NOT NULL,
        常用度 REAL DEFAULT 0.8,
        更新時間 TEXT DEFAULT (DATETIME('now', 'localtime')) NOT NULL
    );
    """)

    # 確保 `台羅音標` 為 `TLPA`
    tlpa_pinyin = convert_tl_to_tlpa(tai_lo_pinyin)

    # 確定 `常用度`（文讀音 0.8 / 白話音 0.6）
    siong_iong_too = 0.8 if piau_im_huat == "文讀音" else 0.6

    # **嘗試插入或更新**
    cursor.execute("""
        INSERT INTO 漢字庫 (漢字, 台羅音標, 常用度, 更新時間)
        VALUES (?, ?, ?, CURRENT_TIMESTAMP)
        ON CONFLICT(漢字, 台羅音標) DO UPDATE
        SET 更新時間 = CURRENT_TIMESTAMP;
    """, (han_ji, tlpa_pinyin, siong_iong_too))

    conn.commit()
    conn.close()

    print(f"✅ 成功寫入資料庫: {han_ji} -> {tlpa_pinyin} (常用度: {siong_iong_too})")


# =========================================================================
# 讀取 Excel 的【缺字表】工作表，並回填至 `漢字庫`
# =========================================================================
def process_excel_to_db(wb, sheet_name, db_path):
    """
    讀取 Excel 的【缺字表】工作表，並將資料回填至 SQLite `漢字庫`。

    :param wb: Excel 活頁簿物件。
    :param sheet_name: Excel 工作表名稱。
    :param db_path: 資料庫檔案路徑。
    """
    sheet = wb.sheets[sheet_name]
    piau_im_huat = get_value_by_name(wb=wb, name="語音類型")

    data = sheet.range("A2").expand("table").value

    # 確保資料為 2D 列表
    if not isinstance(data[0], list):
        data = [data]

    for row in data:
        han_ji = row[0] or ""
        tai_lo_pinyin = row[2] or ""

        if han_ji and tai_lo_pinyin:
            insert_or_update_to_db(db_path, han_ji, tai_lo_pinyin, piau_im_huat)

    print(f"✅ 【缺字表】已成功回填至資料庫 `{db_path}`")


# =============================================================================
# 主流程
# =============================================================================
def process(wb):
    sheet_name = "缺字表"
    db_path = DB_HO_LOK_UE

    process_excel_to_db(wb, sheet_name, db_path)
    return EXIT_CODE_SUCCESS


# =============================================================================
# 主執行函數
# =============================================================================
def main():
    logging.info("🔹 作業開始")

    wb = None
    try:
        wb = xw.apps.active.books.active
    except Exception as e:
        print(f"⚠️ 找不到作用中的 Excel 活頁簿: {e}")
        return EXIT_CODE_NO_FILE

    if not wb:
        return EXIT_CODE_NO_FILE

    try:
        return process(wb)
    except Exception as e:
        print(f"❌ 進行過程發生錯誤: {e}")
        return EXIT_CODE_UNKNOWN_ERROR


if __name__ == "__main__":
    sys.exit(main())
