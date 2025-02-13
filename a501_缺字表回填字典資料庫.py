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
from mod_excel_access import (
    create_dict_by_sheet,
    ensure_sheet_exists,
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
EXIT_CODE_NO_FILE = 1  # 無法找到檔案
EXIT_CODE_INVALID_INPUT = 2  # 輸入錯誤
EXIT_CODE_PROCESS_FAILURE = 3  # 過程失敗
EXIT_CODE_UNKNOWN_ERROR = 99  # 未知錯誤

# =========================================================================
# 作業程序
# =========================================================================
def insert_or_update_to_db(db_path, table_name: str, han_ji: str, tai_gi_im_piau: str, piau_im_huat: str):
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
    cursor.execute(f"""
    CREATE TABLE IF NOT EXISTS {table_name} (
        識別號 INTEGER NOT NULL UNIQUE PRIMARY KEY AUTOINCREMENT,
        漢字 TEXT,
        台羅音標 TEXT,
        常用度 REAL,
        摘要說明 TEXT,
        建立時間 TEXT NOT NULL DEFAULT (DATETIME('now', 'localtime')),
        更新時間 TEXT NOT NULL DEFAULT (DATETIME('now', 'localtime'))
    );
    """)

    # 檢查是否已存在該漢字
    cursor.execute(f"SELECT 識別號 FROM {table_name} WHERE 漢字 = ?", (han_ji,))
    row = cursor.fetchone()

    siong_iong_too = 0.8 if piau_im_huat == "文讀音" else 0.6
    if row:
        # 更新資料
        cursor.execute(f"""
        UPDATE {table_name}
        SET 台羅音標 = ?, 更新時間 = ?
        WHERE 識別號 = ?;
        """, (tai_gi_im_piau, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), row[0]))
    else:
        # 若語音類型為：【文讀音】，設定【常用度】欄位值為 0.8
        cursor.execute(f"""
        INSERT INTO {table_name} (漢字, 台羅音標, 常用度, 摘要說明)
        VALUES (?, ?, ?, NULL);
        """, (han_ji, tai_gi_im_piau, siong_iong_too))

    conn.commit()
    conn.close()


def process_excel_to_db(wb, sheet_name, db_path, table_name):
    """
    讀取 Excel 的【缺字表】工作表，並將資料回填至 SQLite 資料庫。

    :param excel_path: Excel 檔案路徑。
    :param sheet_name: Excel 工作表名稱。
    :param db_path: 資料庫檔案路徑。
    :param table_name: 資料表名稱。
    """
    # wb = xw.Book(excel_path)
    sheet = wb.sheets[sheet_name]
    piau_im_huat = get_value_by_name(wb=wb, name="語音類型")

    # 讀取資料表範圍
    data = sheet.range("A2").expand("table").value

    # 確保資料為 2D 列表
    if not isinstance(data[0], list):
        data = [data]

    for row in data:
        han_ji = row[0]
        tai_gi_im_piau = row[2]

        if han_ji and tai_gi_im_piau:
            insert_or_update_to_db(db_path, table_name, han_ji, tai_gi_im_piau, piau_im_huat)

    print(f"【缺字表】中的資料已成功回填至資料庫： {db_path} 的【{table_name}】資料表中。")


# =============================================================================
# 作業主流程
# =============================================================================
def process(wb):
    # excel_path = "缺字表.xlsx"  # 替換為你的 Excel 檔案路徑
    sheet_name = "缺字表"      # 替換為你的工作表名稱
    db_path = "Ho_Lok_Ue.db"  # 替換為你的資料庫檔案路徑
    # db_path = "QA.sqlite"  # 替換為你的資料庫檔案路徑
    table_name = "漢字庫"         # 替換為你的資料表名稱

    process_excel_to_db(wb, sheet_name, db_path, table_name)
    return EXIT_CODE_SUCCESS


# =============================================================================
# 程式主流程
# =============================================================================
def main():
    # =========================================================================
    # 開始作業
    # =========================================================================
    logging.info("作業開始")
    print(f"🔍 執行程式前，務必確認【缺字表】工作表中之【台語昔標】已填入！！")

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
