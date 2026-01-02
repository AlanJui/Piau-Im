# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import logging
import os
import sqlite3
import sys
from datetime import datetime

import xlwings as xw
from dotenv import load_dotenv

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
# 功能：從 Excel 工作表匯入資料到【漢字庫】資料表（添加模式）
# =========================================================================
def import_data_from_excel(wb, sheet_name="漢字庫"):
    """
    從 Excel 工作表中讀取資料，並以【添加】模式匯入 SQLite 資料庫的【漢字庫】資料表。
    - 檢查工作表是否存在。
    - 讀取資料並驗證。
    - 若資料重複，顯示警示訊息但不中斷程式。
    - 將新資料添加到【漢字庫】資料表。

    :param wb: Excel 活頁簿物件
    :param sheet_name: 工作表名稱，預設為 "漢字庫"
    :return: EXIT_CODE_SUCCESS or EXIT_CODE_FAILURE
    """
    try:
        # 檢查工作表是否存在
        if sheet_name not in [sheet.name for sheet in wb.sheets]:
            print(f"⚠️ 無法找到工作表: {sheet_name}")
            return EXIT_CODE_FAILURE

        # 取得工作表
        sheet = wb.sheets[sheet_name]

        # 讀取資料範圍（從 A2 開始）
        data = sheet.range("A2").expand("table").value

        # 確保資料為 2D 列表
        if not isinstance(data[0], list):
            data = [data]

        # 連接到 SQLite 資料庫
        conn = sqlite3.connect(DB_HO_LOK_UE)
        cursor = conn.cursor()

        # 檢查【漢字庫】資料表是否存在，若不存在則建立
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS 漢字庫 (
            識別號 INTEGER PRIMARY KEY AUTOINCREMENT,
            漢字 TEXT NOT NULL,
            台羅音標 TEXT NOT NULL,
            常用度 REAL DEFAULT 0.1,
            摘要說明 TEXT DEFAULT 'NA',
            更新時間 TEXT DEFAULT (DATETIME('now', 'localtime')) NOT NULL,
            UNIQUE (漢字, 台羅音標)
        );
        """)

        # 插入資料
        for idx, row_data in enumerate(data, start=2):  # Excel A2 起始，Python Index 2
            han_ji = row_data[1]  # B 欄: 漢字
            tai_lo_im_piau = row_data[2]  # C 欄: 台羅音標
            siong_iong_too = row_data[3] if isinstance(row_data[2], (int, float)) else 0.1  # D 欄: 常用度
            summary = row_data[4] if isinstance(row_data[4], str) else "NA"  # E 欄: 摘要說明
            updated_time = row_data[5] if isinstance(row_data[5], str) else datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            # 檢查漢字與台羅音標是否為有效資料
            if not han_ji or not tai_lo_im_piau:
                print(f"⚠️ 跳過無效資料: Excel 第 {idx} 列：缺【漢字】或【台羅音標】")
                continue

            # 處理台羅音標中的斜線分隔格式（例如: go5/gia5）
            tai_lo_list = []
            if '/' in str(tai_lo_im_piau):
                # 將斜線分隔的音標拆分成多個
                tai_lo_list = [piau.strip() for piau in str(tai_lo_im_piau).split('/') if piau.strip()]
                print(f"ℹ️ Excel 第 {idx} 列：偵測到多音標格式 '{tai_lo_im_piau}'，拆分為 {len(tai_lo_list)} 個音標")
            else:
                tai_lo_list = [tai_lo_im_piau]

            # 對每個台羅音標進行處理
            for tai_lo in tai_lo_list:
                # 檢查資料是否重複
                cursor.execute("""
                    SELECT 1 FROM 漢字庫 WHERE 漢字 = ? AND 台羅音標 = ?;
                """, (han_ji, tai_lo))
                if cursor.fetchone():
                    print(f"⚠️ 資料重複: Excel 第 {idx} 列 (漢字='{han_ji}', 台羅音標='{tai_lo}')")
                    continue

                # 插入資料到【漢字庫】資料表
                try:
                    cursor.execute("""
                        INSERT INTO 漢字庫 (漢字, 台羅音標, 常用度, 摘要說明, 更新時間)
                        VALUES (?, ?, ?, ?, ?);
                    """, (han_ji, tai_lo, siong_iong_too, summary, updated_time))
                    print(f"✅ 已新增: 漢字='{han_ji}', 台羅音標='{tai_lo}'")
                except sqlite3.IntegrityError as e:
                    print(f"⚠️ 資料重複或插入失敗: Excel 第 {idx} 列 (錯誤: {e})")

        # 提交變更
        conn.commit()
        print(f"✅ 資料已成功從工作表 '{sheet_name}' 添加到【漢字庫】資料表！")
        return EXIT_CODE_SUCCESS

    except Exception as e:
        print(f"❌ 匯入資料失敗: {e}")
        return EXIT_CODE_FAILURE

    finally:
        # 關閉資料庫連接
        if conn:
            conn.close()

# =========================================================================
# 主程式執行
# =========================================================================
def main():
    # 檢查是否有指定工作表名稱
    if len(sys.argv) > 1:
        sheet_name = sys.argv[1]
    else:
        # sheet_name = "漢字庫"  # 預設工作表名稱
        # sheet_name = "甲骨釋文漢字庫"  # 預設工作表名稱
        sheet_name = "台語字庫"  # 預設工作表名稱

    # 取得當前作用中的 Excel 活頁簿
    try:
        wb = xw.books.active
    except Exception as e:
        print(f"❌ 無法取得作用中的 Excel 活頁簿: {e}")
        return EXIT_CODE_FAILURE

    # 呼叫匯入資料函式
    return import_data_from_excel(wb, sheet_name)

if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)