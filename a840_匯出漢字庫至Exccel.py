# =========================================================================
# a840_匯出漢字庫至Exccel.py
#
# 功能說明：
# 將【Ho_Lok_Ue.db】/【漢字庫】資料表內的所有資料紀錄匯出至
# 作用中 Excel 活頁簿的【漢字庫】工作表，作為備份／檢視之用。
# 與 a850_使用Excel重建漢字庫.py 為【備份／還原】配對工具。
#
# 【漢字庫】工作表欄位結構：
#   A 識別號、B 漢字、C 台羅音標、D 常用度、
#   E 摘要說明、F 更新時間、G 最近揀用時間
#
# 用法：
#   1. 開啟（或新建）Excel 活頁簿，並使之處於作用中；
#   2. 將 Excel 活頁簿，以：「河洛話漢字庫.xlsx」為檔名儲存。
#   3. 執行指令：python a840_匯出漢字庫至Exccel.py
#   4. 再次執行存檔，將 Excel 活頁簿檔，新增之【漢字庫】工作表存檔。
# =========================================================================

# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import logging
import os
import sqlite3
import sys

import xlwings as xw
from dotenv import load_dotenv

from mod_excel_access import ensure_sheet_exists

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

HEADERS = ["識別號", "漢字", "台羅音標", "常用度", "摘要說明", "更新時間", "最近揀用時間"]


# =========================================================================
# 功能：將資料庫之【漢字庫】資料表，備份至 Excel 工作表
# =========================================================================
def export_database_to_excel(wb, sheet_name="漢字庫"):
    """
    將 `漢字庫` 資料表的資料寫入 Excel 的【漢字庫】工作表。

    :param wb: Excel 活頁簿物件
    :param sheet_name: 工作表名稱，預設為 "漢字庫"
    :return: EXIT_CODE_SUCCESS or EXIT_CODE_FAILURE
    """
    try:
        ensure_sheet_exists(wb, sheet_name)
        sheet = wb.sheets[sheet_name]
    except Exception:
        print(f"⚠️ 無法找到工作表: {sheet_name}")
        return EXIT_CODE_FAILURE

    conn = sqlite3.connect(DB_HO_LOK_UE)
    cursor = conn.cursor()

    try:
        # 【漢字庫】資料表現行結構：識別號、漢字、台羅音標、常用度、摘要說明、更新時間、最近揀用時間
        cursor.execute(
            """
            SELECT 識別號, 漢字, 台羅音標, 常用度, 摘要說明, 更新時間, 最近揀用時間
            FROM 漢字庫
            ORDER BY 識別號 ASC;
            """
        )
        rows = cursor.fetchall()

        # 清空舊內容
        sheet.clear()

        # 下列欄位先設為【文字】格式，避免 Excel 自動型別轉換破壞資料：
        # - C欄（台羅音標）：如 'jun7'、'jun2' 會被誤判為日期（6月7日），變成 '2026-06-07 00:00:00'
        # - F、G欄（更新時間、最近揀用時間）：'YYYY-MM-DD HH:MM:SS' 字串會被轉成【日期】值
        sheet.range("C:C").number_format = "@"
        sheet.range("F:G").number_format = "@"

        # 寫入標題列與資料
        sheet.range("A1").value = HEADERS
        if rows:
            sheet.range("A2").value = rows

        print(f"✅ 資料成功匯出至 Excel！（共 {len(rows)} 筆）")
        logging.info("已匯出漢字庫至 Excel：%s 筆", len(rows))
        return EXIT_CODE_SUCCESS

    except Exception as e:
        print(f"❌ 匯出資料失敗: {e}")
        logging.error("匯出漢字庫至 Excel 失敗: %s", e)
        return EXIT_CODE_FAILURE

    finally:
        conn.close()


# =========================================================================
# 主程式執行
# =========================================================================
def main():
    sheet_name = sys.argv[1] if len(sys.argv) > 1 else "漢字庫"

    try:
        wb = xw.apps.active.books.active
    except Exception as e:
        print(f"❌ 無法取得作用中的 Excel 活頁簿: {e}")
        return EXIT_CODE_FAILURE

    if not wb:
        print("❌ 無法作業，因未有任何 Excel 檔案已開啟。")
        return EXIT_CODE_FAILURE

    result = export_database_to_excel(wb, sheet_name)
    if result == EXIT_CODE_SUCCESS:
        try:
            wb.save()
            print(f"✅ 已儲存活頁簿：{wb.fullname or wb.name}")
        except Exception as e:
            print(f"⚠️ 資料已寫入工作表，但儲存失敗: {e}")
    return result


if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)
