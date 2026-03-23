"""a501_標音用工作表回填字典資料庫.py v0.2.0

功能說明：
    1. 讀取 Excel 的【缺字表/人工標音字庫/標音字庫】工作表，並將資料回填至 SQLite 資料庫。

    2. 預設工作表為：人工標音字庫。

使用說明：

        1. 執行此程式前，請先確保 Excel 已開啟包含【缺字表/人工標音字庫/標音字庫】工作表的活頁簿檔案。
        2. 執行此程式，並在命令列參數中指定要回填的工作表名稱 (預設: 人工標音字庫，可選: 標音字庫, 缺字表)。
           例如：python a501_標音用工作表回填字典資料庫.py --sheet 人工標音字庫

變更紀錄：
    v0.2.0 (2024-06-30)：變更原先只支援【缺字表】工作表回填，改為支援【缺字表/人工標音字庫/標音字庫】三個工作表回填；同時新增命令列參數以指定要回填的工作表名稱。
"""

# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import argparse
import logging
import os
import sqlite3
from datetime import datetime
from pathlib import Path

# 載入第三方套件
import xlwings as xw
from dotenv import load_dotenv

# 載入自訂模組/函式
from mod_file_access import save_as_new_file
from mod_帶調符音標 import tng_im_piau, tng_tiau_ho
from mod_標音 import convert_tlpa_to_tl

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
from mod_logging import (
    init_logging,
    logging_exc_error,
    logging_exception,
    logging_process_step,
)

init_logging()


# =========================================================================
# 程式區域函式
# =========================================================================
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

    # 檢查是否已存在該漢字與音標的組合
    cursor.execute(
        f"SELECT 識別號 FROM {table_name} WHERE 漢字 = ? AND 台羅音標 = ?",
        (han_ji, tai_gi_im_piau),
    )
    row = cursor.fetchone()

    siong_iong_too = 0.8 if piau_im_huat == "文讀音" else 0.6
    if row:
        # 更新資料 (如果已經存在相同的漢字和音標，只需更新時間)
        cursor.execute(
            f"""
        UPDATE {table_name}
        SET 更新時間 = ?
        WHERE 識別號 = ?;
        """,
            (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), row[0]),
        )
    else:
        # 檢查是否已存在該漢字 (但音標不同)
        cursor.execute(f"SELECT 識別號 FROM {table_name} WHERE 漢字 = ?", (han_ji,))
        row_han_ji = cursor.fetchone()

        if row_han_ji:
            # 如果漢字存在但音標不同，我們應該新增一筆紀錄，因為一個漢字可以有多個讀音
            # 但因為有 UNIQUE constraint (漢字, 台羅音標)，所以只要音標不同就可以新增
            cursor.execute(
                f"""
            INSERT INTO {table_name} (漢字, 台羅音標, 常用度, 摘要說明)
            VALUES (?, ?, ?, NULL);
            """,
                (han_ji, tai_gi_im_piau, siong_iong_too),
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
# 使用【缺字表】更新【漢字庫】資料庫
# =========================================================================
def khuat_ji_piau_poo_im_piau(wb, sheet_name: str = "人工標音字庫"):
    """
    讀取 Excel 的【缺字表/人工標音字庫/標音字庫】工作表，並將資料回填至 SQLite 資料庫。

    :param excel_path: Excel 檔案路徑。
    :param sheet_name: Excel 工作表名稱。
    :param db_path: 資料庫檔案路徑。
    :param table_name: 資料表名稱。
    """
    # sheet_name = "缺字表"
    # sheet_name = "人工標音字庫"
    sheet = wb.sheets[sheet_name]
    piau_im_huat = wb.names["語音類型"].refers_to_range.value
    db_path = "Ho_Lok_Ue.db"  # 替換為你的資料庫檔案路徑
    table_name = "漢字庫"  # 替換為你的資料表名稱
    # hue_im = wb.names["語音類型"].refers_to_range.value
    # siong_iong_too = 0.8 if hue_im == "文讀音" else 0.6  # 根據語音類型設定常用度

    # 讀取資料表範圍
    # data = sheet.range("A2").expand("table").value

    # 從 A2 開始讀取，並嘗試讀取到 D 欄
    try:
        last_row = sheet.range("A" + str(sheet.cells.last_cell.row)).end("up").row
        if last_row < 2:
            print("Excel 無資料 (至少需要有一列資料)。")
            return

        # 讀取所有資料（ A2:F{last_row} ）
        data = sheet.range(f"A2:D{last_row}").value
    except Exception as e:
        print(f"讀取 Excel 資料失敗: {e}")
        return

    # 確保資料為 2D 列表
    if not isinstance(data[0], list):
        data = [data]

    idx = 0
    for row in data:
        han_ji = row[0]  # 漢字
        tai_gi_im_piau = row[1]  # 台語音標
        hau_ziann_im_piau = row[2]  # 校正音標
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
        f"【{sheet_name}】中的資料已成功回填至資料庫： {db_path} 的【{table_name}】資料表中。"
    )
    return EXIT_CODE_SUCCESS


# =========================================================================
# 使用【缺字表】更新【漢字庫】資料庫
# =========================================================================
def update_han_ji_khoo_by_piau_im_ji_khoo(wb, sheet_name: str = "標音字庫"):
    """
    讀取 Excel 的【缺字表/人工標音字庫/標音字庫】工作表，並將資料回填至 SQLite 資料庫。

    :param excel_path: Excel 檔案路徑。
    :param sheet_name: Excel 工作表名稱。
    :param db_path: 資料庫檔案路徑。
    :param table_name: 資料表名稱。
    """
    # sheet_name = "缺字表"
    # sheet_name = "人工標音字庫"
    sheet = wb.sheets[sheet_name]
    piau_im_huat = wb.names["語音類型"].refers_to_range.value
    db_path = "Ho_Lok_Ue.db"  # 替換為你的資料庫檔案路徑
    table_name = "漢字庫"  # 替換為你的資料表名稱
    hue_im = wb.names["語音類型"].refers_to_range.value
    siong_iong_too = 0.8 if hue_im == "文讀音" else 0.6  # 根據語音類型設定常用度

    # 讀取資料表範圍
    # data = sheet.range("A2").expand("table").value

    # 從 A2 開始讀取，並嘗試讀取到 D 欄
    try:
        last_row = sheet.range("A" + str(sheet.cells.last_cell.row)).end("up").row
        if last_row < 2:
            print("Excel 無資料 (至少需要有一列資料)。")
            return

        # 讀取所有資料（ A2:F{last_row} ）
        data = sheet.range(f"A2:D{last_row}").value
    except Exception as e:
        print(f"讀取 Excel 資料失敗: {e}")
        return

    # 確保資料為 2D 列表
    if not isinstance(data[0], list):
        data = [data]

    idx = 0
    for row in data:
        han_ji = row[0]  # 漢字
        tai_gi_im_piau = row[1]  # 台語音標
        hau_ziann_im_piau = row[2]  # 校正音標
        zo_piau = row[3]  # (儲存格位置)座標

        if han_ji and (not tai_gi_im_piau and tai_gi_im_piau != "N/A"):
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
        f"【{sheet_name}】中的資料已成功回填至資料庫： {db_path} 的【{table_name}】資料表中。"
    )
    return EXIT_CODE_SUCCESS


# =============================================================================
# 作業主流程
# =============================================================================
def process(wb, sheet_name: str = "人工標音字庫"):
    logging_process_step("<----------- 作業開始！---------->")

    try:
        khuat_ji_piau_poo_im_piau(wb, sheet_name)
    except Exception as e:
        logging_exc_error(msg=f"無法將【{sheet_name}】資料回填至資料庫！", error=e)
        return EXIT_CODE_PROCESS_FAILURE

    # ---------------------------------------------------------------------
    # 作業結尾處理
    # ---------------------------------------------------------------------
    # 要求畫面回到指定的工作表
    try:
        wb.sheets[sheet_name].activate()
    except Exception:
        pass  # 如果工作表不存在，忽略錯誤
    # 作業正常結束
    logging_process_step("<----------- 作業結束！---------->")
    return EXIT_CODE_SUCCESS


# =============================================================================
# 程式主流程
# =============================================================================
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

    # 解析命令列參數
    parser = argparse.ArgumentParser(description="將標音用工作表回填至字典資料庫")
    parser.add_argument(
        "-s",
        "--sheet",
        type=str,
        default="人工標音字庫",
        help="指定要回填的工作表名稱 (預設: 人工標音字庫，可選: 標音字庫, 缺字表)",
    )
    args = parser.parse_args()
    sheet_name = args.sheet

    # =========================================================================
    # (2) 設定【作用中活頁簿】：偵測及獲取 Excel 已開啟之活頁簿檔案。
    # =========================================================================
    wb = None
    # 取得【作用中活頁簿】
    try:
        wb = xw.apps.active.books.active  # 取得 Excel 作用中的活頁簿檔案
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
        status_code = process(wb, sheet_name)
        if status_code != EXIT_CODE_SUCCESS:
            msg = f"程式異常終止：{program_name}"
            logging_exc_error(msg=msg, error=None)
            return EXIT_CODE_PROCESS_FAILURE

    except Exception as e:
        msg = f"程式異常終止：{program_name}"
        logging_exc_error(msg=msg, error=e)
        return EXIT_CODE_UNKNOWN_ERROR

    finally:
        # --------------------------------------------------------------------------
        # 儲存檔案
        # --------------------------------------------------------------------------
        try:
            # 要求畫面回到【漢字注音】工作表
            wb.sheets["漢字注音"].activate()
            # 儲存檔案
            file_path = save_as_new_file(wb=wb)
            if not file_path:
                logging_exc_error(msg="儲存檔案失敗！", error=e)
                return EXIT_CODE_SAVE_FAILURE  # 作業異當終止：無法儲存檔案
            else:
                logging_process_step(f"儲存檔案至路徑：{file_path}")
        except Exception as e:
            logging_exc_error(msg="儲存檔案失敗！", error=e)
            return EXIT_CODE_SAVE_FAILURE  # 作業異當終止：無法儲存檔案

        # if wb:
        #     xw.apps.active.quit()  # 確保 Excel 被釋放資源，避免開啟殘留

    # =========================================================================
    # 結束程式
    # =========================================================================
    logging_process_step(f"《========== 程式終止執行：{program_name} ==========》")
    return EXIT_CODE_SUCCESS  # 作業正常結束


if __name__ == "__main__":
    exit_code = main()
