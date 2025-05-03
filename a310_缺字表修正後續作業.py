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
from mod_excel_access import get_value_by_name, save_as_new_file
from mod_帶調符音標 import tng_im_piau, tng_tiau_ho
from mod_標音 import PiauIm  # 漢字標音物件
from mod_標音 import tlpa_tng_han_ji_piau_im  # 台語音標轉台語音標
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
DB_HO_LOK_UE = os.getenv('DB_HO_LOK_UE', 'Ho_Lok_Ue.db')
DB_KONG_UN = os.getenv('DB_KONG_UN', 'Kong_Un.db')

# =========================================================================
# 設定日誌
# =========================================================================
from mod_logging import (
    init_logging,
    logging_exc_error,
    logging_exception,
    logging_process_step,
    logging_warning,
)

init_logging()

# =========================================================================
# 程式區域函式
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
    table_name = "漢字庫"         # 替換為你的資料表名稱
    hue_im = wb.names['語音類型'].refers_to_range.value
    siong_iong_too = 0.8 if hue_im == "文讀音" else 0.6  # 根據語音類型設定常用度

    # 讀取資料表範圍
    data = sheet.range("A2").expand("table").value

    # # 確保資料為 2D 列表
    # if not isinstance(data[0], list):
    #     data = [data]
    # 若資料為空（即表格沒有任何資料），直接跳出處理

    # 若完全無資料或只有空列，視為異常處理
    if not data or data == [[]]:
        raise ValueError("【缺字表】工作表內，無任何資料，略過後續處理作業。")

    # 若只有一列資料（如一筆記錄），資料可能不是 2D list，要包成 list
    if not isinstance(data[0], list):
        data = [data]

    idx = 0
    for row in data:
        han_ji = row[0] # 漢字
        tai_gi_im_piau = row[1] # 台語音標
        hau_ziann_im_piau = row[2] # 台語音標
        zo_piau = row[3] # (儲存格位置)座標

        if han_ji and (tai_gi_im_piau != 'N/A' or hau_ziann_im_piau != 'N/A'):
            # 將 Excel 工作表存放的【台語音標（TLPA）】，改成資料庫保存的【台羅拼音（TL）】
            tlpa_im_piau = tng_im_piau(tai_gi_im_piau)   # 將【音標】使用之【拼音字母】轉換成【TLPA拼音字母】；【音標調符】仍保持
            tlpa_im_piau_cleanned = tng_tiau_ho(tlpa_im_piau).lower()  # 將【音標調符】轉換成【數值調號】
            tl_im_piau = convert_tlpa_to_tl(tlpa_im_piau_cleanned)

            insert_or_update_to_db(db_path, table_name, han_ji, tl_im_piau, piau_im_huat)
            print(f"📌 {idx+1}. 【{han_ji}】==> {zo_piau}：台羅音標：【{tl_im_piau}】、校正音標：【{hau_ziann_im_piau}】、台語音標=【{tai_gi_im_piau}】、座標：{zo_piau}")
            idx += 1

    logging_process_step(f"【缺字表】中的資料已成功回填至資料庫： {db_path} 的【{table_name}】資料表中。")
    return EXIT_CODE_SUCCESS


def update_khuat_ji_piau(wb):
    """
    讀取 Excel 檔案，依據【缺字表】工作表中的資料執行下列作業：
      1. 由 A 欄讀取漢字，從 C 欄取得原始【台語音標】，並轉換為 TLPA+ 格式後更新 D 欄（校正音標）。
      2. 從 E 欄讀取座標字串（可能含有多組座標），每組座標指向【漢字注音】工作表中該漢字儲存格，
         而【台語音標】應填入位於該漢字儲存格上方一列（row - 1）的相同欄位。
         若該儲存格尚無值，則填入校正音標。
    """
    # 取得本函式所需之【選項】參數
    try:
        han_ji_khoo = wb.names["漢字庫"].refers_to_range.value
        piau_im_huat = wb.names["標音方法"].refers_to_range.value
    except Exception as e:
        logging_exc_error("找不到作業所需之選項設定", e)
        return EXIT_CODE_INVALID_INPUT

    piau_im = PiauIm(han_ji_khoo=han_ji_khoo)

    # 取得【缺字表】工作表
    try:
        khuat_ji_piau_sheet = wb.sheets["缺字表"]
    except Exception as e:
        logging_exc_error("找不到名為『缺字表』的工作表", e)
        return EXIT_CODE_INVALID_INPUT

    # 取得【漢字注音】工作表
    try:
        han_ji_piau_im_sheet = wb.sheets["漢字注音"]
    except Exception as e:
        logging_exc_error("找不到名為『漢字注音』的工作表", e)
        return EXIT_CODE_INVALID_INPUT

    row = 2  # 從第 2 列開始（跳過標題列）
    while True:
        han_ji = khuat_ji_piau_sheet.range(f"A{row}").value  # 讀取 A 欄（漢字）
        if not han_ji:  # 若 A 欄為空，則結束迴圈
            break

        # 更新【缺字表】中【校正音標】欄（C 欄）
        hau_ziann_im_piau = khuat_ji_piau_sheet.range(f"C{row}").value
        if hau_ziann_im_piau == "N/A" or not hau_ziann_im_piau:  # 若 C 欄為空，則結束迴圈
            row += 1
            continue

        tlpa_im_piau = tng_im_piau(hau_ziann_im_piau)   # 將【音標】使用之【拼音字母】轉換成【TLPA拼音字母】；【音標調符】仍保持
        tai_gi_im_piau = tng_tiau_ho(tlpa_im_piau).lower()  # 將【音標調符】轉換成【數值調號】
        # 取得原始【台語音標】並轉換為 TLPA+ 格式
        im_piau = khuat_ji_piau_sheet.range(f"B{row}").value
        khuat_ji_piau_sheet.range(f"B{row}").value = tai_gi_im_piau  # 更新 C 欄（校正音標）

        coordinates_str = khuat_ji_piau_sheet.range(f"D{row}").value
        print(f"{row-1}. (A{row}) 【{han_ji}】==> {coordinates_str} ： 原音標：{im_piau}, 校正音標：{tai_gi_im_piau}")

        # 讀取【缺字表】中【座標】欄（E 欄）的內容，該內容可能含有多組座標，如 "(5, 17); (33, 8); (77, 5)"
        if coordinates_str:
            # 利用正規表達式解析所有形如 (row, col) 的座標
            coordinate_tuples = re.findall(r"\((\d+)\s*,\s*(\d+)\)", coordinates_str)
            for tup in coordinate_tuples:
                try:
                    r_coord = int(tup[0])
                    c_coord = int(tup[1])
                except ValueError:
                    continue  # 若轉換失敗，跳過該組座標

                han_ji_cell = (r_coord, c_coord)  # 漢字儲存格座標

                # 根據說明，【台語音標】應填入漢字儲存格上方一列 (row - 1)，相同欄位
                target_row = r_coord - 1
                tai_gi_im_piau_cell = (target_row, c_coord)

                # 將【校正音標】填入【漢字注音】工作表漢字之【台語音標】儲存格
                han_ji_piau_im_sheet.range(tai_gi_im_piau_cell).value = tai_gi_im_piau
                excel_address = han_ji_piau_im_sheet.range(tai_gi_im_piau_cell).address
                excel_address = excel_address.replace("$", "")  # 去除 "$" 符號
                print(f"   台語音標：【{tai_gi_im_piau}】，填入座標：{excel_address} = {tai_gi_im_piau_cell}")

                #--------------------------------------------------------------------------
                # 【漢字標音】
                #--------------------------------------------------------------------------
                # 使用【台語音標】轉換，取得【漢字標音】
                han_ji_piau_im = tlpa_tng_han_ji_piau_im(
                    piau_im=piau_im, piau_im_huat=piau_im_huat, tai_gi_im_piau=tai_gi_im_piau
                )
                # 根據說明，【漢字標音】應填入漢字儲存格下方一列 (row + 1)，相同欄位
                target_row = r_coord + 1
                han_ji_piau_im_cell = (target_row, c_coord)

                # 將【校正音標】填入【漢字注音】工作表漢字之【台語音標】儲存格
                han_ji_piau_im_sheet.range(han_ji_piau_im_cell).value = han_ji_piau_im
                excel_address = han_ji_piau_im_sheet.range(tai_gi_im_piau_cell).address
                excel_address = excel_address.replace("$", "")  # 去除 "$" 符號
                print(f"   漢字標音：【{han_ji_piau_im}】，填入座標：{excel_address} = {han_ji_piau_im_cell}")
                # 將【漢字注音】工作表之【漢字】儲存格之底色，重置為【無底色】
                han_ji_piau_im_sheet.range(han_ji_cell).color = None

        row += 1  # 讀取下一列

    return EXIT_CODE_SUCCESS


# =========================================================================
# 本程式主要處理作業程序
# =========================================================================
def process(wb):
    """
    更新【漢字注音】表中【台語音標】儲存格的內容，依據【標音字庫】中的【校正音標】欄位進行更新，並將【校正音標】覆蓋至原【台語音標】。
    """
    logging_process_step("<----------- 作業開始！---------->")
    try:
        # 取得工作表
        han_ji_piau_im_sheet = wb.sheets['漢字注音']
        han_ji_piau_im_sheet.activate()
    except Exception as e:
        logging_exc_error(msg=f"找不到【漢字注音】工作表 ！", error=e)
        return EXIT_CODE_PROCESS_FAILURE
    logging_process_step(f"已完成作業所需之初始化設定！")

    #-------------------------------------------------------------------------
    # 【缺字表】工作表，原先找不到【音標】之漢字，已補填【台語音標】之後續處理作業
    #-------------------------------------------------------------------------
    try:
        wb.sheets['缺字表'].activate()
        update_khuat_ji_piau(wb)
    except Exception as e:
        logging_exc_error(msg=f"處理【缺字表】作業異常！", error=e)
        return EXIT_CODE_PROCESS_FAILURE
    logging_process_step(f"完成：處理【缺字表】作業")

    #-------------------------------------------------------------------------
    # 將【缺字表】之【漢字】與【台語音標】存入【漢字庫】作業
    #-------------------------------------------------------------------------
    try:
        wb.sheets['缺字表'].activate()
        khuat_ji_piau_poo_im_piau(wb)
    except Exception as e:
        logging_exc_error(
            msg=f"將【缺字表】之【漢字】與【台語音標】存入【漢字庫】作業，發生執行異常！",
            error=e)
        return EXIT_CODE_PROCESS_FAILURE
    logging_process_step(f"完成：將【缺字表】之【漢字】與【台語音標】存入【漢字庫】作業")

    #--------------------------------------------------------------------------
    # 結束作業
    #--------------------------------------------------------------------------
    logging_process_step("<----------- 作業結束！---------->")

    return EXIT_CODE_SUCCESS

# =========================================================================
# 程式主要作業流程
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
    # (1) 開始執行程式
    # =========================================================================
    logging_process_step(f"《========== 程式開始執行：{program_name} ==========》")
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
    except Exception as e:
        msg = f"程式異常終止：{program_name}"
        logging_exc_error(msg=msg, error=e)
        return EXIT_CODE_UNKNOWN_ERROR

    if result_code != EXIT_CODE_SUCCESS:
        msg = f"程式異常終止：{program_name}（非例外，而是返回失敗碼）"
        logging.error(msg)
        return EXIT_CODE_PROCESS_FAILURE

    #--------------------------------------------------------------------------
    # 儲存檔案
    #--------------------------------------------------------------------------
    try:
        # 要求畫面回到【漢字注音】工作表
        wb.sheets['漢字注音'].activate()
        # 儲存檔案
        file_path = save_as_new_file(wb=wb)
        if not file_path:
            logging_exc_error(msg="儲存檔案失敗！", error=e)
            return EXIT_CODE_SAVE_FAILURE    # 作業異當終止：無法儲存檔案
        else:
            logging_process_step(f"儲存檔案至路徑：{file_path}")
    except Exception as e:
        logging_exc_error(msg="儲存檔案失敗！", error=e)
        return EXIT_CODE_SAVE_FAILURE    # 作業異當終止：無法儲存檔案

    # =========================================================================
    # 結束程式
    # =========================================================================
    logging_process_step(f"《========== 程式終止執行：{program_name} ==========》")
    return EXIT_CODE_SUCCESS    # 作業正常結束


if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)

