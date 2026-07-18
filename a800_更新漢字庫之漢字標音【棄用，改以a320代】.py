"""a800_更新漢字庫之漢字標音.py v0.2.7
使用【工作表】的漢字標音資料，更新 SQLite 資料庫的【漢字庫】資料表。
- 將【台語音標】轉換為【台羅音標】後寫入資料庫。
- 【網頁匯入】工作表只用【台語音標】。

功能：
 1. 使用【缺字表】
 2. 使用【人工標音字庫】
 3. 使用【標音字庫】
 4. 使用【網頁匯入】

v0.2.7 2026-02-23: 加入功能 4
"""

# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import logging
import sys
from datetime import datetime

import xlwings as xw

from mod_database import db_manager
from mod_excel_access import get_value_by_name
from mod_標音 import convert_tlpa_to_tl

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


def upsert_han_ji_record(han_ji: str, tai_lo_im_piau: str, siong_iong_too: float):
    """
    插入或更新漢字記錄（使用 UPSERT 語法）

    Args:
        han_ji: 漢字
        tai_lo_im_piau: 台羅音標
        siong_iong_too: 常用度（文讀音 0.8 / 白話音 0.6）

    Returns:
        int: 影響的記錄數
    """
    try:
        cursor = db_manager.execute(
            """
            INSERT INTO 漢字庫 (漢字, 台羅音標, 常用度, 摘要說明, 更新時間)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(漢字, 台羅音標) DO UPDATE
            SET 常用度 = excluded.常用度,
                摘要說明 = excluded.摘要說明,
                更新時間 = excluded.更新時間
        """,
            (
                han_ji,
                tai_lo_im_piau,
                siong_iong_too,
                "NA",
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            ),
        )
        return cursor.rowcount
    except Exception as e:
        logging.error(f"資料庫操作錯誤: {e}")
        raise


# =========================================================================
# 功能 1：使用【缺字表】更新【漢字庫】資料表
# =========================================================================
def update_database_from_missing_characters(wb):
    """
    使用【缺字表】工作表的資料更新 SQLite 資料庫的【漢字庫】資料表。
    """
    sheet_name = "缺字表"
    try:
        sheet = wb.sheets[sheet_name]
    except Exception:
        print(f"⚠️ 無法找到工作表: {sheet_name}")
        return EXIT_CODE_FAILURE

    # 讀取【語音類型】以便設定【常用度】
    gu_im_lui_hing = get_value_by_name(wb=wb, name="語音類型")
    # 確定 `常用度`（文讀音 0.8 / 白話音 0.6）
    siong_iong_too = 0.8 if gu_im_lui_hing == "文讀音" else 0.6

    # 讀取資料範圍
    data = sheet.range("A2").expand("table").value  # 讀取所有資料

    # 確保資料為 2D 列表
    if not isinstance(data[0], list):
        data = [data]

    try:
        # 使用交易管理
        with db_manager.transaction():
            for idx, row_data in enumerate(
                data, start=2
            ):  # Excel A2 起始，Python Index 2
                han_ji = row_data[0]  # A 欄: 漢字
                tai_gi_im_piau = row_data[1]  # B 欄: 台語音標
                # tai_lo_im_piau = row_data[2]  # C 欄: 校正音標
                coordinates = row_data[3]  # D 欄: 座標
                cell_address_list = []
                # 將【座標】欄位的字串轉換為【座標】串列
                coordinates_list = coordinates.split(";")
                # 將【座標】欄位的字串轉換為元組
                for coordinates in coordinates_list:
                    row = col = cell_address = None
                    row_str, col_str = coordinates.split(",")
                    row_str = row_str.strip()
                    row = int(row_str.strip("("))
                    col = int(col_str.strip(")"))
                    # 轉換(row, col) 為 Excel 儲存格位址
                    # 使用 xlwings Range 物件來取得儲存格位址
                    cell_address = sheet.range((row, col)).address
                    cell_address = cell_address.replace("$", "")  # 移除 $ 符號
                    # print(f"📍 位置: {cell_address}")
                    # 加入【儲存格位址】清單
                    cell_address_list.append(cell_address)
                    # print(f"📍 位置: {cell_address_list}")

                if not han_ji or not tai_gi_im_piau or tai_gi_im_piau == "N/A":
                    continue  # 跳過無效資料

                # **轉換台語音標（TLPA）→ 台羅音標（TL）**
                tl_im_piau = convert_tlpa_to_tl(tai_gi_im_piau)

                # **在 INSERT 之前，顯示 Console 訊息**
                print(
                    f"\n📌 第 {idx} 列：漢字='{han_ji}', 台語音標='{tai_gi_im_piau}', 台羅音標='{tl_im_piau}', 儲存格={cell_address_list}"
                )

                # **插入或更新資料庫（使用 UPSERT）**
                rowcount = upsert_han_ji_record(
                    han_ji=han_ji,
                    tai_lo_im_piau=tl_im_piau,
                    siong_iong_too=siong_iong_too,
                )

                if rowcount == 0:
                    print(f"⚠️ 第 {idx} 列資料更新失敗！")
                else:
                    print(f"✅ 第 {idx} 列資料已更新至資料庫。")

        # 交易自動 commit
        print("\n" + "=" * 80)
        print("✅ 資料庫更新完成！")
        return EXIT_CODE_SUCCESS

    except Exception as e:
        print(f"❌ 資料庫更新失敗: {e}")
        logging.exception("資料庫更新失敗")
        return EXIT_CODE_FAILURE


# =========================================================================
# 功能 2：使用【人工標音字庫】更新【漢字庫】資料表
# =========================================================================
def update_database_from_jin_kang_piau_im_ji_khoo(wb):
    """
    使用【人工標音字庫】工作表的資料更新 SQLite 資料庫的【漢字庫】資料表。
    - 將【台語音標】轉換為【台羅音標】後寫入資料庫。

    :param wb: Excel 活頁簿物件
    :return: EXIT_CODE_SUCCESS or EXIT_CODE_FAILURE
    """
    sheet_name = "人工標音字庫"
    try:
        sheet = wb.sheets[sheet_name]
    except Exception:
        print(f"⚠️ 無法找到工作表: {sheet_name}")
        return EXIT_CODE_FAILURE

    # 讀取【語音類型】以便設定【常用度】
    gu_im_lui_hing = get_value_by_name(wb=wb, name="語音類型")
    # 確定 `常用度`（文讀音 0.8 / 白話音 0.6）
    siong_iong_too = 0.8 if gu_im_lui_hing == "文讀音" else 0.6

    # 讀取資料範圍
    data = sheet.range("A2").expand("table").value  # 讀取所有資料

    # 確保資料為 2D 列表
    if not isinstance(data[0], list):
        data = [data]

    try:
        # 使用交易管理
        with db_manager.transaction():
            for idx, row_data in enumerate(
                data, start=2
            ):  # Excel A2 起始，Python Index 2
                han_ji = row_data[0]  # A 欄: 漢字
                tai_gi_im_piau = row_data[1]  # B 欄: 台語音標
                # tai_lo_im_piau = row_data[2]  # C 欄: 校正音標
                coordinates = row_data[3]  # D 欄: 座標
                cell_address_list = []
                # 將【座標】欄位的字串轉換為【座標】串列
                coordinates_list = coordinates.split(";")
                # 將【座標】欄位的字串轉換為元組
                for coordinates in coordinates_list:
                    row = col = cell_address = None
                    row_str, col_str = coordinates.split(",")
                    row_str = row_str.strip()
                    row = int(row_str.strip("("))
                    col = int(col_str.strip(")"))
                    # 轉換(row, col) 為 Excel 儲存格位址
                    # 使用 xlwings Range 物件來取得儲存格位址
                    cell_address = sheet.range((row, col)).address
                    cell_address = cell_address.replace("$", "")  # 移除 $ 符號
                    # print(f"📍 位置: {cell_address}")
                    # 加入【儲存格位址】清單
                    cell_address_list.append(cell_address)
                    # print(f"📍 位置: {cell_address_list}")

                if not han_ji or not tai_gi_im_piau or tai_gi_im_piau == "N/A":
                    continue  # 跳過無效資料

                # **轉換台語音標（TLPA）→ 台羅音標（TL）**
                tl_im_piau = convert_tlpa_to_tl(tai_gi_im_piau)

                # **在 INSERT 之前，顯示 Console 訊息**
                print(
                    f"\n📌 第 {idx} 列：漢字='{han_ji}', 台語音標='{tai_gi_im_piau}', 台羅音標='{tl_im_piau}', 儲存格={cell_address_list}"
                )

                # **插入或更新資料庫（使用 UPSERT）**
                rowcount = upsert_han_ji_record(
                    han_ji=han_ji,
                    tai_lo_im_piau=tl_im_piau,
                    siong_iong_too=siong_iong_too,
                )

                if rowcount == 0:
                    print(f"⚠️ 第 {idx} 列資料更新失敗！")
                else:
                    print(f"✅ 第 {idx} 列資料已更新至資料庫。")

        # 交易自動 commit
        print("\n" + "=" * 80)
        print("✅ 資料庫更新完成！")
        return EXIT_CODE_SUCCESS

    except Exception as e:
        print(f"❌ 資料庫更新失敗: {e}")
        logging.exception("資料庫更新失敗")
        return EXIT_CODE_FAILURE


# =========================================================================
# 功能 3：使用【標音字庫】更新【漢字庫】資料表
# =========================================================================
def update_database_from_piau_im_ji_khoo(wb):
    """
    使用【標音字庫】工作表的資料更新 SQLite 資料庫的【漢字庫】資料表。
    """
    sheet_name = "標音字庫"
    try:
        sheet = wb.sheets[sheet_name]
    except Exception:
        print(f"⚠️ 無法找到工作表: {sheet_name}")
        return EXIT_CODE_FAILURE

    # 讀取【語音類型】以便設定【常用度】
    gu_im_lui_hing = get_value_by_name(wb=wb, name="語音類型")
    # 確定 `常用度`（文讀音 0.8 / 白話音 0.6）
    siong_iong_too = 0.8 if gu_im_lui_hing == "文讀音" else 0.6

    # 讀取資料範圍
    data = sheet.range("A2").expand("table").value  # 讀取所有資料

    # 確保資料為 2D 列表
    if not isinstance(data[0], list):
        data = [data]

    try:
        # 使用交易管理
        with db_manager.transaction():
            for idx, row_data in enumerate(
                data, start=2
            ):  # Excel A2 起始，Python Index 2
                han_ji = row_data[0]  # A 欄: 漢字
                tai_gi_im_piau = row_data[1]  # B 欄: 台語音標
                # tai_lo_im_piau = row_data[2]  # C 欄: 校正音標
                coordinates = row_data[3]  # D 欄: 座標
                cell_address_list = []
                # 將【座標】欄位的字串轉換為【座標】串列
                coordinates_list = coordinates.split(";")
                # 將【座標】欄位的字串轉換為元組
                for coordinates in coordinates_list:
                    row = col = cell_address = None
                    row_str, col_str = coordinates.split(",")
                    row_str = row_str.strip()
                    row = int(row_str.strip("("))
                    col = int(col_str.strip(")"))
                    # 轉換(row, col) 為 Excel 儲存格位址
                    # 使用 xlwings Range 物件來取得儲存格位址
                    cell_address = sheet.range((row, col)).address
                    cell_address = cell_address.replace("$", "")  # 移除 $ 符號
                    # print(f"📍 位置: {cell_address}")
                    # 加入【儲存格位址】清單
                    cell_address_list.append(cell_address)
                    # print(f"📍 位置: {cell_address_list}")

                if not han_ji or not tai_gi_im_piau or tai_gi_im_piau == "N/A":
                    continue  # 跳過無效資料

                # **轉換台語音標（TLPA）→ 台羅音標（TL）**
                tl_im_piau = convert_tlpa_to_tl(tai_gi_im_piau)

                # **在 INSERT 之前，顯示 Console 訊息**
                print(
                    f"\n📌 第 {idx} 列：漢字='{han_ji}', 台語音標='{tai_gi_im_piau}', 台羅音標='{tl_im_piau}', 儲存格={cell_address_list}"
                )

                # **插入或更新資料庫（使用 UPSERT）**
                rowcount = upsert_han_ji_record(
                    han_ji=han_ji,
                    tai_lo_im_piau=tl_im_piau,
                    siong_iong_too=siong_iong_too,
                )

                if rowcount == 0:
                    print(f"⚠️ 第 {idx} 列資料更新失敗！")
                else:
                    print(f"✅ 第 {idx} 列資料已更新至資料庫。")

        # 交易自動 commit
        print("\n" + "=" * 80)
        print("✅ 資料庫更新完成！")
        return EXIT_CODE_SUCCESS

    except Exception as e:
        print(f"❌ 資料庫更新失敗: {e}")
        logging.exception("資料庫更新失敗")
        return EXIT_CODE_FAILURE


# =========================================================================
# 功能 4：使用【網頁匯入】更新【漢字庫】資料表
# =========================================================================
def update_database_from_web_page_import_work_sheet(wb):
    """
    依【網頁匯入】工作表的資料，更新 SQLite 資料庫的【漢字庫】資料表。
    - 此工作表只用【台語音標】。

    :param wb: Excel 活頁簿物件
    :return: EXIT_CODE_SUCCESS or EXIT_CODE_FAILURE
    """
    """
    使用【缺字表】工作表的資料更新 SQLite 資料庫的【漢字庫】資料表。
    """
    sheet_name = "網頁匯入"
    try:
        sheet = wb.sheets[sheet_name]
    except Exception:
        print(f"⚠️ 無法找到工作表: {sheet_name}")
        return EXIT_CODE_FAILURE

    # 讀取【語音類型】以便設定【常用度】
    gu_im_lui_hing = get_value_by_name(wb=wb, name="語音類型")
    # 確定 `常用度`（文讀音 0.8 / 白話音 0.6）
    siong_iong_too = 0.8 if gu_im_lui_hing == "文讀音" else 0.6

    # 讀取資料範圍
    # data = sheet.range("A2").expand("table").value  # 讀取所有資料
    # =========================================================================
    # 修正 Bug: 原來的 expand("table") 遇到 B 欄為空時，會導致 C, D, E, F 欄未被讀入。
    # 改用明確指定範圍 (Range) 的方式讀取。
    # =========================================================================
    last_row = sheet.range("A" + str(sheet.cells.last_cell.row)).end("up").row

    # 檢查是否有資料
    if last_row < 2:
        print("⚠️ 無資料可讀取！(A2 以下為空)")
        return EXIT_CODE_SUCCESS

    # 明確讀取 A2:F{last_row}
    data = sheet.range(f"A2:F{last_row}").value

    # 確保資料為 2D 列表 (若只有一列資料，xlwings 可能回傳 1D list)
    if isinstance(data, list) and len(data) > 0 and not isinstance(data[0], list):
        data = [data]
    elif data is None:
        data = []

    try:
        # 使用交易管理
        with db_manager.transaction():
            # Excel A2 起始，Python Index 2
            for idx, row_data in enumerate(data, start=2):
                # 自工作表之各欄讀取資料
                han_ji = row_data[0]  # A 欄: 漢字
                # han_ji_piau_im = row_data[1]  # B 欄: 漢字標音（應為十五音切語，如 "堅五曾"）
                tai_gi_im_piau = row_data[2]  # C 欄: 台語音標

                # 確保讀入之【漢字】不是：【標點符號】、【換行】、【空白】或【N/A】
                if not han_ji or not tai_gi_im_piau or tai_gi_im_piau == "N/A":
                    continue  # 跳過無效資料

                # 合併聲母、韻母、調號為【台語音標（TLPA）】
                siann = row_data[3]  # D 欄: 聲母
                un = row_data[4]  # E 欄: 韻母
                tiau = int(row_data[5])  # F 欄: 調號
                tlpa_im_piau = f"{siann}{un}{tiau}"
                # **轉換台語音標（TLPA）→ 台羅音標（TL）**
                tl_im_piau = convert_tlpa_to_tl(tlpa_im_piau)

                # **在 INSERT 之前，顯示 Console 訊息**
                print(
                    f"\n📌 第 {idx} 列：漢字='{han_ji}', 台語音標='{tai_gi_im_piau}', 台羅音標='{tl_im_piau}'"
                )

                # **插入或更新資料庫（使用 UPSERT）**
                rowcount = upsert_han_ji_record(
                    han_ji=han_ji,
                    tai_lo_im_piau=tl_im_piau,
                    siong_iong_too=siong_iong_too,
                )

                if rowcount == 0:
                    print(f"⚠️ 資料更新失敗！")
                else:
                    print(f"✅ 已更新至資料庫。")

        # 交易自動 commit
        print("\n" + "=" * 80)
        print("✅ 資料庫更新完成！")
        return EXIT_CODE_SUCCESS

    except Exception as e:
        print(f"❌ 資料庫更新失敗: {e}")
        logging.exception("資料庫更新失敗")
        return EXIT_CODE_FAILURE


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
        # 缺字表更新漢字庫
        return update_database_from_missing_characters(wb)
    elif mode == "2":
        # 人工標音字庫更新漢字庫
        return update_database_from_jin_kang_piau_im_ji_khoo(wb)
    elif mode == "3":
        # 標音字庫更新漢字庫
        return update_database_from_piau_im_ji_khoo(wb)
    elif mode == "4":
        # 【網頁匯入】更新漢字庫
        return update_database_from_web_page_import_work_sheet(wb)
    else:
        print("❌ 錯誤：請輸入有效模式 (1)：缺字表、(2)人工標音字庫、(3)標音字庫")
        return EXIT_CODE_INVALID_INPUT


if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)
