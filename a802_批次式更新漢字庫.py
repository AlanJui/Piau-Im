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
EXIT_CODE_WORKSHEET_IS_EMPTY = 4
EXIT_CODE_UNKNOWN_ERROR = 99


def upsert_han_ji_record(han_ji: str, tai_lo_im_piau: str, siong_iong_too: float):
    """
    插入或更新漢字記錄（使用 UPSERT 語法）

    若記錄不存在，則插入新記錄。
    若記錄已存在且【常用度】不同，則更新【常用度】、【摘要說明】、【更新時間】。
    若記錄已存在但【常用度】相同，則不做任何更新。

    Args:
        han_ji: 漢字
        tai_lo_im_piau: 台羅音標
        siong_iong_too: 常用度（文讀音 0.8 / 白話音 0.6）

    Returns:
        int: 影響的記錄數（0=無異動, 1=新增或更新）
    """
    try:
        cursor = db_manager.execute("""
            INSERT INTO 漢字庫 (漢字, 台羅音標, 常用度, 摘要說明, 更新時間)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(漢字, 台羅音標) DO UPDATE
            SET 常用度 = excluded.常用度,
                摘要說明 = excluded.摘要說明,
                更新時間 = excluded.更新時間
            WHERE 漢字庫.常用度 != excluded.常用度
        """, (han_ji, tai_lo_im_piau, siong_iong_too, "NA", datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        return cursor.rowcount
    except Exception as e:
        logging.error(f"資料庫操作錯誤: {e}")
        raise

# =========================================================================
# 共用函數：從 Excel 工作表更新漢字庫
# =========================================================================
def parse_cell_address(coordinates_str: str, sheet) -> list:
    """
    解析座標字串並轉換為儲存格位址清單

    Args:
        coordinates_str: 座標字串，格式如 "(5, 4); (5, 5)"
        sheet: Excel 工作表物件

    Returns:
        list: 儲存格位址清單，如 ["E5", "F5"]
    """
    cell_address_list = []
    coordinates_list = coordinates_str.split(';')

    for coord in coordinates_list:
        row_str, col_str = coord.split(',')
        row = int(row_str.strip().strip('('))
        col = int(col_str.strip().strip(')'))
        cell_address = sheet.range((row, col)).address.replace('$', '')
        cell_address_list.append(cell_address)

    return cell_address_list


def update_database_from_worksheet(wb, sheet_name: str) -> int:
    """
    從指定工作表讀取資料並更新漢字庫（通用函數）

    Args:
        wb: Excel 活頁簿物件
        sheet_name: 工作表名稱（如：缺字表、人工標音字庫、標音字庫）

    Returns:
        int: 執行狀態碼
    """
    # 1. 取得工作表
    try:
        sheet = wb.sheets[sheet_name]
    except Exception:
        print(f"⚠️ 無法找到工作表: {sheet_name}")
        return EXIT_CODE_FAILURE

    # 2. 讀取常用度設定
    gu_im_lui_hing = get_value_by_name(wb=wb, name="語音類型")
    siong_iong_too = 0.8 if gu_im_lui_hing == "文讀音" else 0.6

    # 3. 讀取資料
    data = sheet.range("A2").expand("table").value

    # 4. 檢查是否為空
    if data is None or (isinstance(data, list) and len(data) == 0):
        print(f"⚠️ 工作表 '{sheet_name}' 無資料（第 2 行以下為空）")
        return EXIT_CODE_WORKSHEET_IS_EMPTY

    # 5. 確保為 2D 列表
    if not isinstance(data[0], list):
        data = [data]

    # 6. 處理資料並更新資料庫
    try:
        with db_manager.transaction():
            for idx, row_data in enumerate(data, start=2):
                han_ji = row_data[0]  # A 欄: 漢字
                tai_gi_im_piau = row_data[1]  # B 欄: 台語音標
                coordinates = row_data[3]  # D 欄: 座標

                # 跳過無效資料
                if not han_ji or not tai_gi_im_piau or tai_gi_im_piau == "N/A":
                    continue

                # 解析儲存格位址
                cell_address_list = parse_cell_address(coordinates, sheet)

                # 轉換台語音標（TLPA）→ 台羅音標（TL）
                tl_im_piau = convert_tlpa_to_tl(tai_gi_im_piau)

                # 顯示處理訊息
                print(f"\n📌 第 {idx} 列：漢字='{han_ji}', 台語音標='{tai_gi_im_piau}', "
                      f"台羅音標='{tl_im_piau}', 儲存格={cell_address_list}")

                # 插入或更新資料庫
                rowcount = upsert_han_ji_record(
                    han_ji=han_ji,
                    tai_lo_im_piau=tl_im_piau,
                    siong_iong_too=siong_iong_too
                )

                if rowcount == 0:
                    # print(f"⇒⇨⮕  資料：【{han_ji}】、【{tl_im_piau}】、【{siong_iong_too}】已存於資料表中，未執行任何更新作業！")
                    print(f"⚠️  資料：【{han_ji} ({tl_im_piau})】、【{siong_iong_too}】已存於資料表中，未執行任何更新作業！")
                else:
                    print(f"✅ 已在資料表，新增【{han_ji}（{tl_im_piau}）】或更新【常用度：{siong_iong_too}】。")

        # 交易自動 commit
        print("\n" + "=" * 80)
        print(f"✅ 使用【{sheet_name}】工作表，更新【漢字庫】已完成！")
        return EXIT_CODE_SUCCESS

    except Exception as e:
        print(f"❌ 資料庫更新失敗: {e}")
        logging.exception(f"更新【{sheet_name}】失敗")
        return EXIT_CODE_FAILURE


# =========================================================================
# 功能 1：使用【缺字表】更新【漢字庫】資料表
# =========================================================================
def update_database_from_missing_characters(wb):
    """使用【缺字表】工作表的資料更新 SQLite 資料庫的【漢字庫】資料表"""
    return update_database_from_worksheet(wb, "缺字表")

# =========================================================================
# 功能 2：使用【人工標音字庫】更新【漢字庫】資料表
# =========================================================================
def update_database_from_jin_kang_piau_im_ji_khoo(wb):
    """使用【人工標音字庫】工作表的資料更新 SQLite 資料庫的【漢字庫】資料表"""
    return update_database_from_worksheet(wb, "人工標音字庫")


# =========================================================================
# 功能 3：使用【標音字庫】更新【漢字庫】資料表
# =========================================================================
def update_database_from_piau_im_ji_khoo(wb):
    """使用【標音字庫】工作表的資料更新 SQLite 資料庫的【漢字庫】資料表"""
    return update_database_from_worksheet(wb, "標音字庫")

# =========================================================================
# 主程式執行
# =========================================================================
def main():
    wb = xw.apps.active.books.active
    try:
        # 缺字表更新漢字庫
        exit_code = update_database_from_missing_characters(wb)
        if exit_code != EXIT_CODE_WORKSHEET_IS_EMPTY and exit_code != EXIT_CODE_SUCCESS:
            return exit_code
        # 人工標音字庫更新漢字庫
        exit_code = update_database_from_jin_kang_piau_im_ji_khoo(wb)
        if exit_code != EXIT_CODE_WORKSHEET_IS_EMPTY and exit_code != EXIT_CODE_SUCCESS:
            return exit_code
        # 標音字庫更新漢字庫
        exit_code = update_database_from_piau_im_ji_khoo(wb)
        if exit_code != EXIT_CODE_WORKSHEET_IS_EMPTY and exit_code != EXIT_CODE_SUCCESS:
            return exit_code
    except Exception as e:
        logging.exception("主程式執行錯誤")
        return EXIT_CODE_UNKNOWN_ERROR

if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)