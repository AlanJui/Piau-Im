# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import logging
import os

# import re
# import sqlite3
import sys
from datetime import datetime
from pathlib import Path

# 載入第三方套件
import xlwings as xw
from dotenv import load_dotenv

# 載入自訂模組/函式
from a320_人工標音更正漢字自動標音 import jin_kang_piau_im_cu_han_ji_piau_im
from mod_excel_access import (
    convert_to_excel_address,
    excel_address_to_row_col,
    get_active_cell,
    get_active_cell_info,
    get_row_col_from_coordinate,
    get_value_by_name,
)
from mod_字庫 import JiKhooDict  # 漢字字庫物件
from mod_標音 import PiauIm

# from mod_標音 import convert_tl_with_tiau_hu_to_tlpa  # 去除台語音標的聲調符號
# from mod_標音 import is_punctuation  # 是否為標點符號
# from mod_標音 import split_hong_im_hu_ho  # 分解漢字標音
# from mod_標音 import tlpa_tng_han_ji_piau_im  # 漢字標音物件

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
EXIT_CODE_FAILURE = 1  # 失敗
EXIT_CODE_NO_FILE = 1  # 無法找到檔案
EXIT_CODE_INVALID_INPUT = 2  # 輸入錯誤
EXIT_CODE_PROCESS_FAILURE = 3  # 過程失敗
EXIT_CODE_UNKNOWN_ERROR = 99  # 未知錯誤

# =========================================================================
# 作業程序
# =========================================================================
def check_han_ji_in_excel(wb, han_ji, excel_cell):
    """
    在【標音字庫】工作表內查詢【漢字】與【Excel座標】是否存在。

    :param wb: Excel 活頁簿物件
    :param han_ji: 要查找的漢字 (str)
    :param excel_cell: 要查找的 Excel 座標 (如 "D9")
    :return: Boolean 值 (True: 找到, False: 未找到)
    """
    sheet_name = "標音字庫"  # Excel 工作表名稱
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

    for row in data:
        row_han_ji = row[0]  # A 欄: 漢字
        coordinates = row[4]  # E 欄: 座標 (可能是 "(9, 4); (25, 9)" 這類格式)

        if row_han_ji == han_ji and coordinates:
            # 將座標解析成一個 set
            coord_list = coordinates.split("; ")
            parsed_coords = {convert_to_excel_address(coord) for coord in coord_list}

            # 檢查 Excel 座標是否在列表內
            if excel_cell in parsed_coords:
                return True

    return False


def check_and_update_pronunciation(wb, han_ji, position, jin_kang_piau_im):
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

    # 建置 PiauIm 物件，供作漢字拼音轉換作業
    han_ji_khoo_field = '漢字庫'
    han_ji_khoo_name = get_value_by_name(wb=wb, name=han_ji_khoo_field)
    piau_im = PiauIm(han_ji_khoo=han_ji_khoo_name)           # 指定漢字自動查找使用的【漢字庫】
    piau_im_huat = get_value_by_name(wb=wb, name='標音方法')   # 指定【台語音標】轉換成【漢字標音】的方法

    # 建置自動及人工漢字標音字庫工作表：（1）【標音字庫】（2）【人工標音字】
    piau_im_sheet_name = '標音字庫'
    piau_im_ji_khoo = JiKhooDict.create_ji_khoo_dict_from_sheet(
        wb=wb,
        sheet_name=piau_im_sheet_name)

    jin_kang_piau_im_sheet_name='人工標音字庫'
    jin_kang_piau_im_ji_khoo = JiKhooDict.create_ji_khoo_dict_from_sheet(
        wb=wb,
        sheet_name=jin_kang_piau_im_sheet_name)

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
        tai_gi_im_piau = row[1]  # B 欄: 台語音標
        kenn_ziann_im_piau = row[2]  # C 欄: 校正音標
        coordinates = row[3]  # D 欄: 座標 (可能是 "(9, 4); (25, 9)" 這類格式)
        correction_pronounce_cell = sheet.range(f"D{idx+2}")  # D 欄: 校正音標

        row, col = get_row_col_from_coordinate(coordinates)  # 取得座標的行列
        cell = sheet.range((row, col))  # 取得該儲存格物件

        if row_han_ji == han_ji and coordinates:
            # 將座標解析成一個 set
            coord_list = coordinates.split("; ")
            parsed_coords = {convert_to_excel_address(coord) for coord in coord_list}

            # 確認該座標是否存在於【標音字庫】中
            # if convert_to_excel_address(str(position)) in parsed_coords:
            position_address = convert_to_excel_address(str(position))
            if position_address in parsed_coords:
                # 檢查【漢字】標注之【人工標音】是否與【台語音標】不同
                if jin_kang_piau_im != tai_gi_im_piau:
                    tai_gi_im_piau, han_ji_piau_im = jin_kang_piau_im_cu_han_ji_piau_im(
                        wb=wb,
                        jin_kang_piau_im=jin_kang_piau_im,
                        piau_im=piau_im,
                        piau_im_huat=piau_im_huat)

                    # 【標音字庫】添加或更新【漢字】及【台語音標】資料
                    jin_kang_piau_im_ji_khoo.add_entry(
                        han_ji=han_ji,
                        tai_gi_im_piau=tai_gi_im_piau,
                        kenn_ziann_im_piau=jin_kang_piau_im,
                        coordinates=(row, col)
                    )
                    # ----- 新增程式邏輯：更新【標音字庫】 -----
                    # Step 1: 在【標音字庫】搜尋該筆【漢字】+【台語音標】
                    existing_entries = piau_im_ji_khoo.ji_khoo_dict.get(han_ji, [])

                    # 標記是否找到
                    entry_found = False

                    for existing_entry in existing_entries:
                        # Step 2: 若找到，移除該筆資料內的座標
                        if (row, col) in existing_entry["coordinates"]:
                            existing_entry["coordinates"].remove((row, col))
                        entry_found = True
                        break  # 找到即可離開迴圈

                    # Step 3: 將此筆資料（校正音標為 'N/A'）於【標音字庫】底端新增
                    piau_im_ji_khoo.add_entry(
                        han_ji=han_ji,
                        tai_gi_im_piau=tai_gi_im_piau,
                        kenn_ziann_im_piau="N/A",  # 預設值
                        coordinates=(row, col)
                    )

                    # 將文字顏色設為【紅色】
                    cell.font.color = (255, 0, 0)
                    # 將儲存格的填滿色彩設為【黄色】
                    cell.color = (255, 255, 0)

                    # 更新【校正音標】為【人工標音】
                    # correction_pronounce_cell.value = jin_kang_piau_im
                    correction_pronounce_cell.value = tai_gi_im_piau
                    print(f"✅ {position}【{han_ji}】： 台語音標 {tai_gi_im_piau} -> 校正標音 {jin_kang_piau_im}")
                    return True

        #----------------------------------------------------------------------
        # 作業結束前處理
        #----------------------------------------------------------------------
        # 將【標音字庫】、【人工標音字庫】，寫入 Excel 工作表
        piau_im_ji_khoo.write_to_excel_sheet(wb=wb, sheet_name=piau_im_sheet_name)
        jin_kang_piau_im_ji_khoo.write_to_excel_sheet(wb=wb, sheet_name=jin_kang_piau_im_sheet_name)

        logging_process_step("作用中【漢字】儲存格之【人工標音】已更新至【標音字庫】。")
        return EXIT_CODE_SUCCESS

    print(f"❌ 未找到匹配的資料或不符合更新條件: {han_ji} ({position})")
    return False



def ut01(wb):
    han_ji = "傀"  # 要查找的漢字
    excel_cell = "D9"  # 要查找的 Excel 座標

    exists = check_han_ji_in_excel(wb, han_ji, excel_cell)
    if exists:
        print(f"✅ 漢字 '{han_ji}' 存在於 {excel_cell}")
    else:
        print(f"❌ 找不到漢字 '{han_ji}' 在 {excel_cell}")

    return EXIT_CODE_SUCCESS


def ut02(wb):
    # 作業流程：獲取當前作用中的 Excel 儲存格
    sheet_name, cell_address = get_active_cell(wb)
    print(f"✅ 目前作用中的儲存格：{sheet_name} 工作表 -> {cell_address}")

    # 將 Excel 儲存格地址轉換為 (row, col) 格式
    row, col = excel_address_to_row_col(cell_address)
    print(f"📌 Excel 位址 {cell_address} 轉換為 (row, col): ({row}, {col})")

    return EXIT_CODE_SUCCESS


# =============================================================================
# 作業主流程
# =============================================================================
def process(wb):
    """
    作業流程：
    1. 取得當前 Excel 作用儲存格 (漢字、座標)
    2. 計算【人工標音】位置與值
    3. 查詢【標音字庫】確認該座標是否已登錄
    4. 若【標正音標】為 'N/A'，則更新為【人工標音】
    """
    # 取得當前 Excel 作用儲存格資訊
    sheet_name, han_ji, active_cell, artificial_pronounce, position = get_active_cell_info(wb)

    print(f"📌 作用儲存格：{active_cell}，位於【{sheet_name}】工作表")
    print(f"📌 漢字：{han_ji}，漢字儲存格座標：{active_cell}")
    print(f"📌 人工標音：{artificial_pronounce}，人工標音儲存格座標：{position}")

    # 執行檢查與更新
    success = check_and_update_pronunciation(wb, han_ji, active_cell, artificial_pronounce)

    return EXIT_CODE_SUCCESS if success else EXIT_CODE_FAILURE


# =============================================================================
# 程式主流程
# =============================================================================
def main():
    # =========================================================================
    # 開始作業
    # =========================================================================
    logging.info("作業開始")

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
            logging.info("a330_以作用儲存格之人工標音更標音字庫.py 程式已執行完畢！")

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
