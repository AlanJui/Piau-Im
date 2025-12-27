# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import logging
import os
import sqlite3
import sys
from pathlib import Path
from typing import Callable

# 載入第三方套件
import xlwings as xw
from dotenv import load_dotenv

# 載入自訂模組/函式
from a310_缺字表修正後續作業 import process as update_khuat_ji_piau_by_jin_kang_piau_im
from a320_人工標音更正漢字自動標音 import process as update_by_jin_kang_piau_im
from a330_以標音字庫更新漢字注音工作表 import process as update_by_piau_im_ji_khoo
from mod_excel_access import (
    convert_to_excel_address,
    ensure_sheet_exists,
    get_row_col_from_coordinate,
    get_value_by_name,
    strip_cell,
)
from mod_file_access import save_as_new_file
from mod_字庫 import JiKhooDict  # 漢字字庫物件
from mod_帶調符音標 import (
    cing_bo_iong_ji_bu,
    is_han_ji,
    kam_si_u_tiau_hu,
    tng_im_piau,
    tng_tiau_ho,
)
from mod_標音 import PiauIm  # 漢字標音物件
from mod_標音 import convert_tl_with_tiau_hu_to_tlpa  # 去除台語音標的聲調符號
from mod_標音 import is_punctuation  # 是否為標點符號
from mod_標音 import split_hong_im_hu_ho  # 分解漢字標音
from mod_標音 import tlpa_tng_han_ji_piau_im  # 台語音標轉台語音標

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
# 本程式主要處理作業程序
# =========================================================================
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


def process(wb):
    """
    更新【漢字注音】表中【台語音標】儲存格的內容，依據【標音字庫】中的【校正音標】欄位進行更新，並將【校正音標】覆蓋至原【台語音標】。
    """
    logging_process_step("<=========== 開始處理流程作業！==========>")
    try:
        # 取得工作表
        target_sheet_name = '漢字注音'
        ensure_sheet_exists(wb, target_sheet_name)
        han_ji_piau_im_sheet = wb.sheets['漢字注音']
        han_ji_piau_im_sheet.activate()
    except Exception as e:
        logging_exc_error(msg=f"找不到【漢字注音】工作表 ！", error=e)
        return EXIT_CODE_PROCESS_FAILURE
    logging_process_step(f"已完成作業所需之初始化設定！")

    #-------------------------------------------------------------------------
    # 將【缺字表】工作表，已填入【台語音標】之資料，登錄至【標音字庫】工作表
    # 使用【缺字表】工作表中的【校正音標】，更正【漢字注音】工作表中之【台語音標】、【漢字標音】；
    # 並依【缺字表】工作表中的【台語音標】儲存格內容，更新【標音字庫】工作表中之【台語音標】及【校正音標】欄位
    #-------------------------------------------------------------------------
    try:
        sheet_name = '缺字表'
        logging_process_step(f"以【{sheet_name}】工作表之【校正音標】，更新【{target_sheet_name}】工作表之【台語音標】與【漢字標音】！")
        print('\n\n')
        print("=" * 100)
        print(f"使用【{sheet_name}】工作表的【校正音標】欄位，更新【{target_sheet_name}】工作表之【台語音標】、【漢字標音】：")
        print("=" * 100)
        # 將【缺字表】工作表中的【台語音標】儲存格內容，更新至【標音字庫】工作表中之【台語音標】及【校正音標】欄位
        # update_khuat_ji_piau(wb=wb)
        # 依據【缺字表】工作表紀錄，並參考【漢字注音】工作表在【人工標音】欄位的內容，更新【缺字表】工作表中的【校正音標】及【台語音標】欄位
        # 即使用者為【漢字】補入查找不到的【台語音標】時，若是在【缺字表】工作表中之【校正音標】直接填寫
        # 則應執行 a310*.py 程式；但使用者若是在【漢字注音】工作表中之【人工標音】欄位填寫，則應執行 a320*.py 程式
        # a300*.py 之本程式
        update_khuat_ji_piau_by_jin_kang_piau_im(wb=wb, sheet_name=sheet_name)
    except Exception as e:
        logging_exc_error(msg=f"處理【缺字表】作業異常！", error=e)
        return EXIT_CODE_PROCESS_FAILURE
    #-------------------------------------------------------------------------
    # 將【漢字注音】工作表，【漢字】填入【人工標音】內容，登錄至【人工標音字庫】及【標音字庫】工作表
    #-------------------------------------------------------------------------
    try:
        sheet_name = '人工標音字庫'
        logging_process_step(f"以【{sheet_name}】工作表之【校正音標】，更新【{target_sheet_name}】工作表之【台語音標】與【漢字標音】！")
        print('\n\n')
        print("=" * 100)
        print(f"使用【{sheet_name}】工作表的【校正音標】欄位，更新【{target_sheet_name}】工作表之【台語音標】、【漢字標音】：")
        print("=" * 100)
        update_by_jin_kang_piau_im(wb=wb, sheet_name=sheet_name )
    except Exception as e:
        logging_exc_error(msg=f"處理【漢字】之【人工標音】作業異常！", error=e)
        return EXIT_CODE_PROCESS_FAILURE
    #-------------------------------------------------------------------------
    # 根據【標音字庫】工作表，更新【漢字注音】工作表中的【台語音標】及【漢字標音】欄位
    #-------------------------------------------------------------------------
    try:
        sheet_name = '標音字庫'
        logging_process_step(f"以【{sheet_name}】工作表之【校正音標】，更新【{target_sheet_name}】工作表之【台語音標】與【漢字標音】！")
        print('\n\n')
        print("=" * 100)
        print(f"使用【{sheet_name}】工作表中的【校正音標】，更新【漢字注音】工作表中的【台語音標】：")
        print("=" * 100)
        update_by_piau_im_ji_khoo(wb, sheet_name=sheet_name)
    except Exception as e:
        logging_exc_error(msg=f"處理以【標音字庫】更新【漢字注音】工作表之作業，發生執行異常！", error=e)
        return EXIT_CODE_PROCESS_FAILURE
    #--------------------------------------------------------------------------
    # 結束作業
    #--------------------------------------------------------------------------
    han_ji_piau_im_sheet.activate()
    logging_process_step("<=========== 完成處理流程作業！==========>")

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
        if result_code != EXIT_CODE_SUCCESS:
            msg = f"程式異常終止：{program_name}"
            logging_exc_error(msg=msg, error=e)
            return EXIT_CODE_PROCESS_FAILURE
    except Exception as e:
        msg = f"程式異常終止：{program_name}"
        logging_exc_error(msg=msg, error=e)
        return EXIT_CODE_UNKNOWN_ERROR

    finally:
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

        # if wb:
        #     xw.apps.active.quit()  # 確保 Excel 被釋放資源，避免開啟殘留

    # =========================================================================
    # 結束程式
    # =========================================================================
    logging_process_step(f"《========== 程式終止執行：{program_name} ==========》")
    return EXIT_CODE_SUCCESS    # 作業正常結束


if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)

