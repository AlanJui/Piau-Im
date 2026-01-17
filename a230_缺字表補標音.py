# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import logging
import os
import re
import sys
from pathlib import Path

import xlwings as xw
from dotenv import load_dotenv

# 載入自訂模組/函式
from mod_logging import init_logging, logging_exc_error, logging_process_step
from mod_帶調符音標 import tng_im_piau, tng_tiau_ho
from mod_標音 import PiauIm, tlpa_tng_han_ji_piau_im

# 載入自訂模組
from mod_程式 import ExcelCell, Program

# =========================================================================
# 載入環境變數
# =========================================================================
load_dotenv()

# 預設檔案名稱從環境變數讀取
DB_HO_LOK_UE = os.getenv('DB_HO_LOK_UE', 'Ho_Lok_Ue.db')
DB_KONG_UN = os.getenv('DB_KONG_UN', 'Kong_Un.db')

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
# 設定日誌
# =========================================================================
init_logging()

# =========================================================================
# 程式區域函式
# =========================================================================
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
        sheet = wb.sheets["缺字表"]
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
        han_ji = sheet.range(f"A{row}").value  # 讀取 A 欄（漢字）
        if not han_ji:  # 若 A 欄為空，則結束迴圈
            break

        # 取得原始【台語音標】並轉換為 TLPA+ 格式
        im_piau = sheet.range(f"C{row}").value

        tlpa_im_piau = tng_im_piau(im_piau)   # 將【音標】使用之【拼音字母】轉換成【TLPA拼音字母】；【音標調符】仍保持
        tai_gi_im_piau = tng_tiau_ho(tlpa_im_piau).lower()  # 將【音標調符】轉換成【數值調號】
        # tai_lo_im_piau = convert_tl_with_tiau_hu_to_tlpa(im_piau)

        # 更新【缺字表】中【校正音標】欄（D 欄）
        sheet.range(f"D{row}").value = tai_gi_im_piau

        print(f"{row-1}. (A{row}) 【{han_ji}】： 取自【校正音標】欄，可能使用【調符】之【台語音標】：{im_piau}, 正規化之【台語音標】：{tai_gi_im_piau}")

        # 讀取【缺字表】中【座標】欄（E 欄）的內容，該內容可能含有多組座標，如 "(5, 17); (33, 8); (77, 5)"
        coordinates_str = sheet.range(f"E{row}").value
        if coordinates_str:
            # 利用正規表達式解析所有形如 (row, col) 的座標
            coordinate_tuples = re.findall(r"\((\d+)\s*,\s*(\d+)\)", coordinates_str)
            for tup in coordinate_tuples:
                try:
                    r_coord = int(tup[0])
                    c_coord = int(tup[1])
                except ValueError:
                    continue  # 若轉換失敗，跳過該組座標

                # 根據說明，【台語音標】應填入漢字儲存格上方一列 (row - 1)，相同欄位
                target_row = r_coord - 1
                tai_gi_im_piau_cell = (target_row, c_coord)

                # 將【校正音標】填入【漢字注音】工作表漢字之【台語音標】儲存格
                han_ji_piau_im_sheet.range(tai_gi_im_piau_cell).value = tai_gi_im_piau
                print(f"更新『台語音標』：座標 {tai_gi_im_piau_cell} 填入音標：{tai_gi_im_piau}")

                #--------------------------------------------------------------------------
                # 【漢字標音】
                #--------------------------------------------------------------------------
                # 使用【台語音標】轉換，取得【漢字標音】
                han_ji_im_piau = tlpa_tng_han_ji_piau_im(
                    piau_im=piau_im, piau_im_huat=piau_im_huat, tai_gi_im_piau=tai_gi_im_piau
                )
                # 根據說明，【漢字標音】應填入漢字儲存格下方一列 (row + 1)，相同欄位
                target_row = r_coord + 1
                han_ji_piau_im_cell = (target_row, c_coord)

                # 將【校正音標】填入【漢字注音】工作表漢字之【台語音標】儲存格
                han_ji_piau_im_sheet.range(han_ji_piau_im_cell).value = han_ji_im_piau
                print(f"更新『漢字注音』：座標 {han_ji_piau_im_cell} 填入音標：{han_ji_im_piau}")

        row += 1  # 讀取下一列

    return EXIT_CODE_SUCCESS


# =========================================================================
# 主程式執行
# =========================================================================


def process(wb, args) -> int:
    """
    查詢漢字讀音並標注

    Args:
        wb: Excel Workbook 物件
        args: 命令列參數

    Returns:
        處理結果代碼
    """
    #--------------------------------------------------------------------------
    # 作業初始化
    #--------------------------------------------------------------------------
    logging_process_step("<=========== 作業開始！==========>")

    try:
        #--------------------------------------------------------------------------
        # 初始化 process config
        #--------------------------------------------------------------------------
        program = Program(wb=wb, args=args, hanji_piau_im_sheet='漢字注音')

        # 建立儲存格處理器
        xls_cell = ExcelCell(
            program=program,
            new_jin_kang_piau_im_ji_khoo_sheet=False,
            new_piau_im_ji_khoo_sheet=False,
            new_khuat_ji_piau_sheet=False,
        )
    except Exception as e:
        logging_exc_error(msg="處理作業異常！", error=e)
        return EXIT_CODE_PROCESS_FAILURE

    #--------------------------------------------------------------------------
    # 作業處理中
    #--------------------------------------------------------------------------
    try:
        # 處理工作表
        sheet_name = program.hanji_piau_im_sheet
        sheet = wb.sheets[sheet_name]
        sheet.activate()

        # 更新【缺字表】中的【台語音標】至【漢字注音】工作表
        update_khuat_ji_piau(wb)

        # 寫回字庫到 Excel
        xls_cell.save_all_piau_im_ji_khoo_dicts()

    except Exception as e:
        logging_exc_error(msg="處理作業異常！", error=e)
        return EXIT_CODE_PROCESS_FAILURE

    #--------------------------------------------------------------------------
    # 處理作業結束
    #--------------------------------------------------------------------------
    print('\n')
    logging_process_step("<=========== 作業結束！==========>")
    return EXIT_CODE_SUCCESS


def main(args):
    # =========================================================================
    # 程式初始化
    # =========================================================================
    # 取得專案根目錄。
    current_file_path = Path(__file__).resolve()
    project_root = current_file_path.parent
    # 取得程式名稱
    # program_file_name = current_file_path.name
    program_name = current_file_path.stem

    # =========================================================================
    # 開始執行程式
    # =========================================================================
    logging_process_step(f"《========== 程式開始執行：{program_name} ==========》")
    logging_process_step(f"專案根目錄為: {project_root}")

    try:
        # 取得 Excel 活頁簿
        wb = None
        wb = xw.apps.active.books.active
    except Exception as e:
        logging.error(f"無法找到作用中的 Excel 工作簿: {e}")
        return EXIT_CODE_NO_FILE

    #===========================================================================
    # 執行處理作業
    #===========================================================================
    try:
        exit_code = process(wb=wb, args=args)

        # 儲存檔案
        if exit_code == EXIT_CODE_SUCCESS:
            try:
                wb.save()
                file_path = wb.fullname
                logging_process_step(f"儲存檔案至路徑：{file_path}")
            except Exception as e:
                logging_exc_error(msg="儲存檔案異常！", error=e)
                return EXIT_CODE_SAVE_FAILURE
    except Exception:
        logging.exception("處理作業失敗！")
        return EXIT_CODE_PROCESS_FAILURE


    #===========================================================================
    # 終止程式
    #===========================================================================
    # 顯示程式結束訊息
    logging_process_step(f"《========== 程式終止執行：{program_name}: {exit_code} ==========》")
    return EXIT_CODE_SUCCESS


if __name__ == "__main__":
    args = sys.argv[1:]  # 獲取命令列參數
    exit_code = main(args)
    if exit_code != EXIT_CODE_SUCCESS:
        print(f"程式異常終止，返回代碼：{exit_code}")
        sys.exit(exit_code)