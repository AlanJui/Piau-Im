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
from mod_ca_ji_tian import HanJiTian
from mod_excel_access import (
    convert_to_excel_address,
    delete_sheet_by_name,
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
# 資料類別：儲存處理配置
# =========================================================================
class ProcessConfig:
    """處理配置資料類別"""

    def __init__(self, wb, args, hanji_piau_im_sheet: str = '漢字注音'):
        self.wb = wb
        self.args = args
        # 【漢字注音】工作表描述
        self.hanji_piau_im_sheet = hanji_piau_im_sheet
        self.TOTAL_LINES = int(wb.names['每頁總列數'].refers_to_range.value)
        self.ROWS_PER_LINE = 4
        self.line_start_row = 3  # 第一行【標音儲存格】所在 Excel 列號: 3
        self.line_end_row = self.line_start_row + (self.TOTAL_LINES * self.ROWS_PER_LINE)
        self.CHARS_PER_ROW = int(wb.names['每列總字數'].refers_to_range.value)
        self.start_col = 4
        self.end_col = self.start_col + self.CHARS_PER_ROW
        self.han_ji_orgin_cell = 'V3'  # 原始漢字儲存格位置
        # 每一行【漢字標音行】組成結構
        self.jin_kang_piau_im_row_offset = 0    # 人工標音儲存格
        self.tai_gi_im_piau_row_offset = 1      # 台語音標儲存格
        self.han_ji_row_offset = 2              # 漢字儲存格
        self.han_ji_piau_im_row_offset = 3      # 漢字標音儲存格
        # 漢字起始列號
        self.han_ji_start_row = self.line_start_row + self.han_ji_row_offset
        # 初始化字典物件
        self.han_ji_khoo_name = wb.names['漢字庫'].refers_to_range.value
        self.db_name = DB_HO_LOK_UE if self.han_ji_khoo_name == '河洛話' else DB_KONG_UN
        self.ji_tian = HanJiTian(self.db_name)
        self.piau_im = PiauIm(han_ji_khoo=self.han_ji_khoo_name)
        # 標音相關
        self.piau_im_huat = wb.names['標音方法'].refers_to_range.value
        self.ue_im_lui_piat = wb.names['語音類型'].refers_to_range.value    # 文讀音或白話音


class CellProcessor:
    """儲存格處理器"""

    def __init__(
        self,
        config: ProcessConfig,
        jin_kang_piau_im_ji_khoo: JiKhooDict,
        piau_im_ji_khoo: JiKhooDict,
        khuat_ji_piau_ji_khoo: JiKhooDict,
    ):
        self.config = config
        self.ji_tian = config.ji_tian
        self.piau_im = config.piau_im
        self.piau_im_huat = config.piau_im_huat
        self.ue_im_lui_piat = config.ue_im_lui_piat
        self.han_ji_khoo = config.han_ji_khoo_name
        self.jin_kang_piau_im_ji_khoo = jin_kang_piau_im_ji_khoo
        self.piau_im_ji_khoo = piau_im_ji_khoo
        self.khuat_ji_piau_ji_khoo = khuat_ji_piau_ji_khoo


# =========================================================================
# 作業處理函數
# =========================================================================
def _initialize_ji_khoo(
    wb,
    new_jin_kang_piau_im_ji_khoo_sheet: bool,
    new_piau_im_ji_khoo_sheet: bool,
    new_khuat_ji_piau_sheet: bool,
) -> tuple[JiKhooDict, JiKhooDict, JiKhooDict]:
    """初始化字庫工作表"""

    # 人工標音字庫
    jin_kang_piau_im_sheet_name = '人工標音字庫'
    if new_jin_kang_piau_im_ji_khoo_sheet:
        delete_sheet_by_name(wb=wb, sheet_name=jin_kang_piau_im_sheet_name)
    jin_kang_piau_im_ji_khoo = JiKhooDict.create_ji_khoo_dict_from_sheet(
        wb=wb,
        sheet_name=jin_kang_piau_im_sheet_name
    )

    # 標音字庫
    piau_im_sheet_name = '標音字庫'
    if new_piau_im_ji_khoo_sheet:
        delete_sheet_by_name(wb=wb, sheet_name=piau_im_sheet_name)
    piau_im_ji_khoo = JiKhooDict.create_ji_khoo_dict_from_sheet(
        wb=wb,
        sheet_name=piau_im_sheet_name
    )

    # 缺字表
    khuat_ji_piau_name = '缺字表'
    if new_khuat_ji_piau_sheet:
        delete_sheet_by_name(wb=wb, sheet_name=khuat_ji_piau_name)
    khuat_ji_piau_ji_khoo = JiKhooDict.create_ji_khoo_dict_from_sheet(
        wb=wb,
        sheet_name=khuat_ji_piau_name
    )

    return jin_kang_piau_im_ji_khoo, piau_im_ji_khoo, khuat_ji_piau_ji_khoo


def _save_ji_khoo_to_excel(
    wb,
    jin_kang_piau_im_ji_khoo: JiKhooDict,
    piau_im_ji_khoo: JiKhooDict,
    khuat_ji_piau_ji_khoo: JiKhooDict,
):
    """儲存字庫到 Excel"""
    jin_kang_piau_im_ji_khoo.write_to_excel_sheet(wb=wb, sheet_name='人工標音字庫')
    piau_im_ji_khoo.write_to_excel_sheet(wb=wb, sheet_name='標音字庫')
    khuat_ji_piau_ji_khoo.write_to_excel_sheet(wb=wb, sheet_name='缺字表')


def process(wb, args) -> int:
    """
    更新【漢字注音】表中【台語音標】儲存格的內容，依據【標音字庫】中的【校正音標】欄位進行更新，並將【校正音標】覆蓋至原【台語音標】。
    Args:
        wb: Excel Workbook 物件
        args: 命令列參數

    Returns:
        處理結果代碼
    """
    #--------------------------------------------------------------------------
    # 作業初始化
    #--------------------------------------------------------------------------
    logging_process_step("<----------- 作業開始！---------->")

    try:
        #--------------------------------------------------------------------------
        # 初始化 process config
        #--------------------------------------------------------------------------
        config = ProcessConfig(wb, args, hanji_piau_im_sheet='漢字注音')

        # 建立字庫工作表
        if args.new:
            jin_kang_piau_im_ji_khoo_dict, piau_im_ji_khoo_dict, khuat_ji_piau_ji_khoo_dict = _initialize_ji_khoo(
                wb=wb,
                new_jin_kang_piau_im_ji_khoo_sheet=True,
                new_piau_im_ji_khoo_sheet=True,
                new_khuat_ji_piau_sheet=True,
            )
        else:
            jin_kang_piau_im_ji_khoo_dict, piau_im_ji_khoo_dict, khuat_ji_piau_ji_khoo_dict = _initialize_ji_khoo(
                wb=wb,
                new_jin_kang_piau_im_ji_khoo_sheet=False,
                new_piau_im_ji_khoo_sheet=False,
                new_khuat_ji_piau_sheet=False,
            )

        # 建立儲存格處理器
        processor = CellProcessor(
            config=config,
            jin_kang_piau_im_ji_khoo=jin_kang_piau_im_ji_khoo_dict,
            piau_im_ji_khoo=piau_im_ji_khoo_dict,
            khuat_ji_piau_ji_khoo=khuat_ji_piau_ji_khoo_dict,
        )

        #======================================================================
        # 處理作業
        #======================================================================
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
            # 寫回字庫到 Excel
            _save_ji_khoo_to_excel(
                wb=wb,
                jin_kang_piau_im_ji_khoo=jin_kang_piau_im_ji_khoo_dict,
                piau_im_ji_khoo=piau_im_ji_khoo_dict,
                khuat_ji_piau_ji_khoo=khuat_ji_piau_ji_khoo_dict,
            )
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
            # 寫回字庫到 Excel
            _save_ji_khoo_to_excel(
                wb=wb,
                jin_kang_piau_im_ji_khoo=jin_kang_piau_im_ji_khoo_dict,
                piau_im_ji_khoo=piau_im_ji_khoo_dict,
                khuat_ji_piau_ji_khoo=khuat_ji_piau_ji_khoo_dict,
            )
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
            # 寫回字庫到 Excel
            _save_ji_khoo_to_excel(
                wb=wb,
                jin_kang_piau_im_ji_khoo=jin_kang_piau_im_ji_khoo_dict,
                piau_im_ji_khoo=piau_im_ji_khoo_dict,
                khuat_ji_piau_ji_khoo=khuat_ji_piau_ji_khoo_dict,
            )
        except Exception as e:
            logging_exc_error(msg=f"處理以【標音字庫】更新【漢字注音】工作表之作業，發生執行異常！", error=e)
            return EXIT_CODE_PROCESS_FAILURE
        #--------------------------------------------------------------------------
        # 結束作業
        #--------------------------------------------------------------------------
        han_ji_piau_im_sheet.activate()
        logging_process_step("<=========== 完成處理流程作業！==========>")

        # =========================================================================
        # 儲存檔案
        # =========================================================================
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
        # 結束作業
        # =========================================================================
        print('=' * 80)
        logging_process_step("已完成【台語音標】和【漢字標音】標注工作。")
        return EXIT_CODE_SUCCESS

    except Exception as e:
        logging.exception("處理作業，發生例外！")
        raise


# =========================================================================
# 主程式
# =========================================================================
def main(args) -> int:
    """主程式 - 從 Excel 呼叫或直接執行"""
    # 取得專案根目錄。
    current_file_path = Path(__file__).resolve()
    project_root = current_file_path.parent
    # 取得程式名稱
    program_name = current_file_path.stem
    # 顯示程式開始訊息
    logging_process_step(f"《========== 程式開始執行：{program_name} ==========》")
    logging_process_step(f"專案根目錄為: {project_root}")

    try:
        # 取得 Excel 活頁簿
        wb = None
        try:
            # 嘗試從 Excel 呼叫取得（RunPython）
            wb = xw.Book.caller()
        except:
            # 若失敗，則取得作用中的活頁簿
            try:
                wb = xw.apps.active.books.active
            except Exception as e:
                logging.error(f"無法找到作用中的 Excel 工作簿: {e}")
                return EXIT_CODE_NO_FILE

        if not wb:
            logging.error("無法取得 Excel 活頁簿")
            return EXIT_CODE_NO_FILE

        # 執行處理
        exit_code = process(wb, args)

        # 顯示程式結束訊息
        logging_process_step(f"《========== 程式終止執行：{program_name} ==========》")
        return exit_code

    except Exception as e:
        logging.exception("程式執行失敗")
        return EXIT_CODE_UNKNOWN_ERROR


def test_01():
    """
    測試程式主要作業流程
    """
    print("\n\n")
    print("=" * 100)
    print("執行測試：test_01()")
    print("=" * 100)
    # 執行主要作業流程
    return EXIT_CODE_SUCCESS


if __name__ == "__main__":
    import argparse
    import sys

    # 解析命令行參數
    parser = argparse.ArgumentParser(
        description='手動更正漢字標音：在【漢字注音】工作表中，依據【標音字庫】工作表之【校正音標】欄位內容，更新【台語音標】及【漢字標音】欄位。',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
使用範例：
  python a300.py          # 執行一般模式
  python a300.py -new     # 建立新的字庫工作表
  python a300.py -test    # 執行測試模式
'''
        )
    parser.add_argument(
        '--test',
        action='store_true',
        help='執行測試模式',
    )
    parser.add_argument(
        '--new',
        action='store_true',
        help='建立新的標音字庫工作表',
    )
    args = parser.parse_args()

    if args.test:
        # 執行測試
        sys.exit(test_01())
    else:
        # 從 Excel 呼叫
        sys.exit(main(args))
