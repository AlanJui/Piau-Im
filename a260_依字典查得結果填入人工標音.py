"""
a260_依字典查得結果填入人工標音.py V0.0.1

在【漢字注音】工作表之【作用儲存格】，可以兩種方式輸入【人工標音】資料：
（1）自【自用字典】查得【台語音標】；（2）直接手動輸入【台語音標】/【台羅拼音】。

修改紀錄：
v0.0.1 2026-2-28: 初始版本，完成基本功能。
"""

# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import sys
from pathlib import Path

# 載入第三方套件
import xlwings as xw

from mod_excel_access import (
    excel_address_to_row_col,
    get_active_cell,
    get_active_cell_address,
    get_line_no_by_row,
    get_row_by_line_no,
)
from mod_logging import (
    init_logging,
    logging_exc_error,
    logging_exception,
    logging_process_step,
    logging_warning,  # noqa: F401
)

# 載入自訂模組/函式
from mod_程式 import ExcelCell, Program

# =========================================================================
# 設定日誌
# =========================================================================
init_logging()

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
# 資料類別：儲存處理配置
# =========================================================================
class CellProcessor(ExcelCell):
    """
    個人字典查詢專用的儲存格處理器
    繼承自 ExcelCell
    覆蓋以下方法以實現個人字典查詢功能：
    - _process_cell(): 處理單一儲存格
    - _process_jin_kang_piau_im(): 處理人工標音邏輯
    其他方法繼承自父類別 ExcelCell
    """

    def __init__(
        self,
        program: Program,
        new_jin_kang_piau_im_ji_khoo_sheet: bool = False,
        new_piau_im_ji_khoo_sheet: bool = False,
        new_khuat_ji_piau_sheet: bool = False,
    ):
        """
        初始化處理器
        :param config: 設定檔物件 (包含標音方法、資料庫連線等)
        :param jin_kang_ji_khoo: 人工標音字庫 (JiKhooDict) - 用於 '=' 查找
        :param piau_im_ji_khoo: 標音字庫
        :param khuat_ji_piau_ji_khoo: 缺字表
        """
        # 調用父類別（MengDianExcelCell）的建構子
        super().__init__(
            program=program,
            new_jin_kang_piau_im_ji_khoo_sheet=new_jin_kang_piau_im_ji_khoo_sheet,
            new_piau_im_ji_khoo_sheet=new_piau_im_ji_khoo_sheet,
            new_khuat_ji_piau_sheet=new_khuat_ji_piau_sheet,
        )

    def _za_ji_tian(self, cell):
        """
        從【自用字典】查詢【台語音標】
        """
        han_ji = cell.value
        tai_gi_im_piau = ""

        if han_ji == "":
            return tai_gi_im_piau

        # (1) 查字典：使用 HanJiTian 類別查詢漢字讀音
        result = self.program.ji_tian.han_ji_ca_piau_im(
            han_ji=han_ji,
            ue_im_lui_piat=self.program.ue_im_lui_piat,
        )

        # 查無此字
        if not result:
            print(f">> 漢字【{han_ji}】查不到讀音資料！")
            return tai_gi_im_piau

        # (2) 在 console 列出字典中，查詢之漢字有那些讀音選項及其常用程度

        # 顯示所有讀音選項
        piau_im_options = self.display_all_piau_im_for_a_han_ji(han_ji, result)

        # (3) 供使用者輸入選擇
        user_input = input("\n請輸入選擇編號 (直接按 Enter 跳過): ").strip()

        if not user_input:
            print(">> 放棄變更！")
            return None

        try:
            # 取得使用者之輸入，並【解析】其輸入是要：（1）引用字典的查找結果；
            # （2）直接輸入【台語音標】或【台羅拼音】
            choice = int(user_input)

            # 解析使用者輸入：
            # （1）【引用字典查找結果】判斷規則：輸入為【數值】，且落在字典查找結果的選項範圍內；
            # （2）【直接輸入台語音標或台羅拼音】判斷規則：輸入為非數值，或數值不在選項範圍內
            case = None

            if case == 1:
                # （1）引用字典查找結果
                if 1 <= choice <= len(piau_im_options):
                    # 顯示使用者輸入之讀音選項
                    print(f"【{han_ji}】讀音，選用：第 {choice} 個選項。")

                    # 依據輸入之【數值】，自讀音選項清單(piau_im_options)，取得對映之【台語音標】及【漢字標音】
                    selected_im_piau, selected_han_ji_piau_im = piau_im_options[
                        choice - 1
                    ]

                    # return [selected_im_piau, selected_han_ji_piau_im]
                    return selected_im_piau
                else:
                    print(f">> 輸入錯誤：{choice} 超出範圍！")
                    return None
            elif case == 2:
                # （2）直接輸入【台語音標】或【台羅拼音】
                # TODO:
                # 1. 解析使用者輸入的【台語音標】或【台羅拼音】，並驗證其格式是否正確。
                # 2. 若格式正確，則將其作為【台語音標】返回；若格式不正確，則提示使用者輸入錯誤。
                return tai_gi_im_piau  # 這裡應該是要返回使用者直接輸入的【台語音標】或【台羅拼音】，但目前尚未實作解析邏輯，因此先返回空字串。
        except ValueError:
            print(f">> 使用者輸入格式有誤：{user_input}")
            return None

        return tai_gi_im_piau

    def _za_ji_tain_au_thiam_jin_kang_piau_im(self, active_cell):
        """查字典後填入工標音"""
        tai_gi_im_piau = ""
        han_ji_piau_im = ""

        # 依據【作用儲存格】之【漢字】，從【自用字典】查詢【台語音標】
        # han_ji = active_cell.value
        tai_gi_im_piau = self._za_ji_tian(active_cell)
        active_cell.offset(-2, 0).value = tai_gi_im_piau  # 人工標音

        self._process_jin_kang_piau_im(cell=active_cell)

        return tai_gi_im_piau, han_ji_piau_im


# =============================================================================
# 作業主流程
# =============================================================================
def process(wb, args) -> int:
    """
    作業流程：
    1. 取得當前 Excel 作用儲存格 (漢字、座標)
    2. 計算【人工標音】位置與值
    3. 查詢【標音字庫】確認該座標是否已登錄
    4. 若【標正音標】為 'N/A'，則更新為【人工標音】

    Args:
        wb: Excel Workbook 物件
        args: 命令列參數

    Returns:
        處理結果代碼
    """
    # --------------------------------------------------------------------------
    # 作業初始化
    # --------------------------------------------------------------------------
    logging_process_step("<=========== 作業開始！==========>")

    try:
        # --------------------------------------------------------------------------
        # 初始化 process config
        # --------------------------------------------------------------------------
        program = Program(wb, args, hanji_piau_im_sheet_name="漢字注音")

        # 建立儲存格處理器
        xls_cell = CellProcessor(
            program=program,
            new_jin_kang_piau_im_ji_khoo_sheet=False,
            new_piau_im_ji_khoo_sheet=False,
            new_khuat_ji_piau_sheet=False,
        )

        # --------------------------------------------------------------------------
        # 處理作業開始
        # --------------------------------------------------------------------------
        source_sheet_name = "漢字注音"
        jin_kang_piau_im_sheet_name = "人工標音字庫"
        piau_im_ji_khoo_sheet_name = "標音字庫"

        # ----------------------------------------------------------------------
        # 取得【作用儲存格】
        # ----------------------------------------------------------------------
        # 指定【漢字注音】工作表為【作用工作表】
        source_sheet = wb.sheets[source_sheet_name]
        source_sheet.activate()

        active_cell_address = get_active_cell_address()
        active_cell = source_sheet.range(active_cell_address)
        row, col = excel_address_to_row_col(active_cell_address)
        current_line_no = get_line_no_by_row(current_row_no=row)  # 計算行號
        jin_kang_piau_im_row, tai_gi_im_piau_row, han_ji_row, han_ji_piau_im_row = (
            get_row_by_line_no(current_line_no)
        )

        han_ji = source_sheet.range((han_ji_row, col)).value
        jin_kang_piau_im = source_sheet.range((jin_kang_piau_im_row, col)).value
        tai_gi_im_piau = source_sheet.range((tai_gi_im_piau_row, col)).value
        han_ji_piau_im = source_sheet.range((han_ji_piau_im_row, col)).value
        han_ji_position = (han_ji_row, col)
        han_ji_cell = source_sheet.range((han_ji_row, col))

        print(
            f"📌 作用儲存格：{active_cell_address} ==> 漢字儲存格座標：{han_ji_position}"
        )
        print(f"📌 漢字：{han_ji}")
        print(
            f"📌 人工標音：{jin_kang_piau_im}，台語音標：{tai_gi_im_piau}，漢字標音：{han_ji_piau_im}"
        )

        # ----------------------------------------------------------------------
        # 查字典後填人工標音
        # Za-Ji-Tain-Au-Thiam-Jin-Kang-Piau-Im
        # ----------------------------------------------------------------------
        tai_gi_im_piau, han_ji_piau_im = xls_cell._za_ji_tain_au_thiam_jin_kang_piau_im(
            active_cell=active_cell,
        )

        # 將【台語音標】和【漢字標音】寫入【漢字注音】工作表之【作用儲存格】
        han_ji_cell.offset(-2, 0).value = tai_gi_im_piau  # 人工標音
        han_ji_cell.offset(-1, 0).value = tai_gi_im_piau  # 台語音標
        han_ji_cell.offset(+1, 0).value = han_ji_piau_im  # 漢字標音
        msg = f"{han_ji}： [{jin_kang_piau_im}] / [{tai_gi_im_piau}] /【{han_ji_piau_im}】"
        print(f"✅ 已更新儲存格：{active_cell_address}，內容為：{msg}")

        # 調整 row 指向【漢字】儲存格所在座標列
        row = han_ji_row

        # -------------------------------------------------------------------------
        # 在【人工標音字庫】工作表對映之【字庫】(dict)，添加或更新一筆【漢字】及
        # 【台語音標】資料
        # -------------------------------------------------------------------------
        xls_cell.jin_kang_piau_im_ji_khoo_dict.add_entry(
            han_ji=han_ji,
            tai_gi_im_piau=tai_gi_im_piau,
            hau_ziann_im_piau=jin_kang_piau_im,
            coordinates=(row, col),
        )
        # -------------------------------------------------------------------------
        # 自【標音字庫】工作表對映之【字庫】(dict)，移除該【漢字】之【座標】資料
        # -------------------------------------------------------------------------
        xls_cell.piau_im_ji_khoo_dict.remove_coordinate(
            han_ji=han_ji,
            coordinate=(row, col),
        )
        # -------------------------------------------------------------------------
        # 更新資料庫中【漢字庫】資料表
        # -------------------------------------------------------------------------
        siong_iong_too_to_use = (
            0.8 if program.ue_im_lui_piat == "文讀音" else 0.6
        )  # 根據語音類型設定常用度
        xls_cell.insert_or_update_to_db(
            table_name=program.table_name,
            han_ji=han_ji,
            tai_gi_im_piau=tai_gi_im_piau,
            ue_im_lui_piat=program.ue_im_lui_piat,
            siong_iong_too=siong_iong_too_to_use,
        )

        # ----------------------------------------------------------------------
        # 將【標音字庫】、【人工標音字庫】，寫入 Excel 工作表
        # ----------------------------------------------------------------------
        xls_cell.piau_im_ji_khoo_dict.write_to_excel_sheet(
            wb=wb, sheet_name=piau_im_ji_khoo_sheet_name
        )
        xls_cell.jin_kang_piau_im_ji_khoo_dict.write_to_excel_sheet(
            wb=wb, sheet_name=jin_kang_piau_im_sheet_name
        )

        logging_process_step(msg="已完成【台語音標】和【漢字標音】標注工作。")
        return EXIT_CODE_SUCCESS
    except Exception as e:
        # 你可以在這裡加上紀錄或處理，例如:
        logging_exception(msg="自動為【漢字】查找【台語音標】作業，發生例外！", error=e)
        # 再次拋出異常，讓外層函式能捕捉
        raise


# =========================================================================
# 主程式
# =========================================================================
def main(args) -> int:
    # =========================================================================
    # (0) 程式初始化
    # =========================================================================
    # 取得專案根目錄。
    current_file_path = Path(__file__).resolve()
    project_root = current_file_path.parent
    # 取得程式名稱
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
        wb = xw.apps.active.books.active  # 取得 Excel 作用中的活頁簿檔案
    except Exception as e:
        msg = "無法找到作用中的 Excel 工作簿！"
        logging_exception(msg=msg, error=e)
        return EXIT_CODE_NO_FILE

    # 若無法取得【作用中活頁簿】，則因無法繼續作業，故返回【作業異常終止代碼】結束。
    if not wb:
        return EXIT_CODE_NO_FILE

    # =========================================================================
    # (3) 執行【處理作業】
    # =========================================================================
    try:
        exit_code = process(wb, args)
    except Exception as e:
        msg = f"程式異常終止：{program_name}"
        logging_exception(msg=msg, error=e)
        return EXIT_CODE_UNKNOWN_ERROR

    if exit_code != EXIT_CODE_SUCCESS:
        msg = f"程式異常終止：{program_name}（非例外，而是返回失敗碼）"
        logging_exc_error(msg=msg, error=None)
        return EXIT_CODE_PROCESS_FAILURE

    # =========================================================================
    # (4) 儲存檔案
    # =========================================================================
    try:
        # 要求畫面回到【漢字注音】工作表
        # wb.sheets['漢字注音'].activate()
        # 儲存檔案
        wb.save()
        file_path = wb.fullname
        logging_process_step(f"儲存檔案至路徑：{file_path}")

    except Exception as e:
        logging_exception(msg="儲存檔案失敗！", error=e)
        return EXIT_CODE_SAVE_FAILURE  # 作業異當終止：無法儲存檔案

    # =========================================================================
    # (5) 結束作業
    # =========================================================================
    return EXIT_CODE_SUCCESS


def ut01():
    # 取得【作用中活頁簿】
    wb = None
    # 取得【作用中活頁簿】
    try:
        wb = xw.apps.active.books.active  # 取得 Excel 作用中的活頁簿檔案
    except Exception as e:
        msg = "無法找到作用中的 Excel 工作簿！"
        logging_exception(msg=msg, error=e)
        return EXIT_CODE_NO_FILE
    # 作業流程：獲取當前作用中的 Excel 儲存格
    sheet_name, cell_address = get_active_cell(wb)
    print(f"✅ 目前作用中的儲存格：{sheet_name} 工作表 -> {cell_address}")

    # 將 Excel 儲存格地址轉換為 (row, col) 格式
    row, col = excel_address_to_row_col(cell_address)
    print(f"📌 Excel 位址 {cell_address} 轉換為 (row, col): ({row}, {col})")

    return EXIT_CODE_SUCCESS


# =============================================================================
# 程式作業模式切換
# =============================================================================
if __name__ == "__main__":
    import argparse
    import sys

    # 解析命令行參數
    parser = argparse.ArgumentParser(
        description="缺字表修正後續作業程式",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用範例：
  python a000.py          # 執行一般模式
  python a000.py -new     # 建立新的字庫工作表
  python a000.py -test    # 執行測試模式
""",
    )
    parser.add_argument(
        "--test",
        action="store_true",
        help="執行測試模式",
    )
    args = parser.parse_args()

    if args.test:
        # 執行測試
        ut01()
    else:
        # 從 Excel 呼叫
        exit_code = main(args)
        if exit_code == EXIT_CODE_SUCCESS:
            print("程式正常完成！")
        else:
            print(f"程式異常終止，錯誤代碼為: {exit_code}")
            sys.exit(exit_code)
