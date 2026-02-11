"""
a310_缺字表修正後續作業.py v0.2.7
"""

# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import os
import re
import sys
from pathlib import Path

# 載入第三方套件
import xlwings as xw
from dotenv import load_dotenv

# from mod_excel_access import save_as_new_file
# from mod_帶調符音標 import kam_si_u_tiau_hu, tng_im_piau, tng_tiau_ho
# from mod_標音 import (
#     PiauIm,  # 漢字標音物件
#     convert_tlpa_to_tl,
#     tlpa_tng_han_ji_piau_im,  # 台語音標轉台語音標
# )
# 載入自訂模組/函式
from mod_excel_access import convert_coord_str_to_excel_address
from mod_標音 import kam_si_u_tiau_hu, tlpa_tng_han_ji_piau_im, tng_im_piau, tng_tiau_ho
from mod_程式 import ExcelCell, Program

# =========================================================================
# 載入環境變數
# =========================================================================
load_dotenv()

# 預設檔案名稱從環境變數讀取
DB_HO_LOK_UE = os.getenv("DB_HO_LOK_UE", "Ho_Lok_Ue.db")
DB_KONG_UN = os.getenv("DB_KONG_UN", "Kong_Un.db")

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
from mod_logging import (  # noqa: E402
    init_logging,
    logging_exc_error,
    logging_exception,
    logging_process_step,
    logging_warning,
)

init_logging()


# =========================================================================
# 主要處理函數
# =========================================================================
class CellProcessor(ExcelCell):
    """
    本程式專用的儲存格處理器
    繼承自 ExcelCell 的類別
    覆蓋以下方法以實現萌典查詢功能：
    - _process_jin_kang_piau_im(): 處理人工標音內容
    - _process_han_ji(): 使用【個人字典】查詢漢字讀音
    - _process_cell(): 處理單一儲存格
    - _process_sheet(): 處理整個工作表
    """

    def __init__(
        self,
        program: Program,
        new_jin_kang_piau_im_ji_khoo_sheet: bool = False,
        new_piau_im_ji_khoo_sheet: bool = False,
        new_khuat_ji_piau_sheet: bool = False,
    ):
        # 調用父類別（MengDianExcelCell）的建構子
        super().__init__(
            program=program,
            new_jin_kang_piau_im_ji_khoo_sheet=new_jin_kang_piau_im_ji_khoo_sheet,
            new_piau_im_ji_khoo_sheet=new_piau_im_ji_khoo_sheet,
            new_khuat_ji_piau_sheet=new_khuat_ji_piau_sheet,
        )

    # =================================================================
    # 覆蓋父類別的方法
    # =================================================================
    def update_hanji_zu_im_sheet_by_khuat_ji_piau(
        self, source_sheet_name: str, target_sheet_name: str
    ) -> int:
        """
        讀取 Excel 檔案，依據【缺字表】工作表中的資料執行下列作業：
        1. 由 A 欄讀取漢字，從 C 欄取得原始輸入之【校正音標】，並轉換為 TLPA+ 格式，然後更新 B 欄（台語音標）。
        2. 從 D 欄讀取座標字串（可能含有多組座標），每組座標指向【漢字注音】工作表中該漢字儲存格，
            將【缺字表】取得之【台語音標】，填入【漢字注音】工作表之【台語音標】欄位（於【漢字】儲存格上方一列（row - 1））;
            並在【漢字】儲存格下方一列（row + 1）填入【漢字標音】。
        """
        # 取得【標音方法】
        wb = self.program.wb
        piau_im_huat = self.program.piau_im_huat
        # 取得【漢字標音】物件
        piau_im = self.program.piau_im

        # -------------------------------------------------------------------------
        # 檢驗工作表是否存在
        # -------------------------------------------------------------------------
        try:
            # 來源、目標工作表
            source_sheet = wb.sheets[source_sheet_name]
            target_sheet = wb.sheets[target_sheet_name]
            # 取得【來源工作表】：【標音字庫】查詢表（dict）
            source_dict = self.get_piau_im_dict_by_name(sheet_name=source_sheet_name)
            target_dict = self.get_piau_im_dict_by_name(sheet_name="標音字庫")
        except Exception as e:
            logging_exc_error(msg="找不到工作表 ！", error=e)
            return EXIT_CODE_PROCESS_FAILURE

        # -------------------------------------------------------------------------
        # 在【缺字表】工作表中，逐列讀取資料進行處理：【校正音標】欄（C）有填音標者，
        # 將【校正音標】正規化為 TLPA+ 格式，並更新【台語音標】欄（B）；
        # 並依據【座標】欄（D）內容，將【校正音標】填入【漢字注音】工作表中相對應之【台語音標】儲存格，
        # 以及使用【校正音標】轉換後之【漢字標音】填入【漢字注音】工作表中相對應之【漢字標音】儲存格。
        # -------------------------------------------------------------------------
        row = 2  # 從第 2 列開始（跳過標題列）
        while True:
            han_ji = source_sheet.range(f"A{row}").value  # 讀取 A 欄（漢字）
            if not han_ji:  # 若 A 欄為空，則結束迴圈
                break

            # 取得原始【台語音標】並轉換為 TLPA+ 格式
            # org_tai_gi_im_piau = source_sheet.range(f"B{row}").value
            tai_gi_im_piau = source_sheet.range(f"B{row}").value

            # 將【校正音標】欄（C 欄）填入之音標，視作【漢字】之【台語音標】，
            # 此音標需符合 TLPA+ 格式（無調符，使用數值調號；沒有台羅拼音的聲母：ts/tsh），
            # 剛取得使用者輸入之【校正音標】，視作【原始台語音標】，需經：（1）轉換成 TLPA 拼音字母；
            # （2）去調符轉調號，依此作業方式，達成符合 TLPA+ 格式。
            org_tai_gi_im_piau = source_sheet.range(f"C{row}").value
            if (
                org_tai_gi_im_piau == "N/A" or not org_tai_gi_im_piau
            ):  # 若【台語音標】欄為空，則結束迴圈
                row += 1
                continue

            # 確保【原始台語音標】為 TLPA+ 格式（無調符，使用數值調號；沒有台羅拼音的聲母：ts/tsh）
            if org_tai_gi_im_piau and kam_si_u_tiau_hu(
                org_tai_gi_im_piau
            ):  # 有【調符】時轉換成【TLPA+格式拼音字母】
                tlpa_im_piau = tng_im_piau(org_tai_gi_im_piau)
                tlpa_im_piau = tng_tiau_ho(tlpa_im_piau).lower()
            else:  # 若非帶調符音標，則直接使用原音標
                tlpa_im_piau = org_tai_gi_im_piau

            # 以此符合 TLPA+ 格式之拼音字母，作為真正使用之【校正音標】。
            hau_ziann_im_piau = tlpa_im_piau

            # 讀取【缺字表】中【座標】欄（D 欄），取得指向【漢字注音】工作表【漢字】之【座標清單】。
            # 欄中內容【格式】，如： "(5, 17); (33, 8); (77, 5)"
            coordinates_str = source_sheet.range(f"D{row}").value
            excel_address_str = convert_coord_str_to_excel_address(
                coord_str=coordinates_str
            )  # B欄（台語音標）儲存格位置
            print("\n")
            print(
                f"{row-1}. (A{row}) 【{han_ji}】：台語音標：{tai_gi_im_piau}, 校正音標：{hau_ziann_im_piau} ==> 【{target_sheet_name}】工作表，儲存格：{excel_address_str} {coordinates_str}"
            )

            # 將【座標】欄位內容解析成 (row, col) 座標：此座標指向【漢字注音】工作表中之【漢字】儲存格位置
            if coordinates_str:
                # 利用正規表達式解析所有形如 (row, col) 的座標
                coordinate_tuples = re.findall(
                    r"\((\d+)\s*,\s*(\d+)\)", coordinates_str
                )
                for tup in coordinate_tuples:
                    try:
                        r_coord = int(tup[0])
                        c_coord = int(tup[1])
                    except ValueError:
                        continue  # 若轉換失敗，跳過該組座標

                    # 指向【漢字注音】工作表，【漢字儲存格】座標
                    han_ji_cell = (r_coord, c_coord)

                    # 根據說明，【台語音標】應填入漢字儲存格上方一列 (row - 1)，相同欄位
                    target_row = r_coord - 1
                    tai_gi_im_piau_cell = (target_row, c_coord)

                    # 對指向【漢字注音】工作表之【漢字儲存格】，填入漢字之【台語音標】
                    tai_gi_im_piau = hau_ziann_im_piau  # 以【校正音標】作為【台語音標】，【漢字注音】工作表之【台語音標】欄位
                    target_sheet.range(tai_gi_im_piau_cell).value = tai_gi_im_piau
                    excel_address_str = target_sheet.range(tai_gi_im_piau_cell).address
                    excel_address_str = excel_address_str.replace(
                        "$", ""
                    )  # 去除 "$" 符號
                    print(
                        f"   台語音標：【{tai_gi_im_piau}】，填入【{target_sheet_name}】工作表之儲存格： {excel_address_str} {tai_gi_im_piau_cell}"
                    )

                    # 轉換【台語音標】，取得【漢字標音】
                    han_ji_piau_im = tlpa_tng_han_ji_piau_im(
                        piau_im=piau_im,
                        piau_im_huat=piau_im_huat,
                        tai_gi_im_piau=tai_gi_im_piau,
                    )

                    # 將【漢字標音】填入【漢字注音】工作表，【漢字】儲存格下之【漢字標音】儲存格（即：row + 1)
                    target_row = r_coord + 1
                    han_ji_piau_im_cell = (target_row, c_coord)

                    # 將【校正音標】填入【漢字注音】工作表漢字之【台語音標】儲存格
                    target_sheet.range(han_ji_piau_im_cell).value = han_ji_piau_im
                    excel_address_str = target_sheet.range(han_ji_piau_im_cell).address
                    excel_address_str = excel_address_str.replace(
                        "$", ""
                    )  # 去除 "$" 符號
                    print(
                        f"   漢字標音：【{han_ji_piau_im}】，填入【{target_sheet_name}】工作表之儲存格： {excel_address_str} {han_ji_piau_im_cell}\n"
                    )

                    # 將【漢字注音】工作表之【漢字】儲存格之底色，重置為【無底色】
                    target_sheet.range(han_ji_cell).color = None

                    # ------------------------------------------------------------------------
                    # 以【缺字表】工作表之【漢字】+【台語音標】作為【資料紀錄索引】，
                    # ------------------------------------------------------------------------
                    # 在【標音字庫】工作表【添增】此筆資料紀錄
                    # hau_ziann_im_piau_to_be = 'N/A' if hau_ziann_im_piau == '' else hau_ziann_im_piau
                    # hau_ziann_im_piau_to_be = "N/A"
                    self.tiau_zing_piau_im_ji_khoo_dict(
                        han_ji=han_ji,
                        tai_gi_im_piau=hau_ziann_im_piau,  # 以【校正音標】作為【台語音標】，【漢字注音】工作表之【台語音標】欄位
                        hau_ziann_im_piau="N/A",  # 更新【校正音標】為 'N/A'
                        coordinates=(r_coord, c_coord),
                    )

                    # 將【座標】自【來源工作表】工作表（缺字表）的【座標】欄移除
                    # source_dict.remove_coordinate_by_hau_ziann_im_piau(
                    #     han_ji=han_ji,
                    #     hau_ziann_im_piau=hau_ziann_im_piau,
                    #     coordinate=(r_coord, c_coord)
                    # )
                    # source_dict.remove_coordinate(
                    #     han_ji=han_ji,
                    #     tai_gi_im_piau=org_tai_gi_im_piau,
                    #     coordinate=(r_coord, c_coord)
                    # )

                    # 更新【缺字表】工作表之【台語音標】欄（B欄）、【校正音標】欄（C欄）、【座標】欄（D欄）
                    source_dict.update_entry(
                        han_ji=han_ji,
                        tai_gi_im_piau="N/A",  # 更新【台語音標】為 'N/A'
                        hau_ziann_im_piau=hau_ziann_im_piau,
                        coordinates=(r_coord, c_coord),
                    )
            # 讀取下一列
            row += 1

        # 依據 Dict 內容，更新【標音字庫】、【缺字表】工作表之資料紀錄
        if row > 2:
            # 更新【標音字庫】工作表（【目標工作表】）
            sheet_name = "標音字庫"
            target_dict.write_to_excel_sheet(wb=wb, sheet_name=sheet_name)
            # 更新【缺字表】工作表（【來源工作表】）
            sheet_name = source_sheet_name
            source_dict.write_to_excel_sheet(wb=wb, sheet_name=sheet_name)
            return EXIT_CODE_SUCCESS
        else:
            logging_warning(
                msg=f"【{sheet_name}】工作表內，無任何資料，略過後續處理作業。"
            )
            return EXIT_CODE_INVALID_INPUT


def process(wb, args) -> int:
    """
    更新【漢字注音】表中【台語音標】儲存格的內容，依據【標音字庫】中的【校正音標】欄位進行更新，並將【校正音標】覆蓋至原【台語音標】。

    Args:
        wb: Excel Workbook 物件
        args: 命令列參數

    Returns:
        處理結果代碼
    """
    # --------------------------------------------------------------------------
    # 作業開始
    # --------------------------------------------------------------------------
    logging_process_step("<=========== 作業開始！==========>")

    # --------------------------------------------------------------------------
    # 作業初始化
    # --------------------------------------------------------------------------
    try:
        program = Program(wb, args, hanji_piau_im_sheet_name="漢字注音")

        # 建立儲存格處理器
        xls_cell = CellProcessor(
            program=program,
            new_jin_kang_piau_im_ji_khoo_sheet=False,
            new_piau_im_ji_khoo_sheet=False,
            new_khuat_ji_piau_sheet=False,
        )
    except Exception as e:
        logging_exc_error(msg="處理作業異常！", error=e)
        return EXIT_CODE_PROCESS_FAILURE

    # --------------------------------------------------------------------------
    # 處理作業開始
    # --------------------------------------------------------------------------
    source_sheet_name = "缺字表"
    target_sheet_name = "漢字注音"
    msg = f"使用【{source_sheet_name}】工作表，更新【{target_sheet_name}】工作表......"
    print("\n")
    print("=" * 80)
    logging_process_step(msg)

    try:
        sheet_name = source_sheet_name
        wb.sheets[sheet_name].activate()
        exit_code = xls_cell.update_hanji_zu_im_sheet_by_khuat_ji_piau(
            source_sheet_name=source_sheet_name,
            target_sheet_name=target_sheet_name,
        )
    except Exception as e:
        logging_exc_error(msg=f"處理【{sheet_name}】作業異常！", error=e)
        return EXIT_CODE_PROCESS_FAILURE

    if exit_code != EXIT_CODE_SUCCESS:
        return exit_code

    # -------------------------------------------------------------------------
    # 將【缺字表】之【漢字】與【台語音標】存入【漢字庫】作業
    # -------------------------------------------------------------------------
    # sheet_name = source_sheet_name
    # msg = f"使用【{sheet_name}】工作表，更新資料庫中之【漢字庫】資料表......"
    # print("\n")
    # print("=" * 80)
    # logging_process_step(msg)
    #
    # try:
    #     sheet_name = source_sheet_name
    #     wb.sheets[sheet_name].activate()
    #     xls_cell.update_han_ji_khoo_db_by_sheet(sheet_name=sheet_name)
    # except Exception as e:
    #     logging_exc_error(
    #         msg=f"將【{sheet_name}】之【漢字】與【台語音標】存入【漢字庫】作業，發生執行異常！",
    #         error=e,
    #     )
    #     return EXIT_CODE_PROCESS_FAILURE
    # finally:
    #     # 關閉資料庫連線
    #     if xls_cell.db_manager:
    #         xls_cell.db_manager.disconnect()
    #         logging_process_step("已關閉資料庫連線")
    #
    # print("\n")
    # print("-" * 80)
    # logging_process_step(
    #     f"完成：將【{sheet_name}】之【漢字】與【台語音標】存入【漢字庫】作業"
    # )
    # print("=" * 80)

    # --------------------------------------------------------------------------
    # 作業結束
    # --------------------------------------------------------------------------
    print("\n")
    logging_process_step("<=========== 作業結束！==========>")
    return EXIT_CODE_SUCCESS


# =========================================================================
# 程式主要作業流程
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
        logging_exception(msg="作業異常終止！", error=e)
        return EXIT_CODE_PROCESS_FAILURE

    # =========================================================================
    # (4) 儲存檔案
    # =========================================================================
    if exit_code == EXIT_CODE_SUCCESS:
        try:
            wb.save()
            file_path = wb.fullname
            logging_process_step(f"儲存檔案至路徑：{file_path}")
        except Exception as e:
            logging_exc_error(msg="儲存檔案異常！", error=e)
            return EXIT_CODE_SAVE_FAILURE

    # =========================================================================
    # 結束程式
    # =========================================================================
    print("\n")
    print("=" * 80)
    logging_process_step(f"《========== 程式終止執行：{program_name} ==========》")
    if exit_code == EXIT_CODE_SUCCESS:
        return EXIT_CODE_SUCCESS  # 作業正常結束
    else:
        msg = f"程式異常終止，返回失敗碼：{exit_code}"
        logging_exc_error(msg=msg, error=None)
        return EXIT_CODE_PROCESS_FAILURE


# =============================================================================
# 測試程式
# =============================================================================
def test_01() -> int:
    """
    測試程式主要作業流程
    """
    print("\n\n")
    print("=" * 100)
    print("執行測試：test_01()")
    print("=" * 100)
    # 執行主要作業流程
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
    parser.add_argument(
        "--new",
        action="store_true",
        help="建立新的標音字庫工作表",
    )
    args = parser.parse_args()

    if args.test:
        # 執行測試
        exit_code = test_01()
    else:
        # 從 Excel 呼叫
        exit_code = main(args)

    # 只在命令列執行時使用 sys.exit()，避免在調試環境中引發 SystemExit 例外
    if exit_code != EXIT_CODE_SUCCESS:
        sys.exit(exit_code)
