"""
a251_三合一字典更換漢字讀音.py

依作用儲存格之漢字，向三合一漢字典（han-ji-tian）API 查詢讀音，
供使用者挑選後，以與 a250 相同流程更新【標音字庫】／【漢字注音】。

與 a250 差異：字典來源改為線上 `/api/v1/reading`（雅俗通／甘／廣韻），
而非本機 Ho_Lok_Ue.db。

憑證見專案根目錄 `.api_key`（範本：`.api_key.example`）。
"""

from __future__ import annotations

import logging
import sys
from pathlib import Path

import xlwings as xw

from a250_更換漢字讀音 import CellProcessor as A250CellProcessor
from a250_更換漢字讀音 import (
    EXIT_CODE_NO_FILE,
    EXIT_CODE_PROCESS_FAILURE,
    EXIT_CODE_SAVE_FAILURE,
    EXIT_CODE_SUCCESS,
)
from mod_han_ji_tian_api import (
    format_grouped_reading_lines,
    han_ji_ca_piau_im_from_api,
    load_api_key_file,
)
from mod_logging import (
    init_logging,
    logging_exc_error,
    logging_exception,
    logging_process_step,
)
from mod_程式 import Program

init_logging()


class CellProcessor(A250CellProcessor):
    """沿用 a250 訂正流程；置換查字典來源，並依字典分組顯示。"""

    def display_all_piau_im_for_a_han_ji(self, han_ji: str, result: list[dict]):
        """依《雅俗通》《甘字典》《廣韻》分組顯示讀音選項。"""
        piau_im_options = []
        for entry in result:
            tai_gi_im_piau, han_ji_piau_im = self._convert_piau_im_by_entry(entry)
            piau_im_options.append((tai_gi_im_piau, han_ji_piau_im))
        format_grouped_reading_lines(han_ji, result, piau_im_options)
        return piau_im_options

    def _han_ji_ca_piau_im_kap_cu_tik(self, cell):
        original = self.program.ji_tian.han_ji_ca_piau_im

        def _api_lookup(
            han_ji,
            ue_im_lui_piat="文讀音",
            display_all_piau_im=False,
            tai_lo_im_piau=None,
        ):
            print("📡 查詢三合一漢字典 API…")
            return han_ji_ca_piau_im_from_api(
                han_ji=han_ji,
                ue_im_lui_piat=ue_im_lui_piat,
                display_all_piau_im=display_all_piau_im,
                tai_lo_im_piau=tai_lo_im_piau,
            )

        self.program.ji_tian.han_ji_ca_piau_im = _api_lookup
        try:
            return super()._han_ji_ca_piau_im_kap_cu_tik(cell)
        finally:
            self.program.ji_tian.han_ji_ca_piau_im = original


def process(wb, args) -> int:
    """查詢三合一字典讀音並標注（介面同 a250.process）。"""
    logging_process_step("<=========== 作業開始！（三合一字典 a251）==========>")
    if load_api_key_file() is None:
        print("⚠️  找不到 `.api_key`（可複製 `.api_key.example`）。仍可嘗試系統環境變數。")

    try:
        program = Program(wb=wb, args=args, hanji_piau_im_sheet_name="漢字注音")
        xls_cell = CellProcessor(
            program=program,
            new_jin_kang_piau_im_ji_khoo_sheet=False,
            new_piau_im_ji_khoo_sheet=False,
            new_khuat_ji_piau_sheet=False,
        )
    except Exception as e:
        logging_exc_error(msg="處理作業異常！", error=e)
        return EXIT_CODE_PROCESS_FAILURE

    try:
        xls_cell._process_sheet()
    except Exception as e:
        logging.error(f"處理錯誤：{e}")
        print(f"❌ 錯誤：{e}")

    print("\n")
    logging_process_step("<=========== 作業結束！（三合一字典 a251）==========>")
    return EXIT_CODE_SUCCESS


def main(args) -> int:
    current_file_path = Path(__file__).resolve()
    project_root = current_file_path.parent
    program_name = current_file_path.stem

    logging_process_step(f"《========== 程式開始執行：{program_name} ==========》")
    logging_process_step(f"專案根目錄為: {project_root}")

    try:
        wb = xw.apps.active.books.active
    except Exception as e:
        logging_exception(msg="無法找到作用中的 Excel 工作簿！", error=e)
        return EXIT_CODE_NO_FILE

    if not wb:
        return EXIT_CODE_NO_FILE

    try:
        exit_code = process(wb, args)
    except Exception as e:
        logging_exception(msg=f"作業程序發生異常，終止執行：{program_name}", error=e)
        return EXIT_CODE_PROCESS_FAILURE

    if exit_code != EXIT_CODE_SUCCESS:
        logging_exc_error(
            msg=f"處理作業發生異常，終止程式執行：{program_name}（處理作業程序，返回失敗碼）"
        )
        return EXIT_CODE_PROCESS_FAILURE

    try:
        wb.save()
        logging_process_step(f"儲存檔案至路徑：{wb.fullname}")
    except Exception as e:
        logging_exception(msg="儲存檔案失敗！", error=e)
        return EXIT_CODE_SAVE_FAILURE

    return EXIT_CODE_SUCCESS


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description="透過【作用儲存格】，向三合一漢字典 API 查詢並更換漢字讀音",
    )
    parser.add_argument("--test", action="store_true", help="僅測試 API 查詢（不寫 Excel）")
    parser.add_argument("--han", default="行", help="--test 時要查的漢字")
    args = parser.parse_args()

    if args.test:
        load_api_key_file()
        rows = han_ji_ca_piau_im_from_api(args.han) or []
        print(f"【{args.han}】API 讀音 {len(rows)} 筆：")
        for i, row in enumerate(rows, 1):
            print(
                f"  {i}. {row.get('台語音標')} (常用度: {row.get('常用度')}) "
                f"{row.get('摘要說明', '')[:40]}"
            )
        sys.exit(0)

    code = main(args)
    if code != EXIT_CODE_SUCCESS:
        print(f"程式異常終止，返回代碼：{code}")
        sys.exit(code)
