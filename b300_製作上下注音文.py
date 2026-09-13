"""b200_製作字幕檔.py v0.1.0

將工作表 I 欄【可匯出SRT】內容輸出成 SRT 文字檔。
預設輸出路徑：%USERPROFILE%\\_tmp\\_tmp.srt

更新紀錄：
 - v0.1.0 2026-09-12: 初版。自 I 欄匯出 SRT，預設使用作用中活頁簿。
"""

# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import os
import sys
from pathlib import Path

import xlwings as xw

from mod_logging import (
    init_logging,
    logging_exc_error,
    logging_exception,
    logging_process_step,
)

# =========================================================================
# 常數定義
# =========================================================================
EXIT_CODE_SUCCESS = 0
EXIT_CODE_NO_FILE = 1
EXIT_CODE_INVALID_INPUT = 2
EXIT_CODE_SAVE_FAILURE = 3
EXIT_CODE_PROCESS_FAILURE = 10
EXIT_CODE_UNKNOWN_ERROR = 99

DEFAULT_SHEET_NAME = "02_計算時長"
HAN_JI_PIAU_IM_COL = 11  # K 欄
START_ROW = 2
HEADER_VALUE = "可匯出SRT"
DEFAULT_OUTPUT_DIR = "output9"
DEFAULT_TEXT_FILE_PATH = Path.home() / "_tmp" / "_tmp.txt"

# =========================================================================
# 設定日誌
# =========================================================================
init_logging()


# =========================================================================
# 作業協助函數
# =========================================================================
def ensure_xlsx_extension(file_name: str) -> str:
    """檔名若無 .xlsx 則補上。"""
    return file_name if file_name.lower().endswith(".xlsx") else f"{file_name}.xlsx"


def resolve_workbook_path(file_arg: str, project_root: Path) -> Path:
    """解析 --file：絕對路徑照用；僅檔名時預設在 output9 尋找。"""
    raw = ensure_xlsx_extension(file_arg.strip())
    path = Path(raw).expanduser()

    if path.is_absolute():
        if path.exists():
            return path
        raise FileNotFoundError(f"找不到指定的 Excel 檔案：{path}")

    candidates = [
        path,
        project_root / path,
        project_root / DEFAULT_OUTPUT_DIR / path.name,
        project_root / DEFAULT_OUTPUT_DIR / path,
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate.resolve()

    searched = project_root / DEFAULT_OUTPUT_DIR / path.name
    raise FileNotFoundError(f"找不到指定的 Excel 檔案：{file_arg}（已於預設目錄尋找：{searched}）")


def resolve_srt_output_path(output_arg: str | None) -> Path:
    """解析輸出 SRT 路徑；未指定時使用 %USERPROFILE%\\_tmp\\_tmp.srt。"""
    if not output_arg:
        return DEFAULT_TEXT_FILE_PATH

    raw = os.path.expandvars(output_arg.strip())
    path = Path(raw).expanduser()
    if path.suffix.lower() != ".srt":
        path = path.with_suffix(".srt") if not path.suffix else path
    return path


def get_workbook(file_arg: str | None, project_root: Path):
    """預設使用作用中活頁簿；僅在指定 --file 時才依 output9 開啟檔案。"""
    if file_arg:
        path = resolve_workbook_path(file_arg, project_root)
        logging_process_step(f"依 --file 開啟活頁簿：{path}")
        return xw.Book(str(path))

    try:
        return xw.Book.caller()
    except Exception:
        pass

    try:
        wb = xw.apps.active.books.active
        if wb:
            logging_process_step(f"使用作用中活頁簿：{wb.fullname}")
            return wb
    except Exception as e:
        logging_exc_error(msg="無法找到作用中的 Excel 工作簿！請先開啟活頁簿，或使用 --file 指定檔名。", error=e)
        return None

    return None


def resolve_srt_sheet(wb, sheet_arg: str | None):
    """決定要讀取 I 欄的工作表。"""
    if sheet_arg:
        sheet_names = [s.name for s in wb.sheets]
        if sheet_arg not in sheet_names:
            raise ValueError(f"活頁簿找不到工作表【{sheet_arg}】，現有工作表：{sheet_names}")
        return wb.sheets[sheet_arg]

    active = wb.sheets.active
    header = active.range((1, HAN_JI_PIAU_IM_COL)).value
    if header and str(header).strip() == HEADER_VALUE:
        return active

    sheet_names = [s.name for s in wb.sheets]
    if DEFAULT_SHEET_NAME in sheet_names:
        return wb.sheets[DEFAULT_SHEET_NAME]

    for sheet in wb.sheets:
        value = sheet.range((1, HAN_JI_PIAU_IM_COL)).value
        if value and str(value).strip() == HEADER_VALUE:
            return sheet

    raise ValueError(f"找不到含【{HEADER_VALUE}】之 I 欄工作表（預設：{DEFAULT_SHEET_NAME}）。")


def collect_srt_blocks(sheet) -> list[str]:
    """自 I 欄第 2 列起讀取 SRT 區塊，遇空列即停。"""
    blocks: list[str] = []
    row = START_ROW
    while True:
        cell_value = sheet.range((row, HAN_JI_PIAU_IM_COL)).value
        if cell_value is None or str(cell_value).strip() == "":
            break
        block = str(cell_value).replace("\r\n", "\n").replace("\r", "\n").strip()
        if block:
            blocks.append(block)
        row += 1
    return blocks


def write_srt_file(blocks: list[str], output_path: Path) -> Path:
    """將 SRT 區塊寫入文字檔。"""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    content = "\n\n".join(blocks)
    if content and not content.endswith("\n"):
        content += "\n"
    output_path.write_text(content, encoding="utf-8", newline="\n")
    return output_path


def process(wb, args) -> int:
    """讀取 I 欄 SRT 並輸出文字檔。"""
    logging_process_step("<=========== 作業開始！==========>")

    try:
        sheet = resolve_srt_sheet(wb, getattr(args, "sheet", None))
        output_path = resolve_srt_output_path(getattr(args, "output", None))

        logging_process_step(f"活頁簿：{wb.fullname}")
        logging_process_step(f"工作表：{sheet.name}")
        logging_process_step(f"來源欄：{xw.utils.col_name(HAN_JI_PIAU_IM_COL)}（自第 {START_ROW} 列起）")
        logging_process_step(f"輸出檔：{output_path}")

        sheet.activate()
        blocks = collect_srt_blocks(sheet)
        if not blocks:
            raise ValueError(f"【{sheet.name}】工作表 I 欄沒有可匯出的 SRT 內容。")

        write_srt_file(blocks, output_path)

        print("=" * 80)
        print(output_path.read_text(encoding="utf-8"), end="")
        print("=" * 80)
        logging_process_step(f"已輸出 {len(blocks)} 則字幕至：{output_path}")
        return EXIT_CODE_SUCCESS

    except Exception as e:
        logging_exception(msg="匯出 SRT 字幕檔時發生例外！", error=e)
        raise


def main(args) -> int:
    """主程式：預設處理作用中活頁簿。"""
    current_file_path = Path(__file__).resolve()
    project_root = current_file_path.parent
    program_name = current_file_path.stem

    logging_process_step(f"《========== 程式開始執行：{program_name} ==========》")
    logging_process_step(f"專案根目錄為: {project_root}")

    try:
        wb = get_workbook(getattr(args, "file", None), project_root)
    except FileNotFoundError as e:
        logging_exc_error(msg=str(e), error=e)
        return EXIT_CODE_NO_FILE
    except Exception as e:
        logging_exc_error(msg="無法開啟 Excel 活頁簿！", error=e)
        return EXIT_CODE_NO_FILE

    if not wb:
        logging_exc_error(msg="無法取得 Excel 活頁簿！請先開啟活頁簿，或使用 --file 指定檔名。", error=None)
        return EXIT_CODE_NO_FILE

    try:
        exit_code = process(wb, args)
    except Exception as e:
        msg = f"作業程序發生異常，終止執行：{program_name}"
        logging_exception(msg=msg, error=e)
        return EXIT_CODE_PROCESS_FAILURE

    if exit_code != EXIT_CODE_SUCCESS:
        msg = f"處理作業發生異常，終止程式執行：{program_name}（處理作業程序，返回失敗碼）"
        logging_exc_error(msg=msg, error=None)
        return EXIT_CODE_PROCESS_FAILURE

    print("\n")
    print("=" * 80)
    logging_process_step(f"《========== 程式終止執行：{program_name} ==========》")
    return EXIT_CODE_SUCCESS


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description="將工作表 I 欄【可匯出SRT】輸出成 SRT 文字檔",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=f"""
使用範例：
  python b200_製作字幕檔.py
      # 預設：作用中活頁簿，輸出至 {DEFAULT_TEXT_FILE_PATH}
  python b200_製作字幕檔.py --file 【字幕】帛書版道德經。第十一章.xlsx
  python b200_製作字幕檔.py --sheet 海海人生
  python b200_製作字幕檔.py --output %USERPROFILE%\\_tmp\\道德經.srt
""",
    )
    parser.add_argument(
        "--file",
        dest="file",
        default=None,
        help=f"活頁簿檔名或路徑；僅檔名時預設於 {DEFAULT_OUTPUT_DIR} 尋找。未指定則使用作用中活頁簿",
    )
    parser.add_argument(
        "--sheet",
        dest="sheet",
        default=None,
        help=f"工作表名稱；未指定時，作用中工作表 I1 若為【{HEADER_VALUE}】則用之，否則用【{DEFAULT_SHEET_NAME}】",
    )
    parser.add_argument(
        "--output",
        dest="output",
        default=None,
        help=f"輸出 SRT 路徑（預設：{DEFAULT_TEXT_FILE_PATH}）",
    )
    args = parser.parse_args()

    exit_code = main(args)
    if exit_code != EXIT_CODE_SUCCESS:
        print(f"程式異常終止，錯誤代碼為: {exit_code}")
        sys.exit(exit_code)
