"""b300_製作上下注音文.py v0.1.2

自【SRT漢字注音】工作表 T 欄（起始儲存格 T4）讀取上下注音文，
輸出成文字檔，並於 Console 顯示處理結果。

換行使用 Windows CRLF（\\r\\n），並將結果放入剪貼簿，方便貼入 PowerPoint。

更新紀錄：
 - v0.1.2 2026-09-13: Console／文字檔改 CRLF；結果同步寫入剪貼簿。
 - v0.1.1 2026-09-13: 改讀作用中活頁簿【SRT漢字注音】T4 起；保留 Console 顯示。
 - v0.1.0 2026-09-12: 初版。
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

DEFAULT_SHEET_NAME = "SRT漢字注音"
SOURCE_COL = 20  # T 欄
START_ROW = 4  # 起始儲存格：T4
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


def resolve_output_path(output_arg: str | None) -> Path:
    """解析輸出文字檔路徑；未指定時使用 %USERPROFILE%\\_tmp\\_tmp.txt。"""
    if not output_arg:
        return DEFAULT_TEXT_FILE_PATH

    raw = os.path.expandvars(output_arg.strip())
    path = Path(raw).expanduser()
    if not path.suffix:
        path = path.with_suffix(".txt")
    return path


def get_workbook(file_arg: str | None, project_root: Path):
    """預設使用作用中活頁簿；僅在指定 --file 時才依路徑開啟檔案。"""
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
        logging_exc_error(msg="無法找到作用中的 Excel 工作簿！請先開啟活頁簿。", error=e)
        return None

    return None


def normalize_lf(text: str) -> str:
    """將各種換行統一成 LF，方便後續再轉成 CRLF。"""
    return str(text).replace("\r\n", "\n").replace("\r", "\n")


def to_crlf(text: str) -> str:
    """轉成 Windows 換行（CRLF），PowerPoint 文字框才會當成段落。"""
    return normalize_lf(text).replace("\n", "\r\n")


def print_crlf(text: str = "", end: str = "\n"):
    """輸出至 Console，換行寫成 CRLF。"""
    sys.stdout.write(to_crlf(f"{text}{end}"))
    sys.stdout.flush()


def copy_to_windows_clipboard(text: str) -> bool:
    """將文字以 CRLF 放入 Windows 剪貼簿。"""
    try:
        import win32clipboard
    except ImportError:
        return False

    payload = to_crlf(text)
    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText(payload, win32clipboard.CF_UNICODETEXT)
    finally:
        win32clipboard.CloseClipboard()
    return True


def resolve_source_sheet(wb, sheet_arg: str | None):
    """取得【SRT漢字注音】工作表。"""
    sheet_name = sheet_arg or DEFAULT_SHEET_NAME
    sheet_names = [sheet.name for sheet in wb.sheets]
    if sheet_name not in sheet_names:
        raise ValueError(f"活頁簿找不到工作表【{sheet_name}】，現有工作表：{sheet_names}")
    return wb.sheets[sheet_name]


def collect_annotation_blocks(sheet) -> list[str]:
    """自 T4 起逐列讀取上下注音文，遇空列即停。"""
    blocks: list[str] = []
    row = START_ROW
    while True:
        cell = sheet.range((row, SOURCE_COL))
        cell_value = cell.value
        addr = f"{xw.utils.col_name(SOURCE_COL)}{row}"
        if cell_value is None or str(cell_value).strip() == "":
            print_crlf(f"{addr}：【空白，停止讀取】")
            break

        block = normalize_lf(cell_value).strip()
        if block:
            blocks.append(block)
            print_crlf("-" * 80)
            print_crlf(f"{addr}：")
            print_crlf(block)
        row += 1
    return blocks


def join_blocks(blocks: list[str]) -> str:
    """以空行分隔各則上下注音文。"""
    content = "\n\n".join(blocks)
    if content and not content.endswith("\n"):
        content += "\n"
    return content


def write_text_file(blocks: list[str], output_path: Path) -> Path:
    """將上下注音文寫入文字檔（Windows CRLF）。"""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(join_blocks(blocks), encoding="utf-8", newline="\r\n")
    return output_path


def process(wb, args) -> int:
    """讀取 T 欄上下注音文並輸出文字檔，同時於 Console 顯示。"""
    logging_process_step("<=========== 作業開始！==========>")

    try:
        sheet = resolve_source_sheet(wb, getattr(args, "sheet", None))
        output_path = resolve_output_path(getattr(args, "output", None))

        logging_process_step(f"活頁簿：{wb.fullname}")
        logging_process_step(f"工作表：{sheet.name}")
        logging_process_step(f"來源儲存格：{xw.utils.col_name(SOURCE_COL)}{START_ROW} 起")
        logging_process_step(f"輸出檔：{output_path}")

        sheet.activate()
        blocks = collect_annotation_blocks(sheet)
        if not blocks:
            raise ValueError(f"【{sheet.name}】工作表 T 欄自 T{START_ROW} 起沒有可匯出的上下注音文。")

        write_text_file(blocks, output_path)
        content = join_blocks(blocks)

        print_crlf("=" * 80)
        print_crlf(content, end="")
        print_crlf("=" * 80)
        if copy_to_windows_clipboard(content):
            logging_process_step("已將上下注音文（CRLF）複製到剪貼簿，可直接貼入 PowerPoint。")
        else:
            logging_process_step("無法寫入剪貼簿；請改從輸出文字檔複製。")
        logging_process_step(f"已輸出 {len(blocks)} 則上下注音文至：{output_path}")
        return EXIT_CODE_SUCCESS

    except Exception as e:
        logging_exception(msg="匯出上下注音文時發生例外！", error=e)
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
        logging_exc_error(msg="無法取得 Excel 活頁簿！請先開啟活頁簿。", error=None)
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
        description="自【SRT漢字注音】工作表 T4 起讀取上下注音文並輸出文字檔",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=f"""
使用範例：
  python b300_製作上下注音文.py
      # 預設：作用中活頁簿【SRT漢字注音】T4 起，輸出至 {DEFAULT_TEXT_FILE_PATH}
  python b300_製作上下注音文.py --output %USERPROFILE%\\_tmp\\上下注音文.txt
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
        help=f"工作表名稱；未指定時使用【{DEFAULT_SHEET_NAME}】",
    )
    parser.add_argument(
        "--output",
        dest="output",
        default=None,
        help=f"輸出文字檔路徑（預設：{DEFAULT_TEXT_FILE_PATH}）",
    )
    args = parser.parse_args()

    exit_code = main(args)
    if exit_code != EXIT_CODE_SUCCESS:
        print(f"程式異常終止，錯誤代碼為: {exit_code}")
        sys.exit(exit_code)
