"""b000_抄寫到SRT漢字注音工作表.py v0.1.1

自【漢字注音】工作表讀取各漢字單元的【漢字列】與【漢字標音列】，
依斷句規則抄入【SRT漢字注音】工作表 Q 欄（漢字）與 R 欄（漢字標音）。

斷句規則（於漢字列判斷）：
 - 讀到換行（\\n 或 =CHAR(10)）時，寫入目前句子並換到下一列。
 - 讀到【。】緊接換行時，將【。】併入目前句子後換到下一列。
 - 讀到【φ】時，寫出殘句後結束。

同一漢字行滿 15 字而尚未遇斷句符時，繼續累積到下一漢字行，不因此斷句。

標音列格式對齊【SRT漢字注音（台羅拼音）】：
 - 句首（及 . ! ? 之後）首字母大寫，並保留聲調符號。
 - 標點轉半形後緊接前一音節（至少 , . ? !）。

更新紀錄：
 - v0.1.1 2026-09-13: 句首大寫保留調符；標點緊接前一音節。
 - v0.1.0 2026-09-13: 初版。預設處理作用中活頁簿。
"""

# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import sys
import unicodedata
from pathlib import Path

import xlwings as xw

from mod_excel_access import (
    END_COL,
    HAN_JI_OFFSET,
    HAN_JI_PIAU_IM_OFFSET,
    ROWS_PER_LINE,
    START_COL,
    START_ROW_NO,
    ensure_sheet_exists,
    get_value_by_name,
)
from mod_logging import (
    init_logging,
    logging_exc_error,
    logging_exception,
    logging_process_step,
)
from mod_標音 import is_han_ji

# =========================================================================
# 常數定義
# =========================================================================
EXIT_CODE_SUCCESS = 0
EXIT_CODE_NO_FILE = 1
EXIT_CODE_INVALID_INPUT = 2
EXIT_CODE_SAVE_FAILURE = 3
EXIT_CODE_PROCESS_FAILURE = 10
EXIT_CODE_UNKNOWN_ERROR = 99

SOURCE_SHEET_NAME = "漢字注音"
TARGET_SHEET_NAME = "SRT漢字注音"
HAN_JI_OUT_COL = 17  # Q 欄
PIAU_IM_OUT_COL = 18  # R 欄
OUTPUT_START_ROW = 4  # 起始儲存格：Q4、R4
OUTPUT_HEADER_ROW = 3  # 對齊【SRT漢字注音（台羅拼音）】Q3／R3 表頭
OUTPUT_HAN_JI_HEADER = "漢字"
OUTPUT_PIAU_IM_HEADER = "漢字標音"

# 標點轉半形，緊接前一個標音（對齊【SRT漢字注音（台羅拼音）】R 欄）
PUNCTUATION_MAP = {
    "，": ",",
    "。": ".",
    "！": "!",
    "？": "?",
    "；": ";",
    "：": ":",
    "、": ",",
    ",": ",",
    ".": ".",
    "!": "!",
    "?": "?",
    ";": ";",
    ":": ":",
    "—": "—",
    "…": "…",
}
SKIP_PUNCTUATION = set("《》〈〉「」『』（）()　")
SENTENCE_END_PUNCTUATIONS = {".", "!", "?"}

DEFAULT_TOTAL_LINES = 120
DEFAULT_CHARS_PER_ROW = 15
DEFAULT_OUTPUT_DIR = "output9"
END_MARK = "φ"
PERIOD_MARK = "。"
NEWLINE_VALUES = {"\n", "\r", "\r\n", chr(10), chr(13)}

# =========================================================================
# 設定日誌
# =========================================================================
init_logging()


def capitalize_im_piau(im_piau: str) -> str:
    """將音標首字母改大寫，保留聲調符號（含預組字與組合字元）。"""
    if not im_piau:
        return im_piau

    text = unicodedata.normalize("NFC", im_piau)
    for index, char in enumerate(text):
        category = unicodedata.category(char)
        if char.isalpha() or category.startswith("L"):
            return text[:index] + char.upper() + text[index + 1 :]
    return text


class PiauImLineBuilder:
    """組建標音行：音節空白分隔、標點緊接、句首大寫（對齊參考工作表）。"""

    def __init__(self):
        self.tokens: list[str] = []
        self.capitalize_next = True

    def reset(self):
        self.tokens.clear()
        self.capitalize_next = True

    def add_piau_im(self, piau_im: str):
        piau_im = str(piau_im).strip()
        if not piau_im:
            return
        if self.capitalize_next:
            piau_im = capitalize_im_piau(piau_im)
            self.capitalize_next = False
        self.tokens.append(piau_im)

    def add_punctuation(self, han_ji_punct: str):
        if han_ji_punct in SKIP_PUNCTUATION:
            return
        punct = PUNCTUATION_MAP.get(han_ji_punct, han_ji_punct)
        if not punct:
            return
        if self.tokens:
            self.tokens[-1] += punct
        else:
            self.tokens.append(punct)
        if punct in SENTENCE_END_PUNCTUATIONS:
            self.capitalize_next = True

    def build(self) -> str:
        return " ".join(self.tokens)


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
        logging_exc_error(msg="無法找到作用中的 Excel 工作簿！請先開啟活頁簿，或使用 --file 指定檔名。", error=e)
        return None

    return None


def _int_named_value(wb, name: str, default: int) -> int:
    """讀取活頁簿名稱為整數；缺值時回傳 default。"""
    value = get_value_by_name(wb, name)
    if value is None or str(value).strip() == "":
        return default
    return int(value)


def _as_row_values(raw, count: int) -> list:
    """將 xlwings 列範圍讀值正規成一維清單。"""
    if raw is None:
        return [None] * count
    if not isinstance(raw, (list, tuple)):
        return [raw]
    if raw and isinstance(raw[0], (list, tuple)):
        return list(raw[0])
    return list(raw)


def is_newline(value, formula=None) -> bool:
    """判斷漢字列儲存格是否為換行控制碼（\\n 或 =CHAR(10)）。"""
    if formula is not None and "=CHAR(10)" in str(formula).upper():
        return True
    if value is None:
        return False
    if value in (10, 13):
        return True
    if isinstance(value, str) and value in NEWLINE_VALUES:
        return True
    return False


def is_end_mark(value) -> bool:
    """判斷是否為文章終止符號 φ。"""
    return value is not None and str(value).strip() == END_MARK


def is_period(value) -> bool:
    """判斷是否為句號【。】。"""
    return value is not None and str(value).strip() == PERIOD_MARK


def is_period_plus_newline(value) -> bool:
    """單一儲存格同時含【。】與換行時亦視為斷句。"""
    if value is None:
        return False
    text = str(value).replace("\r\n", "\n").replace("\r", "\n")
    return text == f"{PERIOD_MARK}\n"


def cell_text(value) -> str:
    """將儲存格值轉成字串；空白則回傳空字串。"""
    if value is None:
        return ""
    return str(value)


def collect_source_cells(sheet, total_lines: int, start_col: int, end_col: int) -> list[dict]:
    """依漢字行列掃描來源儲存格（含公式，供偵測 CHAR(10)）。"""
    cells: list[dict] = []
    chars_per_row = end_col - start_col + 1

    for line_no in range(1, total_lines + 1):
        base_row = START_ROW_NO + (line_no - 1) * ROWS_PER_LINE
        han_ji_row = base_row + HAN_JI_OFFSET
        piau_im_row = base_row + HAN_JI_PIAU_IM_OFFSET
        han_range = sheet.range((han_ji_row, start_col), (han_ji_row, end_col))
        piau_range = sheet.range((piau_im_row, start_col), (piau_im_row, end_col))
        han_values = _as_row_values(han_range.value, chars_per_row)
        han_formulas = _as_row_values(han_range.formula, chars_per_row)
        piau_values = _as_row_values(piau_range.value, chars_per_row)

        for offset, col in enumerate(range(start_col, end_col + 1)):
            cells.append(
                {
                    "line_no": line_no,
                    "han_ji_row": han_ji_row,
                    "col": col,
                    "han_ji": han_values[offset] if offset < len(han_values) else None,
                    "han_formula": han_formulas[offset] if offset < len(han_formulas) else None,
                    "piau_im": piau_values[offset] if offset < len(piau_values) else None,
                }
            )
    return cells


def build_sentences(source_cells: list[dict]) -> list[tuple[str, str]]:
    """依斷句規則將來源儲存格組成（漢字, 漢字標音）列。"""
    sentences: list[tuple[str, str]] = []
    han_chars: list[str] = []
    piau_line = PiauImLineBuilder()

    def flush():
        han_text = "".join(han_chars).strip()
        piau_text = piau_line.build().strip()
        han_chars.clear()
        piau_line.reset()
        if not han_text and not piau_text:
            return
        sentences.append((han_text, piau_text))

    def append_pair(han_ji, piau_im):
        han_text = cell_text(han_ji)
        if not han_text:
            return
        han_chars.append(han_text)
        if len(han_text) == 1 and not is_han_ji(han_text):
            piau_line.add_punctuation(han_text)
            return
        piau_text = cell_text(piau_im).strip()
        if piau_text:
            piau_line.add_piau_im(piau_text)

    index = 0
    total = len(source_cells)
    while index < total:
        cell = source_cells[index]
        han_ji = cell["han_ji"]
        formula = cell["han_formula"]
        addr = f"{xw.utils.col_name(cell['col'])}{cell['han_ji_row']}"

        if is_end_mark(han_ji):
            print(f"{addr} [漢字列] = 【文章終止 φ】")
            flush()
            break

        if is_period_plus_newline(han_ji):
            append_pair(PERIOD_MARK, cell["piau_im"])
            print(f"{addr} [漢字列] = 【句號＋換行】")
            flush()
            index += 1
            continue

        if is_newline(han_ji, formula):
            print(f"{addr} [漢字列] = 【換行】")
            flush()
            index += 1
            continue

        if han_ji is None or cell_text(han_ji).strip() == "":
            index += 1
            continue

        append_pair(han_ji, cell["piau_im"])

        if is_period(han_ji):
            next_cell = source_cells[index + 1] if index + 1 < total else None
            if next_cell and is_newline(next_cell["han_ji"], next_cell["han_formula"]):
                next_addr = f"{xw.utils.col_name(next_cell['col'])}{next_cell['han_ji_row']}"
                print(f"{addr}+{next_addr} [漢字列] = 【。】＋【換行】")
                flush()
                index += 2
                continue

        index += 1

    flush()
    return sentences


def clear_output_columns(sheet, start_row: int):
    """清除目標工作表既有的 Q／R 欄抄寫結果。"""
    last_row = max(sheet.used_range.last_cell.row, start_row)
    if last_row >= start_row:
        sheet.range((start_row, HAN_JI_OUT_COL), (last_row, PIAU_IM_OUT_COL)).clear_contents()


def write_sentences(sheet, sentences: list[tuple[str, str]]) -> int:
    """將句子寫入 Q／R 欄，並補上表頭。"""
    header_han = sheet.range((OUTPUT_HEADER_ROW, HAN_JI_OUT_COL)).value
    header_piau = sheet.range((OUTPUT_HEADER_ROW, PIAU_IM_OUT_COL)).value
    if header_han is None or str(header_han).strip() == "":
        sheet.range((OUTPUT_HEADER_ROW, HAN_JI_OUT_COL)).value = OUTPUT_HAN_JI_HEADER
    if header_piau is None or str(header_piau).strip() == "":
        sheet.range((OUTPUT_HEADER_ROW, PIAU_IM_OUT_COL)).value = OUTPUT_PIAU_IM_HEADER

    clear_output_columns(sheet, OUTPUT_START_ROW)
    if not sentences:
        return 0

    end_row = OUTPUT_START_ROW + len(sentences) - 1
    sheet.range((OUTPUT_START_ROW, HAN_JI_OUT_COL), (end_row, PIAU_IM_OUT_COL)).value = sentences
    return len(sentences)


def resolve_source_sheet(wb):
    """取得【漢字注音】工作表。"""
    sheet_names = [sheet.name for sheet in wb.sheets]
    if SOURCE_SHEET_NAME not in sheet_names:
        raise ValueError(f"活頁簿找不到工作表【{SOURCE_SHEET_NAME}】，現有工作表：{sheet_names}")
    return wb.sheets[SOURCE_SHEET_NAME]


# =========================================================================
# 本程式主要處理作業程序
# =========================================================================
def process(wb, _args) -> int:
    """自【漢字注音】抄寫漢字與標音至【SRT漢字注音】Q／R 欄。"""
    logging_process_step("<=========== 作業開始！==========>")

    try:
        source_sheet = resolve_source_sheet(wb)
        target_sheet = ensure_sheet_exists(wb, TARGET_SHEET_NAME)
        if target_sheet is None:
            raise ValueError(f"無法建立或開啟工作表【{TARGET_SHEET_NAME}】。")

        total_lines = _int_named_value(wb, "每頁總列數", DEFAULT_TOTAL_LINES)
        chars_per_row = _int_named_value(wb, "每列總字數", DEFAULT_CHARS_PER_ROW)
        start_col = START_COL
        end_col = min(START_COL + chars_per_row - 1, END_COL)

        logging_process_step(f"活頁簿：{wb.fullname}")
        logging_process_step(f"來源工作表：{source_sheet.name}")
        logging_process_step(f"目標工作表：{target_sheet.name}")
        logging_process_step(f"掃描範圍：漢字列 {start_col}–{end_col} 欄，共 {total_lines} 行")

        source_sheet.activate()
        source_cells = collect_source_cells(source_sheet, total_lines, start_col, end_col)
        sentences = build_sentences(source_cells)
        written = write_sentences(target_sheet, sentences)
        target_sheet.activate()

        print("=" * 80)
        for offset, (han_text, piau_text) in enumerate(sentences):
            row = OUTPUT_START_ROW + offset
            print(f"Q{row}：{han_text}")
            print(f"R{row}：{piau_text}")
        print("=" * 80)
        logging_process_step(f"已抄寫 {written} 列至【{TARGET_SHEET_NAME}】Q／R 欄。")
        return EXIT_CODE_SUCCESS

    except Exception as e:
        logging_exception(msg="抄寫至【SRT漢字注音】工作表時發生例外！", error=e)
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

    try:
        wb.save()
        logging_process_step(f"已儲存檔案：{wb.fullname}")
    except Exception as e:
        logging_exception(msg="儲存檔案失敗！", error=e)
        return EXIT_CODE_SAVE_FAILURE

    print("\n")
    print("=" * 80)
    logging_process_step(f"《========== 程式終止執行：{program_name} ==========》")
    return EXIT_CODE_SUCCESS


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description="將【漢字注音】的漢字列與漢字標音列抄入【SRT漢字注音】Q／R 欄",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用範例：
  python b000_抄寫到SRT漢字注音工作表.py
      # 預設：處理 Excel 作用中活頁簿（請先開啟檔案）
  python b000_抄寫到SRT漢字注音工作表.py --file 【字幕】帛書版道德經。第十一章.xlsx
      # 僅檔名時，預設於 output9 目錄尋找
""",
    )
    parser.add_argument(
        "--file",
        dest="file",
        default=None,
        help=f"活頁簿檔名或路徑；僅檔名時預設於 {DEFAULT_OUTPUT_DIR} 尋找。未指定則使用作用中活頁簿",
    )
    args = parser.parse_args()

    exit_code = main(args)
    if exit_code != EXIT_CODE_SUCCESS:
        print(f"程式異常終止，錯誤代碼為: {exit_code}")
        sys.exit(exit_code)
