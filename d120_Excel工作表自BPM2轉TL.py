""" d120_Excel工作表自BPM2轉TL.py v0.0.1
摘要：
    將 Excel 檔案中的【台語注音二式】工作表反轉換成【台羅拼音】工作表。
    1. 自【台語注音二式】工作表複製成【台羅拼音】工作表。
    2. 漢字標音欄預設依表頭自動判斷（code 等）；可用 --col_addr 指定。
    3. 將標音欄由【台語注音二式】轉回【台羅拼音】後寫回標的工作表。
    4. 作業對象：source_dir_path 下的【台語注音二式字庫】.xlsx（單一活頁簿）。

指令：
    python d120_Excel工作表自BPM2轉TL.py --col_addr=[欄位址]

使用方法：
    1. 來源檔目錄路徑： %UserProfile%\\work\\rime-tlpa\\src
    2. Excel 檔案：【台語注音二式字庫】.xlsx
    3. 來源工作表名稱：台語注音二式
    4. 標的工作表名稱：台羅拼音
    5. 轉換路徑：台語注音二式（MPS2／BPM2）→ TLPA → 台羅拼音（TL）
"""  # noqa: N999

from __future__ import annotations

import argparse
import logging
import re
import sys
import time
from pathlib import Path

import xlwings as xw

from d000_定義檔 import DEFAULT_CODE_COL, expand_dir_path, source_dir_path
from mod_convert_TLPA_to_MPS2 import convert_MPS2_to_TLPA
from mod_標音 import convert_tlpa_to_tl

# =========================================================================
# 設定日誌
# =========================================================================
logging.basicConfig(
    filename="process_log.txt",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)

EXIT_CODE_SUCCESS = 0
EXIT_CODE_FAILURE = 1

WORKBOOK_NAME = "【台語注音二式字庫】.xlsx"
SOURCE_SHEET = "台語注音二式"
TARGET_SHEET = "台羅拼音"

PHONETIC_HEADER_KEYS = (
    "code",
    "台語注音二式",
    "注音二式",
    "台羅音標",
    "台羅拼音",
    "拼音",
    "音標",
    "標音",
)


def say(msg: str = "", *, end: str = "\n") -> None:
    """立即輸出到 Console（flush），避免長時間無回應像當掉。"""
    print(msg, end=end, flush=True)


def fmt_elapsed(seconds: float) -> str:
    if seconds < 60:
        return f"{seconds:.1f}s"
    mins, secs = divmod(int(seconds), 60)
    return f"{mins}m{secs:02d}s"


def parse_col_addr(col_addr: str) -> int:
    """將欄位址轉成 1-based 欄號。支援：B、B:B、2、B2 等。"""
    raw = str(col_addr).strip().upper()
    if not raw:
        raise ValueError("欄位址不可為空白")

    if raw.isdigit():
        col = int(raw)
        if col < 1:
            raise ValueError(f"欄號必須 >= 1：{col_addr}")
        return col

    match = re.match(r"^([A-Z]+)", raw)
    if not match:
        raise ValueError(f"無法辨識欄位址：{col_addr}")

    letters = match.group(1)
    col = 0
    for ch in letters:
        col = col * 26 + (ord(ch) - ord("A") + 1)
    return col


def col_num_to_letter(col: int) -> str:
    """1-based 欄號 → Excel 欄字母（如 3 → C）。"""
    letters = []
    n = col
    while n > 0:
        n, rem = divmod(n - 1, 26)
        letters.append(chr(ord("A") + rem))
    return "".join(reversed(letters))


def detect_phonetic_col(sheet: xw.Sheet) -> tuple[int, str]:
    """依第 1 列表頭自動找出標音欄；找不到則退回預設 B 欄。"""
    header = sheet.range("A1:Z1").value
    if header is None:
        return parse_col_addr(DEFAULT_CODE_COL), f"預設 {DEFAULT_CODE_COL}"

    if not isinstance(header, (list, tuple)):
        header = [header]

    for idx, cell in enumerate(header, start=1):
        if cell is None:
            continue
        name = str(cell).strip().lower()
        for key in PHONETIC_HEADER_KEYS:
            if key.lower() == name or key.lower() in name:
                return idx, str(cell).strip()

    return parse_col_addr(DEFAULT_CODE_COL), f"預設 {DEFAULT_CODE_COL}（表頭未匹配）"


def resolve_code_col(sheet: xw.Sheet, col_addr_override: str | None) -> tuple[int, str]:
    """決定標音欄：--col_addr > 表頭自動判斷。"""
    if col_addr_override:
        col = parse_col_addr(col_addr_override)
        return col, f"參數 --col_addr={col_addr_override}"

    col, label = detect_phonetic_col(sheet)
    return col, f"表頭自動判斷：{label}（{col_num_to_letter(col)}）"


def convert_bpm2_code_to_tl(code: str) -> str:
    """
    將 code 欄之台語注音二式轉成台羅拼音。
    多音節以空白分隔，逐音節轉換後再以空白接回。
    轉換路徑：MPS2／BPM2 → TLPA → 台羅拼音（TL）。
    """
    if not code:
        return ""

    syllables = str(code).strip().split()
    converted = []
    for syl in syllables:
        if syl.startswith("#"):
            converted.append(syl)
            continue
        body = syl.lower()
        tlpa = convert_MPS2_to_TLPA(body)
        converted.append(convert_tlpa_to_tl(tlpa))
    return " ".join(converted)


def get_workbook(workbook_path: Path):
    """取得目標活頁簿：若已在 Excel 開啟則直接使用，否則開啟檔案。"""
    workbook_name = workbook_path.name
    for app in xw.apps:
        for book in app.books:
            if book.name == workbook_name or Path(book.fullname or "").name == workbook_name:
                say(f"📌 使用已開啟之活頁簿：{book.name}")
                return book, False

    if not workbook_path.exists():
        raise FileNotFoundError(f"找不到活頁簿檔案：{workbook_path}")

    size_mb = workbook_path.stat().st_size / (1024 * 1024)
    say(f"📌 正在開啟活頁簿（約 {size_mb:.1f} MB，大檔可能需數十秒）…")
    say(f"   {workbook_path}")
    t0 = time.perf_counter()
    wb = xw.Book(str(workbook_path))
    say(f"✅ 活頁簿已開啟（耗時 {fmt_elapsed(time.perf_counter() - t0)}）")
    return wb, True


def copy_source_to_target(wb, source_sheet: str, target_sheet: str) -> xw.Sheet:
    """將來源工作表複製為標的工作表；若標的已存在則先刪除。"""
    sheet_names = [s.name for s in wb.sheets]
    if source_sheet not in sheet_names:
        raise ValueError(f"找不到來源工作表：{source_sheet}")

    if target_sheet in sheet_names:
        say(f"⚠️ 標的工作表【{target_sheet}】已存在，正在刪除…")
        t0 = time.perf_counter()
        wb.sheets[target_sheet].delete()
        say(f"   已刪除（耗時 {fmt_elapsed(time.perf_counter() - t0)}）")

    say(f"⏳ 正在複製工作表【{source_sheet}】→【{target_sheet}】…")
    t0 = time.perf_counter()
    source = wb.sheets[source_sheet]
    source.copy(after=source, name=target_sheet)
    target = wb.sheets[target_sheet]
    say(f"✅ 複製完成（耗時 {fmt_elapsed(time.perf_counter() - t0)}）")
    return target


def convert_code_column(target: xw.Sheet, code_col: int) -> int:
    """將標的工作表指定欄之台語注音二式轉成台羅拼音並寫回。"""
    say("⏳ 正在掃描資料列數…")
    t0 = time.perf_counter()
    last_row = target.range("A" + str(target.cells.last_cell.row)).end("up").row
    say(f"   資料列至第 {last_row} 列（掃描耗時 {fmt_elapsed(time.perf_counter() - t0)}）")
    if last_row < 2:
        say("⚠️ 標的工作表無資料列可轉換。")
        return 0

    total = last_row - 1
    say(f"⏳ 正在設定文字格式並讀取標音欄（共 {total} 列）…")
    t0 = time.perf_counter()
    target.range((2, code_col), (last_row, code_col)).number_format = "@"
    codes = target.range((2, code_col), (last_row, code_col)).value
    if last_row == 2:
        codes = [codes]
    say(f"✅ 讀取完成（耗時 {fmt_elapsed(time.perf_counter() - t0)}）")

    say(f"⏳ 開始反轉換拼音（共 {total} 列）…")
    t0 = time.perf_counter()
    new_codes = []
    converted_count = 0
    processed = 0
    next_pct_mark = 5
    for idx, code in enumerate(codes, start=2):
        processed += 1
        if code is None or str(code).strip() == "":
            new_codes.append(code)
        elif str(code).strip().startswith("##"):
            new_codes.append(code)
        else:
            tl = convert_bpm2_code_to_tl(code)
            new_codes.append(tl)
            converted_count += 1
            if converted_count <= 5:
                say(f"   範例 ({idx}) {code} → {tl}")

        pct = processed * 100 / total if total else 100
        if pct >= next_pct_mark or processed == total:
            elapsed = time.perf_counter() - t0
            rate = processed / elapsed if elapsed > 0 else 0
            remain = (total - processed) / rate if rate > 0 else 0
            say(
                f"\r   轉換進度 {processed}/{total}（{pct:5.1f}%）"
                f" 已轉 {converted_count} 筆"
                f" 耗時 {fmt_elapsed(elapsed)}"
                f" 預估剩餘 {fmt_elapsed(remain)}   ",
                end="",
            )
            while next_pct_mark <= pct:
                next_pct_mark += 5

    say()
    say(f"✅ 拼音轉換完成：有效 {converted_count} 筆（耗時 {fmt_elapsed(time.perf_counter() - t0)}）")

    say(f"⏳ 正在寫回 Excel 標音欄（{total} 列，COM 寫入可能需較久）…")
    t0 = time.perf_counter()
    target.range((2, code_col)).options(transpose=True).value = new_codes
    say(f"✅ 寫回完成（耗時 {fmt_elapsed(time.perf_counter() - t0)}）")
    return converted_count


def process(col_addr_override: str | None) -> int:
    workbook_path = expand_dir_path(source_dir_path) / WORKBOOK_NAME
    say(f"Excel：{WORKBOOK_NAME}")
    say(f"來源：{SOURCE_SHEET} → 標的：{TARGET_SHEET}")
    item_t0 = time.perf_counter()

    wb = None
    opened_by_script = False
    try:
        wb, opened_by_script = get_workbook(workbook_path)
        target = copy_source_to_target(wb, SOURCE_SHEET, TARGET_SHEET)

        code_col, col_desc = resolve_code_col(target, col_addr_override)
        say(f"標音欄：第 {code_col} 欄（{col_num_to_letter(code_col)}）← {col_desc}")

        count = convert_code_column(target, code_col)

        say("⏳ 正在儲存活頁簿（大檔可能需數十秒）…")
        t0 = time.perf_counter()
        wb.save()
        say(f"✅ 已儲存（耗時 {fmt_elapsed(time.perf_counter() - t0)}）：{wb.fullname or wb.name}")
        say(f"✅ 轉換完成：共 {count} 列，合計耗時 {fmt_elapsed(time.perf_counter() - item_t0)}")
        logging.info(
            "d120 轉換完成：%s 列 → %s（col=%s）",
            count,
            TARGET_SHEET,
            col_num_to_letter(code_col),
        )
        return EXIT_CODE_SUCCESS
    except Exception as e:
        say(f"❌ 作業失敗：{e}")
        logging.error("d120 作業失敗：%s", e, exc_info=True)
        return EXIT_CODE_FAILURE
    finally:
        if wb is not None and opened_by_script:
            say("⏳ 正在關閉活頁簿…")
            wb.close()
            say("✅ 活頁簿已關閉")


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="將【台語注音二式字庫】之【台語注音二式】工作表反轉換成【台羅拼音】工作表"
    )
    parser.add_argument(
        "--col_addr",
        default=None,
        help=f"標音欄位址（未指定則依表頭自動判斷，或退回 {DEFAULT_CODE_COL}）",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    say("<=========== d120 作業開始 ===========>")
    say("提示：開啟／複製／寫回／存檔階段若暫停數十秒屬正常，請留意進度訊息。")
    if args.col_addr:
        say(f"標音欄覆寫：{args.col_addr}")
    result = process(args.col_addr)
    say("<=========== d120 作業結束 ===========>")
    return result


if __name__ == "__main__":
    sys.exit(main())
