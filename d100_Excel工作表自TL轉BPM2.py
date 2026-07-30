""" d100_Excel工作表自TL轉BPM2.py v0.0.1
摘要：
    將 Excel 檔案中的【台羅拼音】工作表轉換成【台語注音二式】工作表。
    1. 自【台羅拼音】工作表複製成【台語注音二式】工作表。
    2. 【台羅拼音】工作表，漢字標音欄預設【欄位址】為：B Column，可用 --col_addr 參數指定其他欄位址。
    3. 將【台羅拼音工作表】的漢字標音欄位，轉換成【台語注音二式】，置入【台語注音二式】工作表的漢字標音欄位。
    4. 依 d000_定義檔.Dict_List 逐一處理全部 Excel 活頁簿。

指令：
    python d100_Excel工作表自TL轉BPM2.py --col_addr=[欄位址]

使用方法：
    1. 來源檔目錄路徑： %UserProfile%\\work\\rime-tlpa\\src
    2. 標的檔目錄路徑： %UserProfile%\\work\\rime-tlpa\\src
    3. Excel 檔案：Dict_List 每一筆的 WorkBook檔案
    4. 來源工作表名稱：Dict_List 每一筆的 WorkSheet名稱（如：台羅拼音）
    5. 標的工作表名稱：Dict_List 的標的工作表名稱（預設：台語注音二式）
"""  # noqa: N999

from __future__ import annotations

import argparse
import logging
import re
import sys
from pathlib import Path

import xlwings as xw

from d000_定義檔 import (
    Dict_List,
    expand_dir_path,
    get_code_col_addr,
    get_source_sheet_name,
    get_target_sheet_name,
    source_dir_path,
)
from mod_convert_TLPA_to_MPS2 import convert_TLPA_to_MPS2
from mod_標音 import convert_tl_to_tlpa

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
DEFAULT_COL_ADDR = "B"


def parse_col_addr(col_addr: str) -> int:
    """
    將欄位址轉成 1-based 欄號。
    支援：B、B:B、2、B2 等形式。
    """
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


def convert_tl_code_to_bpm2(code: str) -> str:
    """
    將 code 欄之台羅拼音轉成台語注音二式。
    多音節以空白分隔，逐音節轉換後再以空白接回。
    轉換路徑：台羅 → TLPA → 台語注音二式（MPS2／BPM2）。
    """
    if not code:
        return ""

    syllables = str(code).strip().split()
    converted = []
    for syl in syllables:
        # 註解列（## ...）原樣保留
        if syl.startswith("#"):
            converted.append(syl)
            continue
        tlpa = convert_tl_to_tlpa(syl.lower()) or syl.lower()
        converted.append(convert_TLPA_to_MPS2(tlpa))
    return " ".join(converted)


def get_workbook(workbook_path: Path):
    """取得目標活頁簿：若已在 Excel 開啟則直接使用，否則開啟檔案。"""
    workbook_name = workbook_path.name
    for app in xw.apps:
        for book in app.books:
            if book.name == workbook_name or Path(book.fullname or "").name == workbook_name:
                print(f"📌 使用已開啟之活頁簿：{book.name}")
                return book, False

    if not workbook_path.exists():
        raise FileNotFoundError(f"找不到活頁簿檔案：{workbook_path}")

    print(f"📌 開啟活頁簿：{workbook_path}")
    return xw.Book(str(workbook_path)), True


def copy_source_to_target(wb, source_sheet: str, target_sheet: str) -> xw.Sheet:
    """將來源工作表複製為標的工作表；若標的已存在則先刪除。"""
    sheet_names = [s.name for s in wb.sheets]
    if source_sheet not in sheet_names:
        raise ValueError(f"找不到來源工作表：{source_sheet}")

    if target_sheet in sheet_names:
        print(f"⚠️ 標的工作表【{target_sheet}】已存在，將先刪除再重新複製。")
        wb.sheets[target_sheet].delete()

    source = wb.sheets[source_sheet]
    source.copy(after=source, name=target_sheet)
    target = wb.sheets[target_sheet]
    print(f"✅ 已將【{source_sheet}】複製為【{target_sheet}】")
    return target


def convert_code_column(target: xw.Sheet, code_col: int) -> int:
    """將標的工作表指定欄之台羅拼音轉成台語注音二式並寫回。"""
    last_row = target.range("A" + str(target.cells.last_cell.row)).end("up").row
    if last_row < 2:
        print("⚠️ 標的工作表無資料列可轉換。")
        return 0

    target.range((2, code_col), (last_row, code_col)).number_format = "@"

    codes = target.range((2, code_col), (last_row, code_col)).value
    if last_row == 2:
        codes = [codes]

    new_codes = []
    converted_count = 0
    for idx, code in enumerate(codes, start=2):
        if code is None or str(code).strip() == "":
            new_codes.append(code)
            continue
        # 整格以 ## 開頭者視為註解／保留列
        if str(code).strip().startswith("##"):
            new_codes.append(code)
            continue
        bpm2 = convert_tl_code_to_bpm2(code)
        new_codes.append(bpm2)
        converted_count += 1
        if converted_count <= 10 or converted_count % 1000 == 0:
            print(f"  ({idx}) {code} → {bpm2}")

    target.range((2, code_col)).options(transpose=True).value = new_codes
    return converted_count


def process_one(dict_key: str, dict_cfg: dict, col_addr_override: str | None) -> None:
    """處理 Dict_List 中的單一字典項目。"""
    workbook_path = expand_dir_path(source_dir_path) / dict_cfg["WorkBook檔案"]
    source_sheet = get_source_sheet_name(dict_cfg)
    target_sheet = get_target_sheet_name(dict_cfg)
    col_addr = col_addr_override or get_code_col_addr(dict_cfg)
    code_col = parse_col_addr(col_addr)

    print(f"\n----- [{dict_key}] {dict_cfg['WorkBook檔案']} -----")
    print(f"來源：{source_sheet} → 標的：{target_sheet}；標音欄：{col_addr}")

    wb = None
    opened_by_script = False
    try:
        wb, opened_by_script = get_workbook(workbook_path)
        target = copy_source_to_target(wb, source_sheet, target_sheet)
        count = convert_code_column(target, code_col)
        wb.save()
        print(f"✅ [{dict_key}] 轉換完成：共 {count} 列。已儲存：{wb.fullname or wb.name}")
        logging.info(
            "d100 [%s] 轉換完成：%s 列 → %s（col=%s）",
            dict_key,
            count,
            target_sheet,
            col_addr,
        )
    finally:
        if wb is not None and opened_by_script:
            wb.close()


def process_all(col_addr_override: str | None) -> int:
    failures: list[str] = []
    for dict_key, dict_cfg in Dict_List.items():
        try:
            process_one(dict_key, dict_cfg, col_addr_override)
        except Exception as e:
            failures.append(dict_key)
            print(f"❌ [{dict_key}] 作業失敗：{e}")
            logging.error("d100 [%s] 作業失敗：%s", dict_key, e, exc_info=True)

    if failures:
        print(f"\n⚠️ 失敗項目：{', '.join(failures)}")
        return EXIT_CODE_FAILURE
    return EXIT_CODE_SUCCESS


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="批次將 Dict_List 各 Excel【台羅拼音】工作表轉成【台語注音二式】工作表"
    )
    parser.add_argument(
        "--col_addr",
        default=None,
        help=f"全域覆寫標音欄位址（未指定則用各筆設定或預設 {DEFAULT_COL_ADDR}）",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    print("<=========== d100 作業開始 ===========>")
    print(f"待處理項目數：{len(Dict_List)}（{', '.join(Dict_List.keys())}）")
    if args.col_addr:
        print(f"全域標音欄覆寫：{args.col_addr}")
    result = process_all(args.col_addr)
    print("<=========== d100 作業結束 ===========>")
    return result


if __name__ == "__main__":
    sys.exit(main())
