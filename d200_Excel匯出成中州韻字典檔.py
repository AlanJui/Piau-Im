""" d200_Excel匯出成中州韻字典檔.py v0.0.1
摘要：
    將 Excel 檔案中【工作表】的資料轉換成中州韻字典檔。
    1. 打開 Excel 檔案，讀取【台羅拼音工作表】/【台語注音二式工作表】的【字典檔內容】（A欄：E欄）資料。
    2. 套用【Rime 字典檔的標頭內容】及【字典檔內容】，製作成【中州韻字典檔】。
    3. 將【中州韻字典檔】存放在【target_dir_path】目錄路徑下。
    4. 依 d000_定義檔.Dict_List 逐一處理全部 Excel 活頁簿（各匯出 TL 與 BPM2）。

指令：
    python d200_Excel匯出成中州韻字典檔.py

參考：
    - a860_將Excel檔中的台羅拼音轉換成台語注音二式.py
"""  # noqa: N999

from __future__ import annotations

import logging
import sys
from pathlib import Path

import xlwings as xw

from d000_定義檔 import (
    Dict_List,
    expand_dir_path,
    format_rime_header,
    get_source_sheet_name,
    get_target_sheet_name,
    source_dir_path,
    target_dir_path,
)

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


def resolve_target_dir() -> Path:
    return expand_dir_path(target_dir_path)


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


def cell_to_str(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def read_sheet_body(sheet: xw.Sheet) -> list[str]:
    """
    讀取工作表 A–E 欄資料本體（略過表頭第 1 列），
    轉成以 Tab 分隔的 RIME 字典資料行。
    stem 空白時填入 NA，避免 Tab 欄位在編輯器中被誤刪。
    """
    last_row = sheet.range("A" + str(sheet.cells.last_cell.row)).end("up").row
    if last_row < 2:
        return []

    values = sheet.range((2, 1), (last_row, 5)).value
    if last_row == 2:
        values = [values]

    lines: list[str] = []
    for row in values:
        if row is None:
            continue
        cells = list(row) if isinstance(row, (list, tuple)) else [row]
        while len(cells) < 5:
            cells.append(None)

        text = cell_to_str(cells[0])
        code = cell_to_str(cells[1])
        weight = cell_to_str(cells[2])
        stem = cell_to_str(cells[3]) or "NA"
        create = cell_to_str(cells[4])

        # 整列皆空則略過
        if not any([text, code, weight, stem != "NA", create]):
            continue

        lines.append(f"{text}\t{code}\t{weight}\t{stem}\t{create}")
    return lines


def export_one_dict(
    wb,
    *,
    dict_key: str,
    dict_cfg: dict,
    sheet_name: str,
    system_name: str,
    scheme_key: str,
    target_dir: Path,
) -> Path | None:
    """自指定工作表匯出一支中州韻字典檔；工作表不存在則略過。"""
    sheet_names = [s.name for s in wb.sheets]
    if sheet_name not in sheet_names:
        print(f"⚠️ [{dict_key}] 找不到工作表【{sheet_name}】，略過 {scheme_key} 匯出。")
        return None

    scheme_cfg = dict_cfg.get(scheme_key) or {}
    dict_name = scheme_cfg.get("輸入方案名稱")
    if not dict_name:
        raise ValueError(f"Dict_List['{dict_key}']['{scheme_key}'] 缺少『輸入方案名稱』")

    header = format_rime_header(
        dict_cfg,
        輸入方案名稱=dict_name,
        漢字標音系統名稱=system_name,
    )
    body_lines = read_sheet_body(wb.sheets[sheet_name])

    target_dir.mkdir(parents=True, exist_ok=True)
    out_path = target_dir / f"{dict_name}.dict.yaml"
    with open(out_path, "w", encoding="utf-8", newline="\n") as f:
        f.write(header)
        if not header.endswith("\n"):
            f.write("\n")
        if body_lines:
            f.write("\n")
            f.write("\n".join(body_lines))
            f.write("\n")

    print(f"✅ [{dict_key}] 【{sheet_name}】→ {out_path.name}（{len(body_lines)} 筆）")
    logging.info("d200 [%s] exported %s rows to %s", dict_key, len(body_lines), out_path)
    return out_path


def process_one(dict_key: str, dict_cfg: dict, target_dir: Path) -> list[Path]:
    """處理 Dict_List 中的單一字典項目，匯出 TL／BPM2（若工作表存在）。"""
    workbook_path = expand_dir_path(source_dir_path) / dict_cfg["WorkBook檔案"]
    sheet_tl = get_source_sheet_name(dict_cfg)
    sheet_bpm2 = get_target_sheet_name(dict_cfg)

    print(f"\n----- [{dict_key}] {dict_cfg['WorkBook檔案']} -----")
    print(f"TL 表：{sheet_tl}；BPM2 表：{sheet_bpm2}")

    export_specs = (
        (sheet_tl, "台羅拼音", "TL"),
        (sheet_bpm2, "台語注音二式", "BPM2"),
    )

    wb = None
    opened_by_script = False
    exported: list[Path] = []
    try:
        wb, opened_by_script = get_workbook(workbook_path)
        for sheet_name, system_name, scheme_key in export_specs:
            path = export_one_dict(
                wb,
                dict_key=dict_key,
                dict_cfg=dict_cfg,
                sheet_name=sheet_name,
                system_name=system_name,
                scheme_key=scheme_key,
                target_dir=target_dir,
            )
            if path is not None:
                exported.append(path)
        if not exported:
            raise RuntimeError(f"[{dict_key}] 未匯出任何字典檔（請確認工作表是否存在）")
        return exported
    finally:
        if wb is not None and opened_by_script:
            wb.close()


def process_all() -> int:
    target_dir = resolve_target_dir()
    failures: list[str] = []
    all_exported: list[Path] = []

    for dict_key, dict_cfg in Dict_List.items():
        try:
            exported = process_one(dict_key, dict_cfg, target_dir)
            all_exported.extend(exported)
        except Exception as e:
            failures.append(dict_key)
            print(f"❌ [{dict_key}] 作業失敗：{e}")
            logging.error("d200 [%s] 作業失敗：%s", dict_key, e, exc_info=True)

    print(f"\n✅ 合計匯出 {len(all_exported)} 個字典檔 → {target_dir}")
    if failures:
        print(f"⚠️ 失敗項目：{', '.join(failures)}")
        return EXIT_CODE_FAILURE
    return EXIT_CODE_SUCCESS


def main() -> int:
    print("<=========== d200 作業開始 ===========>")
    print(f"待處理項目數：{len(Dict_List)}（{', '.join(Dict_List.keys())}）")
    print(f"來源目錄：{expand_dir_path(source_dir_path)}")
    print(f"標的目錄：{resolve_target_dir()}")
    result = process_all()
    print("<=========== d200 作業結束 ===========>")
    return result


if __name__ == "__main__":
    sys.exit(main())
