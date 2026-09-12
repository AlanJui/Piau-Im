# -*- coding: utf-8 -*-
"""Read Show times from an SRT file and write them into Excel via xlwings."""
from __future__ import annotations

import argparse
import re
from pathlib import Path

import xlwings as xw

TIME_RE = re.compile(
    r"(\d{1,2}:\d{2}:\d{2}[,.]\d{1,3})\s*-->\s*(\d{1,2}:\d{2}:\d{2}[,.]\d{1,3})"
)
FONT = "Microsoft JhengHei"
SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_SRT = SCRIPT_DIR / "_tmp.srt"
DEFAULT_XLSX = SCRIPT_DIR.parent / "SRT編輯器.xlsx"
TARGET_SHEET = "02_計算時長"
SHOW_START_ROW = 2
SHOW_END_ROW = 62
SHOW_COL = "B"


def normalize_srt_time(value: str) -> str:
    text = value.strip().replace(".", ",", 1) if "," not in value else value.strip()
    hours, minutes, rest = text.split(":")
    seconds, millis = rest.split(",")
    return f"{int(hours):02d}:{int(minutes):02d}:{int(seconds):02d},{int(millis):03d}"


def parse_srt_shows(path: Path) -> list[tuple[int, str]]:
    raw = path.read_text(encoding="utf-8-sig").replace("\r\n", "\n").strip()
    if not raw:
        raise ValueError(f"SRT is empty: {path}")

    rows: list[tuple[int, str]] = []
    for block in (b for b in raw.split("\n\n") if b.strip()):
        lines = [ln.strip() for ln in block.split("\n") if ln.strip()]
        if len(lines) < 2:
            continue
        index = int(lines[0]) if lines[0].isdigit() else len(rows) + 1
        match = TIME_RE.search(lines[1])
        if not match:
            raise ValueError(f"Cannot parse time line in cue {index}: {lines[1]}")
        rows.append((index, normalize_srt_time(match.group(1))))
    if not rows:
        raise ValueError(f"No cues found in: {path}")
    return rows


def resolve_srt(specified: Path | None) -> Path:
    if specified is None:
        srt_path = DEFAULT_SRT
    else:
        srt_path = specified.expanduser()
        srt_path = (
            srt_path.resolve()
            if srt_path.is_absolute()
            else (Path.cwd() / srt_path).resolve()
        )
    if not srt_path.is_file():
        raise SystemExit(f"SRT not found: {srt_path}")
    return srt_path


def excel_app() -> xw.App:
    app = xw.apps.active if xw.apps.count else xw.App(visible=True)
    app.visible = True
    return app


def find_open_book(path: Path) -> xw.Book | None:
    target = path.resolve()
    for app in xw.apps:
        for book in app.books:
            try:
                full = Path(book.fullname).resolve()
            except Exception:
                continue
            if full == target:
                return book
    return None


def write_show_column(sheet: xw.Sheet, shows: list[str]) -> None:
    capacity = SHOW_END_ROW - SHOW_START_ROW + 1
    if len(shows) > capacity:
        print(f"warning: SRT has {len(shows)} cues; only B{SHOW_START_ROW}:B{SHOW_END_ROW} ({capacity}) written")
        shows = shows[:capacity]

    rng = sheet.range(f"{SHOW_COL}{SHOW_START_ROW}:{SHOW_COL}{SHOW_END_ROW}")
    rng.number_format = "@"
    values = [[show] for show in shows]
    values.extend([[None] for _ in range(capacity - len(shows))])
    rng.value = values
    rng.font.name = FONT


def paste_into_editor(rows: list[tuple[int, str]]) -> Path:
    xlsx_path = DEFAULT_XLSX.resolve()
    if not xlsx_path.is_file():
        raise SystemExit(f"Excel not found: {xlsx_path}")

    book = find_open_book(xlsx_path)
    if book is None:
        book = excel_app().books.open(str(xlsx_path))
    else:
        book.app.visible = True

    try:
        sheet = book.sheets[TARGET_SHEET]
    except Exception as exc:
        raise SystemExit(f"Worksheet not found: {TARGET_SHEET}") from exc

    write_show_column(sheet, [show for _, show in rows])
    sheet.activate()
    sheet.range(f"{SHOW_COL}{SHOW_START_ROW}").select()
    book.activate()
    return xlsx_path


def open_new_workbook(rows: list[tuple[int, str]], save_path: Path) -> Path:
    app = excel_app()
    app.display_alerts = False
    wb = app.books.add()
    sheet = wb.sheets[0]
    sheet.name = "Show"

    data = [["編號", "Show"], *[[index, show] for index, show in rows]]
    last_row = len(rows) + 1
    sheet.range("A1").value = data
    sheet.range(f"B2:B{last_row}").number_format = "@"
    sheet.range(f"A1:B{last_row}").font.name = FONT
    sheet.range(f"A1:B{last_row}").api.HorizontalAlignment = -4108
    header = sheet.range("A1:B1")
    header.font.bold = True
    header.font.color = (255, 255, 255)
    header.color = (31, 78, 121)
    sheet.range("A1").column_width = 10
    sheet.range("B1").column_width = 18
    sheet.range("A2").select()
    app.api.ActiveWindow.FreezePanes = True
    sheet.range("A1").select()

    save_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(save_path))
    wb.activate()
    return save_path


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Paste SRT Show times into SRT編輯器.xlsx, or create a new workbook with --new."
    )
    parser.add_argument(
        "srt",
        nargs="?",
        type=Path,
        help="Path to .srt file (default: _tmp.srt next to this script)",
    )
    parser.add_argument(
        "--new",
        action="store_true",
        help="Create a new Excel workbook instead of writing to 02_計算時長",
    )
    args = parser.parse_args()

    srt_path = resolve_srt(args.srt)
    rows = parse_srt_shows(srt_path)
    if args.new:
        xlsx_path = open_new_workbook(rows, srt_path.with_name("_tmp_show.xlsx"))
        print("mode=new workbook")
    else:
        xlsx_path = paste_into_editor(rows)
        print(f"mode=paste {TARGET_SHEET}!{SHOW_COL}{SHOW_START_ROW}:{SHOW_COL}{SHOW_END_ROW}")

    print(f"srt={srt_path}")
    print(f"cues={len(rows)}")
    print(xlsx_path)


if __name__ == "__main__":
    main()
