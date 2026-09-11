# -*- coding: utf-8 -*-
"""Read Show times from an SRT file and open them in a new Excel workbook."""
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


def open_in_excel(rows: list[tuple[int, str]], save_path: Path) -> Path:
    app = xw.apps.active if xw.apps.count else xw.App(visible=True)
    app.visible = True
    app.display_alerts = False
    wb = app.books.add()
    sheet = wb.sheets[0]
    sheet.name = "Show"

    data = [["編號", "Show"], *[[index, show] for index, show in rows]]
    last_row = len(rows) + 1
    sheet.range("A1").value = data
    sheet.range(f"B2:B{last_row}").number_format = "@"
    sheet.range(f"A1:B{last_row}").font.name = FONT
    sheet.range(f"A1:B{last_row}").api.HorizontalAlignment = -4108  # xlCenter
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
        description="Export SRT Show times to a new Excel workbook and open it."
    )
    parser.add_argument(
        "srt",
        nargs="?",
        type=Path,
        help="Path to .srt file (default: _tmp.srt next to this script)",
    )
    args = parser.parse_args()

    srt_path = resolve_srt(args.srt)
    rows = parse_srt_shows(srt_path)
    xlsx_path = srt_path.with_name("_tmp_show.xlsx")
    xlsx_path = open_in_excel(rows, xlsx_path)
    print(f"srt={srt_path}")
    print(f"cues={len(rows)}")
    print(xlsx_path)


if __name__ == "__main__":
    main()
