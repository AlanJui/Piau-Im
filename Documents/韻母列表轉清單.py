"""
將【韻母列表】工作表 B4:L30 的韻母表格，改寫成清單後寫入同表 N 欄起。

表格直欄（第 4 列類別，合併儲存格向右沿用）：
    B        韻腹
    C、D     陰聲韻
    E、F、G  陽聲韻
    H–L      入聲韻

清單每列三欄，對應既有標題列（第 5 列）：
    N  韻腹
    O  韻母拼音
    P  韻母類別

空白儲存格與佔位符「－」略過。同一韻腹依 C → L 由左至右展開，
韻腹之間不插空白列。結果自 N6 寫入，結束列依筆數而定。

活頁簿若已在 Excel 開啟，會寫入該視窗並存檔，不另行開新檔。
"""

from pathlib import Path

import xlwings as xw

WORKBOOK_PATH = Path(__file__).resolve().parent / "_聲韻調對照表.xlsx"
SHEET_NAME = "韻母列表"
SOURCE_RANGE = "B4:L30"
LIST_START_ROW = 6
LIST_COLUMNS = (14, 16)  # N:P
CLEAR_LAST_ROW = 400
# 非韻母拼音之佔位符（含全形連字號「－」）
PLACEHOLDERS = {"－", "-", "—", "–", "─"}
# Excel 常數：xlVAlignCenter
XL_VALIGN_CENTER = -4108
FONT_NAME = "Noto Sans TC Medium"
FONT_SIZE = 16


def _is_empty(value: object) -> bool:
    """空白或佔位符視為無韻母。"""
    if value is None:
        return True
    text = str(value).strip()
    return text == "" or text in PLACEHOLDERS


def _categories_by_column(header_row: list) -> dict[int, str]:
    """依第 4 列標題決定 C–L 各欄的韻母類別。合併儲存格只在左上角有值，向右沿用。"""
    categories: dict[int, str] = {}
    current: str | None = None
    # B4:L30 讀出後，索引 0 為 B 欄，1–10 為 C–L 欄
    for index in range(1, 11):
        value = header_row[index] if index < len(header_row) else None
        if not _is_empty(value):
            current = str(value).strip()
        if current is None:
            raise ValueError(f"第 {index + 2} 欄（C 起算）找不到韻母類別標題")
        categories[index] = current
    return categories


def table_to_list(table: list[list]) -> list[list[str]]:
    """將 B4:L30 二維表格改寫成［韻腹, 韻母拼音, 韻母類別］清單。"""
    if not table or len(table) < 3:
        raise ValueError(f"{SOURCE_RANGE} 資料不足，無法轉成清單")

    categories = _categories_by_column(table[0])
    records: list[list[str]] = []
    # 索引 0、1 為第 4、5 列標題；自第 6 列（索引 2）起為韻腹
    for data_row in table[2:]:
        if not data_row or _is_empty(data_row[0]):
            continue
        nucleus = str(data_row[0]).strip()
        for index in range(1, 11):
            spelling = data_row[index] if index < len(data_row) else None
            if _is_empty(spelling):
                continue
            records.append([nucleus, str(spelling).strip(), categories[index]])
    return records


def _as_table(raw: object) -> list[list]:
    """xlwings 讀取多列範圍時為二維清單；單列時補成二維。"""
    if raw is None:
        return []
    if not isinstance(raw, list):
        return [[raw]]
    if raw and not isinstance(raw[0], list):
        return [raw]
    return raw


def write_list(sheet: xw.Sheet, records: list[list[str]]) -> str:
    """清除 N6:P 舊清單後寫入，並套用與既有清單相同的字型。回傳寫入範圍位址。"""
    if not records:
        raise ValueError("清單是空的，未寫入工作表")

    end_row = LIST_START_ROW + len(records) - 1
    sheet.range((LIST_START_ROW, LIST_COLUMNS[0]), (CLEAR_LAST_ROW, LIST_COLUMNS[1])).clear_contents()

    target = sheet.range((LIST_START_ROW, LIST_COLUMNS[0]), (end_row, LIST_COLUMNS[1]))
    target.value = records
    target.font.name = FONT_NAME
    target.font.size = FONT_SIZE
    target.api.VerticalAlignment = XL_VALIGN_CENTER
    return f"N{LIST_START_ROW}:P{end_row}"


def main() -> None:
    if not WORKBOOK_PATH.is_file():
        raise FileNotFoundError(f"找不到活頁簿：{WORKBOOK_PATH}")

    workbook = xw.Book(str(WORKBOOK_PATH))
    sheet = workbook.sheets[SHEET_NAME]
    records = table_to_list(_as_table(sheet.range(SOURCE_RANGE).value))
    address = write_list(sheet, records)
    workbook.save()

    written = _as_table(sheet.range(address).value)
    if written != records:
        raise RuntimeError("寫入後讀回的清單與來源表格不一致")

    print(f"已將 {SOURCE_RANGE} 改寫成 {len(records)} 筆清單，寫入【{SHEET_NAME}】{address}")


if __name__ == "__main__":
    main()
