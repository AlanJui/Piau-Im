"""a830_BanLam中州韻字典轉台羅拼音.py v0.1.0
【功能摘要】：
將【BanLam】中州韻字典檔，轉換成【台羅拼音】字典檔。

中州韻字典檔：banlam.dict.yaml 為 GitHub 之 a-thok/rime-hokkien 專案（RIME福建話/閩南話
輸入方案）使用之字典檔。其特色為：內含【泉/漳/廈/閩南地方腔】使用之單字/辭彙。

此字典所收錄之漢字讀音，其羅馬拼音系統雖以【台羅拼音】為基礎，但為兼容【泉/漳/廈/閩南方音】各種
腔調，所以在【聲母】、【韻母】、【聲調】做了調整，成為非標準之【台羅拼音】。

原為 yaml 格式之文字檔：banlam.dict.yaml，現已置入【Ban_Lam中州韻字典轉換.xlsx】Excel 檔案
中。本程式之目的便是要將存放在【balam.dict.yaml】工作表，在【B欄】的非標準【台羅拼音】，
轉換成標準之【台羅拼音】，並寫入【ji_khoo_ban_lam.dict.yaml】工作表之【B欄】。

【注意事項】：
 1. 當【A欄】內容值為【##】，表示該行為【註解】，【B欄】之內容無須轉換，可直接。
 2. 這個字典檔的格式，不太重視一致性，以致【C欄】之內容，不是每筆資料（每行）都有；且有 10, 1%, 0.1% 等
    各種數值格式。所以轉換作業無須處理【C欄】之內容。
 3. 【轉換規則】以【正規式】描述，記載於【轉換規則】工作表之【A欄】到【B欄】。自【C欄】開始，各欄
    之內容均請忽略，那些只是為製作【A欄】與【B欄】內容，之工作底稿。
"""

# =========================================================================
# 載入程式所需套件/模組
# =========================================================================
import logging
import re
import sys
from pathlib import Path

import xlwings as xw

# =========================================================================
# 設定日誌
# =========================================================================
logging.basicConfig(
    filename="process_log.txt",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)

# =========================================================================
# 常數定義
# =========================================================================
EXIT_CODE_SUCCESS = 0
EXIT_CODE_FAILURE = 1
EXIT_CODE_NO_FILE = 2

# 活頁簿檔名（實際檔名為 Ban-Lam…；說明中之 Ban_Lam／balam 為筆誤）
WORKBOOK_CANDIDATES = [
    Path(__file__).resolve().parent / "Ban-Lam中州韻字典轉換.xlsx",
    Path(__file__).resolve().parent / "Ban_Lam中州韻字典轉換.xlsx",
]
WORKBOOK_NAME_KEYWORDS = ("Ban-Lam", "Ban_Lam", "BanLam")

SHEET_RULES = "轉換規則"
SHEET_SOURCE = "banlam.dict.yaml"
SHEET_TARGET = "ji_khoo_ban_lam.dict.yaml"

COMMENT_MARK = "##"  # A 欄為此值時，B 欄不轉換、直接複製
PROGRESS_EVERY = 5000


# =========================================================================
# 轉換規則
# =========================================================================
def fix_replacement(repl: str) -> str:
    """
    將常見編輯器風格之置換語法 $1、$2 轉成 Python re.sub 使用的 \\1、\\2。
    """
    return re.sub(r"\$(\d)", lambda m: "\\" + m.group(1), repl)


def load_conversion_rules(wb) -> list[tuple[str, str]]:
    """
    自【轉換規則】工作表讀取 A 欄（搜尋）、B 欄（更換）。
    第 1 列為標題，略過；A 欄空白之列略過；C 欄以後忽略。
    規則依列序由上而下套用（同型衝突時，先列者先生效）。
    """
    sheet = wb.sheets[SHEET_RULES]
    last_row = sheet.range("A" + str(sheet.cells.last_cell.row)).end("up").row
    if last_row < 2:
        raise ValueError(f"【{SHEET_RULES}】工作表無轉換規則。")

    rows = sheet.range(f"A2:B{last_row}").value
    if rows is None:
        raise ValueError(f"【{SHEET_RULES}】工作表無轉換規則。")
    if not isinstance(rows[0], list):
        rows = [rows]

    rules: list[tuple[str, str]] = []
    for row in rows:
        if not row or row[0] is None or str(row[0]).strip() == "":
            continue
        pattern = str(row[0])
        replacement = fix_replacement("" if row[1] is None else str(row[1]))
        # 預編譯驗證正規式是否合法
        try:
            re.compile(pattern)
        except re.error as e:
            raise ValueError(f"無效之搜尋正規式：{pattern!r}（{e}）") from e
        rules.append((pattern, replacement))

    if not rules:
        raise ValueError(f"【{SHEET_RULES}】未讀得任何有效規則。")
    return rules


def convert_syllable(syllable: str, rules: list[tuple[str, str]]) -> str:
    """對單一音節依序套用全部轉換規則。"""
    result = syllable
    for pattern, replacement in rules:
        result = re.sub(pattern, replacement, result)
    return result


def convert_code(code: str, rules: list[tuple[str, str]]) -> str:
    """
    將非標準台羅拼音轉成標準台羅拼音。
    多音節以空白分隔，逐音節套用規則後再以空白接回。
    """
    if code is None:
        return ""
    text = str(code)
    if text.strip() == "":
        return text

    parts = text.split(" ")
    return " ".join(convert_syllable(p, rules) if p else p for p in parts)


# =========================================================================
# Excel 存取
# =========================================================================
def get_workbook():
    """
    取得目標活頁簿：優先使用已開啟者；否則依候選路徑開啟。
    回傳 (wb, opened_by_script)。
    """
    for app in xw.apps:
        for book in app.books:
            if any(k in book.name for k in WORKBOOK_NAME_KEYWORDS):
                print(f"📌 使用已開啟之活頁簿：{book.name}")
                return book, False

    for path in WORKBOOK_CANDIDATES:
        if path.exists():
            print(f"📌 開啟活頁簿：{path}")
            return xw.Book(str(path)), True

    raise FileNotFoundError(
        "找不到【Ban-Lam中州韻字典轉換.xlsx】。"
        "請先開啟該檔，或置於專案根目錄後再執行。"
    )


def ensure_sheets(wb) -> tuple:
    names = [s.name for s in wb.sheets]
    for required in (SHEET_RULES, SHEET_SOURCE, SHEET_TARGET):
        if required not in names:
            raise ValueError(f"活頁簿缺少工作表：{required}（現有：{names}）")
    return wb.sheets[SHEET_SOURCE], wb.sheets[SHEET_TARGET]


# =========================================================================
# 主流程
# =========================================================================
def process() -> int:
    wb = None
    opened_by_script = False
    try:
        wb, opened_by_script = get_workbook()
        source, target = ensure_sheets(wb)

        print(f"📌 讀取轉換規則：【{SHEET_RULES}】")
        rules = load_conversion_rules(wb)
        print(f"✅ 已載入 {len(rules)} 條轉換規則（依列序套用）")

        last_row = source.range("A" + str(source.cells.last_cell.row)).end("up").row
        if last_row < 2:
            print("⚠️ 來源工作表無資料列。")
            return EXIT_CODE_FAILURE

        print(f"📌 讀取來源【{SHEET_SOURCE}】A2:B{last_row} …")
        data = source.range(f"A2:B{last_row}").value
        if data is None:
            print("⚠️ 來源工作表無資料列。")
            return EXIT_CODE_FAILURE
        if not isinstance(data[0], list):
            data = [data]

        converted_codes: list[str | None] = []
        converted_count = 0
        copied_comment_count = 0

        for idx, row in enumerate(data, start=2):
            han_ji = row[0] if row else None
            old_code = row[1] if row and len(row) > 1 else None

            if han_ji == COMMENT_MARK:
                # 註解列：B 欄直接複製，不轉換
                converted_codes.append(old_code)
                copied_comment_count += 1
            elif old_code is None or str(old_code).strip() == "":
                converted_codes.append(old_code)
            else:
                new_code = convert_code(str(old_code), rules)
                converted_codes.append(new_code)
                converted_count += 1

            done = idx - 1
            if done % PROGRESS_EVERY == 0:
                print(f"… 已處理 {done} / {len(data)} 列 …")

        # 寫入標的工作表 B 欄（文字格式，避免 Excel 誤判日期）
        print(f"📌 寫入標的【{SHEET_TARGET}】B2:B{last_row} …")
        target.range(f"B2:B{last_row}").number_format = "@"
        # xlwings 寫入直欄需為二維欄向量
        target.range("B2").value = [[c] for c in converted_codes]

        wb.save()
        print(
            f"✅ 轉換完成：轉換 {converted_count} 列，註解直接複製 {copied_comment_count} 列，"
            f"合計資料列 {len(data)}。"
        )
        print(f"✅ 已儲存：{wb.fullname or wb.name}")
        logging.info(
            "a830 BanLam 轉換完成：converted=%s, comments=%s, total=%s",
            converted_count,
            copied_comment_count,
            len(data),
        )
        return EXIT_CODE_SUCCESS

    except Exception as e:
        print(f"❌ 作業失敗：{e}")
        logging.error("a830 BanLam 轉換失敗：%s", e, exc_info=True)
        return EXIT_CODE_FAILURE
    finally:
        if wb is not None and opened_by_script:
            wb.close()


def main() -> int:
    print("<=========== a830 BanLam→台羅 作業開始 ===========>")
    print(f"來源工作表：{SHEET_SOURCE}（B 欄：非標準台羅）")
    print(f"標的工作表：{SHEET_TARGET}（B 欄：標準台羅）")
    print(f"轉換規則：{SHEET_RULES}（A→B，忽略 C 欄以後）")
    result = process()
    print("<=========== a830 BanLam→台羅 作業結束 ===========>")
    return result


if __name__ == "__main__":
    sys.exit(main())
