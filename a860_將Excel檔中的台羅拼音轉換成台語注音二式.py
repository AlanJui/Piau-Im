"""
a860_將Excel檔中的台羅拼音轉換成台語注音二式.py

【功能說明】：
將Excel檔案中，存放【台羅拼音】的工作表，轉換成存放【台語注音二式】的資料表。

【作業步驟】：
 1. 將【來源工作表】複製成【標的工作表】；
 2. 轉換【來源工作表】/【code】欄（即B欄）的【台羅拼音】（此處之拼音音節大多為2個以上）
   ，轉換後【台語注音二式】後，寫入【標的工作表】/【code】欄（即C欄）。

【作業環境】：

 - Python 套件：xlwings
 - 目錄：C:/Users/AlanJui/work/rime-tlpa/src/
 - 檔名：【漢字正字】中州韻輸入法字庫.xlsx
 - 來源工作表：RIME_Dict
 - 標的工作表：RIME_Dict_BPM2
"""

# =========================================================================
# 載入程式所需套件/模組
# =========================================================================
import logging
import sys
from pathlib import Path

import xlwings as xw

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

# =========================================================================
# 常數定義（對應上方【作業環境】說明）
# =========================================================================
EXIT_CODE_SUCCESS = 0
EXIT_CODE_FAILURE = 1
EXIT_CODE_NO_FILE = 2

WORKBOOK_DIR = Path(r"C:\Users\AlanJui\work\rime-tlpa\src")
WORKBOOK_NAME = "【漢字正字】中州韻輸入法字庫.xlsx"
WORKBOOK_PATH = WORKBOOK_DIR / WORKBOOK_NAME
SOURCE_SHEET = "RIME_Dict"
TARGET_SHEET = "RIME_Dict_BPM2"

# 實際表頭為：A=text、B=code、C=weight、D=stem、E=create
# 規格文「寫入 code 欄（即C欄）」依表頭名稱「code」為準（位於 B 欄）；
# 勿寫入 C 欄，以免覆寫 weight。
CODE_COL = 2  # B 欄


# =========================================================================
# 拼音轉換
# =========================================================================
def convert_tl_code_to_bpm2(code: str) -> str:
    """
    將【code】欄之台羅拼音轉成台語注音二式。
    多音節以空白分隔（如 'kio3 si7'），逐音節轉換後再以空白接回。
    轉換路徑：台羅 → TLPA → 台語注音二式（MPS2／BPM2）。
    """
    if not code:
        return ""

    syllables = str(code).strip().split()
    converted = []
    for syl in syllables:
        tlpa = convert_tl_to_tlpa(syl.lower()) or syl.lower()
        converted.append(convert_TLPA_to_MPS2(tlpa))
    return " ".join(converted)


# =========================================================================
# Excel 作業
# =========================================================================
def get_workbook():
    """
    取得目標活頁簿：若已在 Excel 開啟則直接使用，否則開啟檔案。
    """
    for app in xw.apps:
        for book in app.books:
            if book.name == WORKBOOK_NAME or Path(book.fullname or "").name == WORKBOOK_NAME:
                print(f"📌 使用已開啟之活頁簿：{book.name}")
                return book, False  # (wb, opened_by_script)

    if not WORKBOOK_PATH.exists():
        raise FileNotFoundError(f"找不到活頁簿檔案：{WORKBOOK_PATH}")

    print(f"📌 開啟活頁簿：{WORKBOOK_PATH}")
    return xw.Book(str(WORKBOOK_PATH)), True


def copy_source_to_target(wb) -> xw.Sheet:
    """
    將【來源工作表】複製為【標的工作表】。
    若標的工作表已存在，先刪除再複製，確保內容與來源一致。
    """
    sheet_names = [s.name for s in wb.sheets]
    if SOURCE_SHEET not in sheet_names:
        raise ValueError(f"找不到來源工作表：{SOURCE_SHEET}")

    if TARGET_SHEET in sheet_names:
        print(f"⚠️ 標的工作表【{TARGET_SHEET}】已存在，將先刪除再重新複製。")
        wb.sheets[TARGET_SHEET].delete()

    source = wb.sheets[SOURCE_SHEET]
    # 複製到來源工作表之後，並命名為標的工作表
    source.copy(after=source, name=TARGET_SHEET)
    target = wb.sheets[TARGET_SHEET]
    print(f"✅ 已將【{SOURCE_SHEET}】複製為【{TARGET_SHEET}】")
    return target


def convert_code_column(target: xw.Sheet) -> int:
    """
    將標的工作表【code】欄（B 欄）之台羅拼音轉成台語注音二式並寫回。
    回傳成功轉換之列數。
    """
    # 確認表頭
    header = target.range("A1:E1").value
    if not header or str(header[1]).strip().lower() != "code":
        raise ValueError(f"標的工作表 B1 應為 'code'，實際為：{header}")

    last_row = target.range("A" + str(target.cells.last_cell.row)).end("up").row
    if last_row < 2:
        print("⚠️ 標的工作表無資料列可轉換。")
        return 0

    # 先設為文字格式，避免轉換後之拼音被 Excel 誤判為日期
    target.range((2, CODE_COL), (last_row, CODE_COL)).number_format = "@"

    codes = target.range((2, CODE_COL), (last_row, CODE_COL)).value
    # 單列時 xlwings 回傳純量，統一成 list
    if last_row == 2:
        codes = [codes]

    new_codes = []
    converted_count = 0
    for idx, code in enumerate(codes, start=2):
        if code is None or str(code).strip() == "":
            new_codes.append(code)
            continue
        bpm2 = convert_tl_code_to_bpm2(code)
        new_codes.append(bpm2)
        converted_count += 1
        print(f"  ({idx}) {code} → {bpm2}")

    # 寫回 B 欄（縱向）
    target.range((2, CODE_COL)).options(transpose=True).value = new_codes
    return converted_count


def process() -> int:
    wb = None
    opened_by_script = False
    try:
        wb, opened_by_script = get_workbook()
        target = copy_source_to_target(wb)
        count = convert_code_column(target)
        wb.save()
        print(f"✅ 轉換完成：共 {count} 列。【code】欄已改為台語注音二式。")
        print(f"✅ 已儲存：{wb.fullname or wb.name}")
        logging.info("a860 轉換完成：%s 列 → %s", count, TARGET_SHEET)
        return EXIT_CODE_SUCCESS
    except Exception as e:
        print(f"❌ 作業失敗：{e}")
        logging.error("a860 作業失敗：%s", e, exc_info=True)
        return EXIT_CODE_FAILURE
    finally:
        # 僅關閉由本程式自行開啟的活頁簿；使用者原本開著者保持開啟
        if wb is not None and opened_by_script:
            wb.close()


# =========================================================================
# 主程式
# =========================================================================
def main() -> int:
    print("<=========== a860 作業開始 ===========>")
    print(f"來源：{WORKBOOK_NAME} / {SOURCE_SHEET}")
    print(f"標的：{TARGET_SHEET}（code 欄：台羅拼音 → 台語注音二式）")
    result = process()
    print("<=========== a860 作業結束 ===========>")
    return result


if __name__ == "__main__":
    sys.exit(main())
