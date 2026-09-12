"""b100_字幕加注漢字標音.py v0.1.1

自【01_漢字標音】工作表【Q欄】第 4 列起，讀取儲存格內漢字，
依活頁簿層級名稱【漢字標音】（env!C8）指定之標音方法轉換，
將【漢字標音】（不含漢字，作法同【海海人生】工作表 H 欄）寫入對映【R欄】。

更新紀錄：
 - v0.1.1 2026-09-12: R 欄改為只填漢字標音；--file 預設於 output9 尋找；預設使用作用中活頁簿。
 - v0.1.0 2026-09-12: 初版。參考 a200 之程式骨架，改為處理字幕工作表 Q/R 欄。
"""

# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import os
import sys
from pathlib import Path

import xlwings as xw
from dotenv import load_dotenv

from mod_ca_ji_tian import HanJiTian
from mod_excel_access import get_value_by_name
from mod_logging import (
    init_logging,
    logging_exc_error,
    logging_exception,
    logging_process_step,
    logging_warning,
)
from mod_標音 import PiauIm, ca_ji_tng_piau_im, is_han_ji

# =========================================================================
# 常數定義
# =========================================================================
EXIT_CODE_SUCCESS = 0
EXIT_CODE_NO_FILE = 1
EXIT_CODE_INVALID_INPUT = 2
EXIT_CODE_SAVE_FAILURE = 3
EXIT_CODE_PROCESS_FAILURE = 10
EXIT_CODE_UNKNOWN_ERROR = 99

SHEET_NAME = "01_漢字標音"
HAN_BUN_COL = 17  # Q 欄：漢文
PIAU_IM_COL = 18  # R 欄：漢字標音
START_ROW = 4
PIAU_IM_HUAT_NAME = "漢字標音"
PIAU_IM_HUAT_FALLBACK = ("env", "C8")
DEFAULT_OUTPUT_DIR = "output9"

# 標點轉半形，緊接前一個標音（參考 a520／【海海人生】H 欄）
PUNCTUATION_MAP = {
    "，": ",",
    "。": ".",
    "！": "!",
    "？": "?",
    "；": ";",
    "：": ":",
    "、": ",",
    "—": "—",
    "…": "…",
}
# 書名號等不進入標音行（【海海人生】H2：《海海人生》→ Hái-hái jîn-sing）
SKIP_PUNCTUATION = set("《》〈〉「」『』（）()　")
SENTENCE_END_PUNCTUATIONS = {".", "!", "?"}

# =========================================================================
# 設定日誌
# =========================================================================
init_logging()
load_dotenv()

DB_HO_LOK_UE = os.getenv("DB_HO_LOK_UE", "Ho_Lok_Ue.db")
DB_KONG_UN = os.getenv("DB_KONG_UN", "Kong_Un.db")


# =========================================================================
# 作業協助函數
# =========================================================================
def _named_value(wb, name: str, default: str = "") -> str:
    """讀取活頁簿層級名稱；若無值則回傳 default。"""
    value = get_value_by_name(wb, name)
    if value is None:
        return default
    return str(value).strip()


def get_piau_im_huat(wb) -> str:
    """自活頁簿名稱【漢字標音】讀取標音方法，必要時改讀 env!C8。"""
    value = _named_value(wb, PIAU_IM_HUAT_NAME)
    if value:
        return value

    sheet_name, cell_addr = PIAU_IM_HUAT_FALLBACK
    try:
        fallback = wb.sheets[sheet_name].range(cell_addr).value
    except Exception:
        fallback = None
    if fallback:
        return str(fallback).strip()

    raise ValueError("無法自活頁簿名稱【漢字標音】或 env!C8 讀取漢字標音方法。")


def get_han_ji_khoo_and_db(wb) -> tuple[str, str]:
    """依名稱【漢字庫】決定字典名稱與資料庫檔名。"""
    han_ji_khoo_name = _named_value(wb, "漢字庫", "河洛話") or "河洛話"
    db_name = DB_HO_LOK_UE if han_ji_khoo_name == "河洛話" else DB_KONG_UN
    return han_ji_khoo_name, db_name


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


class HanJiPiauImLookup:
    """查字典並依指定標音方法轉換；同一漢字只查一次。"""

    def __init__(self, ji_tian: HanJiTian, piau_im: PiauIm, han_ji_khoo: str, piau_im_huat: str, ue_im_lui_piat: str):
        self.ji_tian = ji_tian
        self.piau_im = piau_im
        self.han_ji_khoo = han_ji_khoo
        self.piau_im_huat = piau_im_huat
        self.ue_im_lui_piat = ue_im_lui_piat
        self._cache: dict[str, str] = {}
        self.missing_chars: list[str] = []

    def lookup(self, han_ji: str) -> str:
        if han_ji in self._cache:
            return self._cache[han_ji]

        result = self.ji_tian.han_ji_ca_piau_im(han_ji=han_ji, ue_im_lui_piat=self.ue_im_lui_piat)
        han_ji_piau_im = ""
        if result:
            _tai_gi_im_piau, han_ji_piau_im = ca_ji_tng_piau_im(
                entry=result[0],
                han_ji_khoo=self.han_ji_khoo,
                piau_im=self.piau_im,
                piau_im_huat=self.piau_im_huat,
            )
            han_ji_piau_im = str(han_ji_piau_im or "").strip()

        if not han_ji_piau_im:
            logging_warning(f"漢字【{han_ji}】查無可用之【{self.piau_im_huat}】標音。")
            if han_ji not in self.missing_chars:
                self.missing_chars.append(han_ji)

        self._cache[han_ji] = han_ji_piau_im
        return han_ji_piau_im


class PiauImLineBuilder:
    """組建標音行：音節空白分隔、標點緊接、句首大寫（同【海海人生】H 欄）。"""

    def __init__(self):
        self.tokens: list[str] = []
        self.capitalize_next = True

    def add_piau_im(self, piau_im: str):
        piau_im = str(piau_im).strip()
        if not piau_im:
            return
        if self.capitalize_next:
            piau_im = piau_im[0].upper() + piau_im[1:]
            self.capitalize_next = False
        self.tokens.append(piau_im)

    def add_punctuation(self, han_ji_punct: str):
        if han_ji_punct in SKIP_PUNCTUATION:
            return
        punct = PUNCTUATION_MAP.get(han_ji_punct)
        if punct is None:
            return
        if self.tokens:
            self.tokens[-1] += punct
        else:
            self.tokens.append(punct)
        if punct in SENTENCE_END_PUNCTUATIONS:
            self.capitalize_next = True

    def build(self) -> str:
        return " ".join(self.tokens)


def build_piau_im_line(text: str, lookup: HanJiPiauImLookup) -> str:
    """將文句轉成僅含漢字標音的一行（不附漢字）。"""
    if text is None:
        return ""

    builder = PiauImLineBuilder()
    for ch in str(text):
        if is_han_ji(ch):
            piau_im = lookup.lookup(ch)
            builder.add_piau_im(piau_im if piau_im else ch)
        else:
            builder.add_punctuation(ch)
    return builder.build()


def get_workbook(file_arg: str | None, project_root: Path):
    """預設使用作用中活頁簿；僅在指定 --file 時才依 output9 開啟檔案。"""
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


def process(wb, _args) -> int:
    """讀取 Q 欄漢文，將漢字標音寫入 R 欄。"""
    logging_process_step("<=========== 作業開始！==========>")

    sheet_name = SHEET_NAME
    try:
        sheet_names = [s.name for s in wb.sheets]
        if sheet_name not in sheet_names:
            raise ValueError(f"活頁簿找不到工作表【{sheet_name}】，現有工作表：{sheet_names}")

        han_ji_khoo_name, db_name = get_han_ji_khoo_and_db(wb)
        piau_im_huat = get_piau_im_huat(wb)
        ue_im_lui_piat = _named_value(wb, "語音類型", "文讀音") or "文讀音"

        logging_process_step(f"活頁簿：{wb.fullname}")
        logging_process_step(f"工作表：{sheet_name}")
        logging_process_step(f"漢字庫：{han_ji_khoo_name}（{db_name}）")
        logging_process_step(f"語音類型：{ue_im_lui_piat}")
        logging_process_step(f"漢字標音方法：{piau_im_huat}")

        ji_tian = HanJiTian(db_name)
        piau_im = PiauIm(han_ji_khoo=han_ji_khoo_name)
        lookup = HanJiPiauImLookup(
            ji_tian=ji_tian,
            piau_im=piau_im,
            han_ji_khoo=han_ji_khoo_name,
            piau_im_huat=piau_im_huat,
            ue_im_lui_piat=ue_im_lui_piat,
        )

        sheet = wb.sheets[sheet_name]
        sheet.activate()

        ji_tian.connect()
        try:
            processed = 0
            row = START_ROW
            while True:
                han_bun_cell = sheet.range((row, HAN_BUN_COL))
                han_bun = han_bun_cell.value
                if han_bun is None or str(han_bun).strip() == "":
                    break

                piau_im_line = build_piau_im_line(str(han_bun), lookup)
                piau_im_cell = sheet.range((row, PIAU_IM_COL))
                piau_im_cell.value = piau_im_line

                q_addr = f"{xw.utils.col_name(HAN_BUN_COL)}{row}"
                r_addr = f"{xw.utils.col_name(PIAU_IM_COL)}{row}"
                print("-" * 80)
                print(f"{q_addr}：{han_bun}")
                print(f"{r_addr}：{piau_im_line}")

                processed += 1
                row += 1
        finally:
            ji_tian.disconnect()

        print("=" * 80)
        logging_process_step(f"已完成字幕加注，共處理 {processed} 列。")
        if lookup.missing_chars:
            missing = "、".join(lookup.missing_chars)
            logging_warning(f"下列漢字查無標音，已於標音行原樣保留：{missing}")
        return EXIT_CODE_SUCCESS

    except Exception as e:
        logging_exception(
            msg=f"在【{sheet_name}】工作表，為 Q 欄漢字加注漢字標音時發生例外！",
            error=e,
        )
        raise


def main(args) -> int:
    """主程式：預設處理作用中活頁簿。"""
    current_file_path = Path(__file__).resolve()
    project_root = current_file_path.parent
    program_name = current_file_path.stem

    logging_process_step(f"《========== 程式開始執行：{program_name} ==========》")
    logging_process_step(f"專案根目錄為: {project_root}")

    wb = None
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


def test_01():
    """測試：依字典將例句轉成漢字標音行。"""
    print("=" * 70)
    print("測試字幕加注漢字標音")
    print("=" * 70)

    try:
        ji_tian = HanJiTian(DB_HO_LOK_UE)
        piau_im = PiauIm(han_ji_khoo="河洛話")
        lookup = HanJiPiauImLookup(
            ji_tian=ji_tian,
            piau_im=piau_im,
            han_ji_khoo="河洛話",
            piau_im_huat="台羅拼音",
            ue_im_lui_piat="文讀音",
        )
        samples = [
            "《道德經。第十一章》",
            "卅輻同一轂，當其無，有車之用也。",
            "故有之以為利，無之以為用。",
        ]
        ji_tian.connect()
        try:
            for text in samples:
                piau_im_line = build_piau_im_line(text, lookup)
                print(f"原文：{text}")
                print(f"標音：{piau_im_line}")
                print("-" * 70)
        finally:
            ji_tian.disconnect()

        if lookup.missing_chars:
            print(f"缺字：{'、'.join(lookup.missing_chars)}")
        print("測試完成")
    except Exception as e:
        print(f"測試失敗：{e}")
        import traceback

        traceback.print_exc()


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description="自【01_漢字標音】工作表 Q 欄讀取漢文，將漢字標音寫入 R 欄",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用範例：
  python b100_字幕加注漢字標音.py
      # 預設：處理 Excel 作用中活頁簿（請先開啟檔案）
  python b100_字幕加注漢字標音.py --file 【字幕】帛書版道德經。第十一章.xlsx
      # 僅檔名時，預設於 output9 目錄尋找
  python b100_字幕加注漢字標音.py --test
""",
    )
    parser.add_argument(
        "--test",
        action="store_true",
        help="執行測試模式（不寫入 Excel）",
    )
    parser.add_argument(
        "--file",
        dest="file",
        default=None,
        help=f"活頁簿檔名或路徑；僅檔名時預設於 {DEFAULT_OUTPUT_DIR} 尋找。未指定則使用作用中活頁簿",
    )
    args = parser.parse_args()

    if args.test:
        test_01()
    else:
        exit_code = main(args)
        if exit_code != EXIT_CODE_SUCCESS:
            print(f"程式異常終止，錯誤代碼為: {exit_code}")
            sys.exit(exit_code)
