"""c100_漢字標音儲存格改用變調標示格式.py v0.1.1

將【漢字注音】工作表各個【漢字標音】儲存格，依【漢字】儲存格底色，
改寫成【變調標音格式】或【本調標音格式】。

判斷規則：
- 漢字儲存格底色為淺灰 #E8E8E8 → 變調標音格式（例：duan⁵⁻³）
- 其餘漢字儲存格 → 本調標音格式（例：duan⁵）

掃表迴路邏輯參考 a200 / a201。
"""

# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import re
import sys
from pathlib import Path
from typing import Optional

# 載入第三方套件
import xlwings as xw

# 載入自訂模組/函式
from mod_logging import (
    init_logging,
    logging_exc_error,
    logging_exception,
    logging_process_step,
    logging_warning,
)
from mod_標音 import (
    PiauIm,
    convert_tl_with_tiau_hu_to_tlpa,
    im_piau_iong_tiau_ho,
    is_han_ji,
    is_punctuation,
    split_tai_gi_im_piau,
)
from mod_程式 import ExcelCell, Program

# =========================================================================
# 常數定義
# =========================================================================
EXIT_CODE_SUCCESS = 0
EXIT_CODE_NO_FILE = 1
EXIT_CODE_INVALID_INPUT = 2
EXIT_CODE_SAVE_FAILURE = 3
EXIT_CODE_PROCESS_FAILURE = 10
EXIT_CODE_UNKNOWN_ERROR = 99

# 變調標示底色（淺灰）
BIAN_TIAU_FILL_COLOR = (232, 232, 232)  # #E8E8E8

# 上標數字
SUPERSCRIPT_DIGITS = str.maketrans("0123456789", "⁰¹²³⁴⁵⁶⁷⁸⁹")
SUPERSCRIPT_MINUS = "\u207b"  # ⁻

# 十五音／雅俗通用中文調號
CHINESE_TONE_MARKS = {
    1: "一",
    2: "二",
    3: "三",
    4: "四",
    5: "五",
    7: "七",
    8: "八",
}

# 台羅調號 → 閩拼調號
TL_TO_BP_TONE = {
    0: 0,
    1: 1,  # 陰平
    2: 3,  # 上聲
    3: 5,  # 陰去
    4: 7,  # 陰入
    5: 2,  # 陽平
    7: 6,  # 陽去
    8: 8,  # 陽入
}

# 舒聲變調（台羅調號）
SANDHI_SHU_SIANN = {
    1: 7,  # 高變中
    2: 1,  # 高降取高
    3: 2,  # 低變高降
    5: 7,  # 低升取升／取中
    7: 3,  # 中變低
}

# 使用閩拼調號系統顯示之本調／變調標音法
BP_TONE_PIAU_IM_HUAT = {"閩拼調號", "閩拼調符", "閩拼注音"}

# 方音符號齒音連 i 特殊轉換（與 TPS_piau_im 相同）
TPS_REMAP = {
    "ㄗㄧ": "ㄐㄧ",
    "ㄘㄧ": "ㄑㄧ",
    "ㄙㄧ": "ㄒㄧ",
    "ㆡㄧ": "ㆢㄧ",
}

# =========================================================================
# 設定日誌
# =========================================================================
init_logging()


# =========================================================================
# 變調／本調格式輔助函式
# =========================================================================
def to_superscript(tone: int) -> str:
    """將調號轉為上標數字。"""
    return str(tone).translate(SUPERSCRIPT_DIGITS)


def normalize_tl_tone(tiau_ho: str | int) -> int:
    """正規化台羅調號（6 → 7）。"""
    tiau = int(str(tiau_ho).strip())
    return 7 if tiau == 6 else tiau


def get_sandhi_tone_tl(tiau_ho: int, un_bu: str) -> int:
    """
    依台羅八音變調規則，求變調後調號。

    舒聲：1→7、2→1、3→2、5→7、7→3
    入聲 -p/t/k：4→8、8→4
    入聲 -h：4→2、8→3
    """
    if tiau_ho in SANDHI_SHU_SIANN:
        return SANDHI_SHU_SIANN[tiau_ho]

    if tiau_ho in (4, 8):
        if un_bu.endswith(("p", "t", "k")):
            return {4: 8, 8: 4}[tiau_ho]
        # 入聲韻尾 h／ㆷ
        return {4: 2, 8: 3}[tiau_ho]

    return tiau_ho


def to_display_tone(tl_tone: int, piau_im_huat: str) -> int:
    """依標音法，將台羅調號轉為顯示用調號。"""
    if piau_im_huat in BP_TONE_PIAU_IM_HUAT:
        return TL_TO_BP_TONE.get(tl_tone, tl_tone)
    return tl_tone


def is_bian_tiau_fill(cell) -> bool:
    """判斷漢字儲存格是否塗成變調標示淺灰底色。"""
    return cell.color == BIAN_TIAU_FILL_COLOR


def _normalize_siann_bu(siann_bu: Optional[str]) -> str:
    if siann_bu in ("", None, "Ø", "ø"):
        return ""
    return siann_bu


def get_syllable_body(
    piau_im: PiauIm,
    piau_im_huat: str,
    siann_bu: str,
    un_bu: str,
    tiau_ho: str,
) -> str:
    """
    取得不含調號的音節主體（本調／變調格式共用）。

    十五音、雅俗通之本調／變調皆採十五音順序：聲＋韻。
    """
    siann_bu = _normalize_siann_bu(siann_bu)

    if piau_im_huat in BP_TONE_PIAU_IM_HUAT:
        result = piau_im._get_BP_syllable(
            siann_bu, un_bu, tiau_ho, with_tone_number=False
        )
        if not result or not isinstance(result, tuple):
            return ""
        return f"{result[0]}{result[1]}"

    if piau_im_huat in ("十五音", "雅俗通"):
        if siann_bu == "":
            siann = "英"
        else:
            siann = piau_im.Siann_Bu_Dict[siann_bu]["十五音"]
        un = piau_im.Un_Bu_Dict[un_bu]["十五音"]
        return f"{siann}{un}"

    if piau_im_huat == "方音符號":
        key_siann = "ø" if siann_bu == "" else siann_bu
        siann = piau_im.Siann_Bu_Dict[key_siann]["方音符號"]
        un = piau_im.Un_Bu_Dict[un_bu]["方音符號"]
        body = f"{siann}{un}"
        for src, dst in TPS_REMAP.items():
            body = body.replace(src, dst)
        return body

    if piau_im_huat in ("注音二式", "台語注音二式"):
        mps2 = piau_im.MPS2_piau_im(siann_bu, un_bu, tiau_ho)
        return re.sub(r"\d+$", "", mps2 or "")

    if piau_im_huat == "台語音標":
        dict_key = "台語音標"
    elif piau_im_huat == "台羅拼音":
        dict_key = "台羅拼音"
    elif piau_im_huat == "白話字":
        dict_key = "白話字"
    else:
        logging_warning(f"尚未支援之標音方法本調／變調格式：{piau_im_huat}")
        return ""

    if siann_bu == "":
        siann = ""
    else:
        siann = piau_im.Siann_Bu_Dict[siann_bu][dict_key] or ""
    un = piau_im.Un_Bu_Dict[un_bu][dict_key]
    return f"{siann}{un}"


def format_bun_tiau(
    body: str,
    display_tone: int,
    piau_im_huat: str,
) -> str:
    """組出本調標音格式。"""
    if piau_im_huat in ("十五音", "雅俗通"):
        return f"{body}{CHINESE_TONE_MARKS[display_tone]}"
    return f"{body}{to_superscript(display_tone)}"


def format_bian_tiau(
    body: str,
    display_orig: int,
    display_sandhi: int,
    piau_im_huat: str,
) -> str:
    """組出變調標音格式。"""
    if piau_im_huat in ("十五音", "雅俗通"):
        # 十五音／雅俗通使用半形連字號
        return (
            f"{body}{CHINESE_TONE_MARKS[display_orig]}"
            f"-{CHINESE_TONE_MARKS[display_sandhi]}"
        )
    return (
        f"{body}{to_superscript(display_orig)}"
        f"{SUPERSCRIPT_MINUS}{to_superscript(display_sandhi)}"
    )


def convert_to_bun_or_bian_tiau_format(
    piau_im: PiauIm,
    piau_im_huat: str,
    tai_gi_im_piau: str,
    use_bian_tiau: bool,
) -> str:
    """
    將台語音標轉換為本調或變調標音格式。

    Args:
        piau_im: 標音轉換物件
        piau_im_huat: 漢字標音法（env【標音方法】）
        tai_gi_im_piau: 台語音標（TLPA+）
        use_bian_tiau: True=變調格式；False=本調格式
    """
    if not tai_gi_im_piau or str(tai_gi_im_piau).strip() == "":
        return ""

    im_piau = str(tai_gi_im_piau).strip()
    if not im_piau_iong_tiau_ho(im_piau):
        im_piau = convert_tl_with_tiau_hu_to_tlpa(im_piau)

    siann_bu, un_bu, tiau_ho = split_tai_gi_im_piau(im_piau)
    tl_tone = normalize_tl_tone(tiau_ho)
    sandhi_tone = get_sandhi_tone_tl(tl_tone, un_bu)

    body = get_syllable_body(
        piau_im=piau_im,
        piau_im_huat=piau_im_huat,
        siann_bu=siann_bu,
        un_bu=un_bu,
        tiau_ho=str(tl_tone),
    )
    if not body:
        return ""

    display_orig = to_display_tone(tl_tone, piau_im_huat)
    display_sandhi = to_display_tone(sandhi_tone, piau_im_huat)

    if use_bian_tiau:
        return format_bian_tiau(body, display_orig, display_sandhi, piau_im_huat)
    return format_bun_tiau(body, display_orig, piau_im_huat)


# =========================================================================
# 資料類別：儲存格處理器
# =========================================================================
class CellProcessor(ExcelCell):
    """將漢字標音改寫為本調／變調標示格式。"""

    def __init__(
        self,
        program: Program,
        new_jin_kang_piau_im_ji_khoo_sheet: bool = False,
        new_piau_im_ji_khoo_sheet: bool = False,
        new_khuat_ji_piau_sheet: bool = False,
    ):
        super().__init__(
            program=program,
            new_jin_kang_piau_im_ji_khoo_sheet=new_jin_kang_piau_im_ji_khoo_sheet,
            new_piau_im_ji_khoo_sheet=new_piau_im_ji_khoo_sheet,
            new_khuat_ji_piau_sheet=new_khuat_ji_piau_sheet,
        )

    def _process_non_han_ji(self, cell) -> str:
        """標點／空白：標點字元略過，不寫入漢字標音列。"""
        cell_value = cell.value
        if cell_value is None or str(cell_value).strip() == "":
            return "【空白字元】"

        str_value = str(cell_value).strip()

        if is_punctuation(str_value):
            return f"【標點符號】略過：{str_value}"

        return f"【非漢字】：{str_value}"

    def _process_han_ji(self, cell) -> str:
        """依漢字底色，將漢字標音改為本調或變調格式。"""
        han_ji = cell.value
        tai_gi_im_piau_cell = cell.offset(-1, 0)
        han_ji_piau_im_cell = cell.offset(1, 0)

        tai_gi_im_piau = tai_gi_im_piau_cell.value
        if not tai_gi_im_piau or str(tai_gi_im_piau).strip() == "":
            return f"【{han_ji}】台語音標為空，略過。"

        use_bian_tiau = is_bian_tiau_fill(cell)
        piau_im_huat = self.program.piau_im_huat

        try:
            new_piau_im = convert_to_bun_or_bian_tiau_format(
                piau_im=self.program.piau_im,
                piau_im_huat=piau_im_huat,
                tai_gi_im_piau=str(tai_gi_im_piau).strip(),
                use_bian_tiau=use_bian_tiau,
            )
        except Exception as e:
            logging_exception(
                msg=f"轉換【{han_ji}】本調／變調格式失敗（台語音標：{tai_gi_im_piau}）",
                error=e,
            )
            return f"【{han_ji}】轉換失敗：{e}"

        if not new_piau_im:
            return f"【{han_ji}】無法產生本調／變調格式（台語音標：{tai_gi_im_piau}）。"

        han_ji_piau_im_cell.value = new_piau_im
        fmt_name = "變調" if use_bian_tiau else "本調"
        return f"【{han_ji}】[{tai_gi_im_piau}] → {new_piau_im}（{fmt_name}）"

    def _process_cell(self, cell) -> int:
        """
        處理單一漢字儲存格。

        Returns:
            0=漢字、1=文章終結、2=換行、3=空白／標點
        """
        cell_value = cell.value

        if cell_value == "φ":
            print("【文字終結】")
            return 1
        if cell_value == "\n":
            print("【換行】")
            return 2
        if cell_value is None or str(cell_value).strip() == "":
            print("【空白】")
            return 3
        if not is_han_ji(cell_value):
            msg = self._process_non_han_ji(cell)
            print(msg)
            return 3

        msg = self._process_han_ji(cell)
        print(msg)
        return 0

    def _process_sheet(self, sheet) -> None:
        """掃描整張【漢字注音】工作表。"""
        config = self.program
        total_lines = config.TOTAL_LINES
        rows_per_line = config.ROWS_PER_LINE
        line_start_row = config.line_start_row
        start_col = config.start_col
        end_col = config.end_col
        han_ji_row_offset = config.han_ji_row_offset

        active_cell = sheet.range(f"{xw.utils.col_name(start_col)}{line_start_row}")
        active_cell.select()

        is_eof = False
        for line_no in range(1, total_lines + 1):
            if is_eof or line_no > total_lines:
                break

            self._show_separtor_line(f"處理第 {line_no} 行...")
            row = line_start_row + (line_no - 1) * rows_per_line + han_ji_row_offset

            for col in range(start_col, end_col + 1):
                active_cell = sheet.range((row, col))
                active_cell.select()

                print("-" * 80)
                print(f"儲存格：{xw.utils.col_name(col)}{row}（{row}, {col}）")

                status_code = self._process_cell(active_cell)
                if status_code == 1:
                    is_eof = True
                    break
                if status_code == 2:
                    break


def process(wb, args) -> int:
    """將漢字標音改為本調／變調標示格式。"""
    logging_process_step("<=========== 作業開始！==========>")
    sheet_name = "漢字注音"

    try:
        program = Program(wb, args, hanji_piau_im_sheet_name=sheet_name)
        logging_process_step(f"【標音方法】：{program.piau_im_huat}")

        xls_cell = CellProcessor(
            program=program,
            new_jin_kang_piau_im_ji_khoo_sheet=False,
            new_piau_im_ji_khoo_sheet=False,
            new_khuat_ji_piau_sheet=False,
        )

        sheet = wb.sheets[sheet_name]
        sheet.activate()
        xls_cell._process_sheet(sheet=sheet)

        print("=" * 80)
        logging_process_step(msg="已完成【漢字標音】本調／變調標示格式轉換。")
        return EXIT_CODE_SUCCESS

    except Exception as e:
        logging_exception(
            msg=f"在【{sheet_name}】工作表改寫本調／變調標音格式時發生例外！",
            error=e,
        )
        raise


def main(args) -> int:
    """主程式 - 從 Excel 呼叫或直接執行。"""
    current_file_path = Path(__file__).resolve()
    project_root = current_file_path.parent
    program_name = current_file_path.stem

    logging_process_step(f"《========== 程式開始執行：{program_name} ==========》")
    logging_process_step(f"專案根目錄為: {project_root}")

    wb = None
    try:
        wb = xw.Book.caller()
    except Exception:
        try:
            wb = xw.apps.active.books.active
        except Exception as e:
            logging_exc_error(msg="無法找到作用中的 Excel 工作簿！", error=e)
            return EXIT_CODE_NO_FILE

    if not wb:
        logging_exc_error(msg="無法取得 Excel 活頁簿！", error=None)
        return EXIT_CODE_NO_FILE

    try:
        exit_code = process(wb, args)
    except Exception as e:
        logging_exception(msg=f"作業程序發生異常，終止執行：{program_name}", error=e)
        return EXIT_CODE_PROCESS_FAILURE

    if exit_code != EXIT_CODE_SUCCESS:
        logging_exc_error(
            msg=f"處理作業發生異常，終止程式執行：{program_name}",
            error=None,
        )
        return EXIT_CODE_PROCESS_FAILURE

    try:
        if not Program.save_workbook_as_new_file(wb=wb):
            return EXIT_CODE_SAVE_FAILURE
    except Exception as e:
        logging_exception(msg="儲存檔案失敗！", error=e)
        return EXIT_CODE_SAVE_FAILURE

    print("\n")
    print("=" * 80)
    logging_process_step(f"《========== 程式終止執行：{program_name} ==========》")
    return EXIT_CODE_SUCCESS


def test_format_conversion() -> None:
    """單元測試：對照說明文件之閩拼變調／本調範例。"""
    piau_im = PiauIm("河洛話")
    piau_im_huat = "閩拼調符"

    expected = {
        ("tuan3", True): "duan⁵⁻³",
        ("hong5", False): "hong²",
        ("ze3", True): "ze⁵⁻³",
        ("u2", False): "wu³",
        ("zing7", True): "zing⁶⁻⁵",
        ("ciu1", True): "ciu¹⁻⁶",
        ("khong1", False): "kong¹",
        ("san1", True): "san¹⁻⁶",
        ("liam2", False): "liam³",
        ("siu1", True): "siu¹⁻⁶",
        ("bi5", False): "bbi²",
        ("sin1", True): "sin¹⁻⁶",
        ("liok8", False): "liok⁸",
    }

    print("=" * 70)
    print("測試本調／變調標音格式轉換（閩拼調符）")
    print("=" * 70)

    failed = 0
    for (tl, use_bian), want in expected.items():
        got = convert_to_bun_or_bian_tiau_format(
            piau_im=piau_im,
            piau_im_huat=piau_im_huat,
            tai_gi_im_piau=tl,
            use_bian_tiau=use_bian,
        )
        fmt = "變調" if use_bian else "本調"
        ok = got == want
        if not ok:
            failed += 1
        print(f"{'OK' if ok else 'NG'} {tl:8} ({fmt}) → {got!r}  期望 {want!r}")

    # 十五音格式抽樣
    sni_bian = convert_to_bun_or_bian_tiau_format(
        piau_im, "十五音", "tuan3", use_bian_tiau=True
    )
    sni_bun = convert_to_bun_or_bian_tiau_format(
        piau_im, "十五音", "tuan3", use_bian_tiau=False
    )
    print(f"十五音 本調 tuan3 → {sni_bun!r}（期望含「地觀三」）")
    print(f"十五音 變調 tuan3 → {sni_bian!r}（期望含「地觀三-二」）")

    print("=" * 70)
    if failed:
        print(f"測試失敗：{failed} 筆")
        raise SystemExit(1)
    print("測試完成：全部通過")


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description="將【漢字標音】改寫為本調／變調標示格式",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用範例：
  python c100_漢字標音儲存格改用變調標示格式.py
  python c100_漢字標音儲存格改用變調標示格式.py --test
""",
    )
    parser.add_argument(
        "--test",
        action="store_true",
        help="執行格式轉換單元測試",
    )
    args = parser.parse_args()

    if args.test:
        test_format_conversion()
    else:
        exit_code = main(args)
        if exit_code != EXIT_CODE_SUCCESS:
            print(f"程式異常終止，錯誤代碼為: {exit_code}")
            sys.exit(exit_code)
