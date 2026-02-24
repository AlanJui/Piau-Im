"""
a940_自Excel轉製html檔.py v0.0.4

功能：
    參考 a400_製作標音網頁.py 之作法，將 Excel 檔中的【漢字標音】（即：雅俗通十五音）
    轉輸出成類似風格的 HTML 檔。

    預期 Excel 格式：
    Column A: 漢字 / 標點符號 / 換行符號(\n)
    Column B: 漢字標音 (若為空，則視為標點符號，不加 ruby)

    輸出：
    docs/ [檔名].html
"""

import logging
import sys
from pathlib import Path

import xlwings as xw

from mod_logging import (
    logging_exc_error,
    logging_exception,
    logging_process_step,
    logging_warning,
)
from mod_標音 import format_han_ji_piau_im
from mod_程式 import Program

# 嘗試載入 mod_標音
try:
    from mod_標音 import PiauIm
except ImportError:
    PiauIm = None
    split_tai_gi_im_piau = None
    print("警告：無法載入 mod_標音")

# 十五音聲調對照表
TONE_MAP = {
    "一": "上平",
    "二": "上上",
    "三": "上去",
    "四": "上入",
    "五": "下平",
    "六": "下上",
    "七": "下去",
    "八": "下入",
}

# =========================================================================
# 常數定義
# =========================================================================
# 定義 Exit Code
EXIT_CODE_SUCCESS = 0  # 成功
EXIT_CODE_NO_FILE = 1  # 無法找到檔案
EXIT_CODE_INVALID_INPUT = 2  # 輸入錯誤
EXIT_CODE_SAVE_FAILURE = 3  # 儲存失敗
EXIT_CODE_PROCESS_FAILURE = 10  # 過程失敗
EXIT_CODE_UNKNOWN_ERROR = 99  # 未知錯誤


def tai_gi_im_piau_tng_huan(
    siann: str, un: str, tiau: str, piau_im_huat: str, piau_im_object
) -> str:
    """根據【標音方式】設定，將【台語音標】切分成聲母、韻母、調號後，轉換成對映的【漢字標音】。

    args:
        siann: 台語音標的聲母部分 (例如 "k")
        un: 台語音標的韻母部分 (例如 "ian")
        tiau: 台語音標的調號部分 (例如 "5")
        piau_im_huat: 標音標準（反切/注音/羅馬拼音）設定 (例如 "台語音標")
    """
    # 確保 siann, un, tiau 不為 None，若為 None 則轉為空字串
    siann = siann if siann is not None else ""
    un = un if un is not None else ""
    tiau = tiau if tiau is not None else ""

    # 將【聲母】、【韻母】、【聲調】，合併成【台語音標】
    tai_gi_im_piau = f"{siann}{un}{tiau}"

    # 標音法為：【十五音】或【雅俗通】，且【聲母】為空值，則將【聲母】設為【ø】
    if (piau_im_huat == "十五音" or piau_im_huat == "雅俗通") and (
        siann == "" or siann is None
    ):
        siann = "ø"

    ok = False
    han_ji_piau_im = ""
    try:
        han_ji_piau_im = piau_im_object.han_ji_piau_im_tng_huan(
            piau_im_huat=piau_im_huat,
            siann_bu=siann,
            un_bu=un,
            tiau_ho=tiau,
        )
        if han_ji_piau_im:  # 傳回非空字串，表示【漢字標音】之轉換成功
            ok = True
        else:
            logging_warning(
                f"【台語音標】：[{tai_gi_im_piau}]，轉換成【{piau_im_huat}漢字標音】拚音/注音系統失敗！"
            )
    except Exception as e:
        logging_exception(
            f"piau_im.han_ji_piau_im_tng_huan() 發生執行時期錯誤: 【台語音標】：{tai_gi_im_piau}",
            e,
        )
        han_ji_piau_im = ""
        ok = False

    # 若 ok 為 False，表示轉換失敗，則將【台語音標】直接傳回
    if not ok:
        return tai_gi_im_piau
    else:
        return format_han_ji_piau_im(han_ji_piau_im)


def export_excel_to_html(program, output_path):
    # 連接 Excel
    try:
        wb = program.wb

        # (1) 預設使用【網頁匯入】工作表
        try:
            source_sheet_name = program.hanji_piau_im_sheet_name
            sheet = wb.sheets[source_sheet_name]
        except Exception:
            sheet = wb.sheets.active
            print(f"找無【{source_sheet_name}】，使用: {sheet.name}")

        # --------------------------------------------------------------------------
        # 初始化 process config
        # --------------------------------------------------------------------------
        # program = Program(wb, args, hanji_piau_im_sheet_name="網頁匯入")

        # 嘗試取得網頁標題
        try:
            title = program.title
            if title is None:
                title = sheet.name
        except Exception:
            title = sheet.name

        # 取得標音方式設定
        han_ji_piau_im_hong_sik = program.han_ji_piau_im_format  # 標音方式
        siong_pinn_piau_im = program.siong_pinn_piau_im  # 上邊標音
        zian_pinn_piau_im = program.zian_pinn_piau_im  # 右邊標音

        # 去除前後空白
        if han_ji_piau_im_hong_sik:
            han_ji_piau_im_hong_sik = str(han_ji_piau_im_hong_sik).strip()
        if siong_pinn_piau_im:
            siong_pinn_piau_im = str(siong_pinn_piau_im).strip()
        if zian_pinn_piau_im:
            zian_pinn_piau_im = str(zian_pinn_piau_im).strip()

        # 若未設定，給予預設值
        if han_ji_piau_im_hong_sik is None:
            han_ji_piau_im_hong_sik = "上邊"
        if siong_pinn_piau_im is None:
            siong_pinn_piau_im = "台語音標"
        if zian_pinn_piau_im is None:
            zian_pinn_piau_im = ""

    except Exception as e:
        print("無法連接到 Excel。請確認 Excel 已開啟且有活動工作簿。")
        print(f"錯誤訊息: {e}")
        return

    if PiauIm:  # 確保 PiauIm 已成功載入
        try:
            piau_im_object = program.piau_im
        except Exception as e:
            print(f"無法初始化 PiauIm 物件: {e}")

    # --------------------------------------------------------------------------
    # 讀取資料
    # 從 A2 開始讀取，並嘗試讀取到 F 欄
    try:
        last_row = sheet.range("A" + str(sheet.cells.last_cell.row)).end("up").row
        if last_row < 2:
            print("Excel 無資料 (至少需要有一列資料)。")
            return

        # 讀取 A, B, C 欄
        # A: 漢字
        # B: 漢字標音
        # C: 台語音標
        # D: 台語音標【聲母】
        # E: 台語音標【韻母】
        # F: 台語音標【調號】
        data = sheet.range(f"A2:F{last_row}").value
    except Exception as e:
        print(f"讀取 Excel 資料失敗: {e}")
        return

    # 建構內容 HTML
    content_lines = []

    # 初始區塊
    # content_lines.append('<div class="fifteen_yin">')
    content_lines.append('<div class="Siang_Pai">')

    # 標題
    content_lines.append(f'<p class="title"><span>《</span>{title}<span>》</span></p>')

    # 內容開始
    in_paragraph = False

    def start_paragraph_if_needed():
        nonlocal in_paragraph
        if not in_paragraph:
            content_lines.append("<p>")
            in_paragraph = True

    def end_paragraph_if_needed():
        nonlocal in_paragraph
        if in_paragraph:
            content_lines.append("</p>")
            in_paragraph = False

    # 強制開始第一段
    start_paragraph_if_needed()

    for row in data:
        if row is None:
            continue

        # 確保 row 為 list
        if not isinstance(row, list):
            # 單欄情況(雖然這裡是 A:F)
            row = [row]

        # 補足長度
        # while len(row) < 3:
        while len(row) < 6:
            row.append(None)

        han_ji = row[0]
        sip_goo_im_piau_im = row[1]
        tai_gi_im_piau = row[2]

        # 過濾無效資料，務必需有：漢字、漢字標音及至台語音標
        if han_ji is None:
            han_ji = ""
        if sip_goo_im_piau_im is None:
            sip_goo_im_piau_im = ""
        if tai_gi_im_piau is None:
            tai_gi_im_piau = ""

        han_ji = str(han_ji)
        sip_goo_im_piau_im = str(sip_goo_im_piau_im)
        tai_gi_im_piau = str(tai_gi_im_piau)

        # 若【漢字】儲存格，其內容為：換行控制符
        if han_ji == "\\n" or han_ji == "\\r\\n" or han_ji == "\n" or han_ji == "\r\n":
            end_paragraph_if_needed()
            start_paragraph_if_needed()
            continue

        # 轉換標音
        # 根據設定 (siong_pinn_piau_im, zian_pinn_piau_im) 轉換內容
        # 預設邏輯：
        # - B欄 (piau_im_1) 為主要標音來源 (通常是十五音代碼，如 "堅五曾")
        # - 若有 C欄 (piau_im_2)，則視為第二標音來源 (若未留空)

        # 取得【台語音標】的【聲母】、【韻母】、【調號】
        siann = row[3]
        un = row[4]
        tiau = int(row[5]) if row[5] is not None else ""  # 調號可能為空值，需處理

        # 準備上邊標音內容
        top_content = ""
        if "上" in han_ji_piau_im_hong_sik and tai_gi_im_piau:
            converted_1 = tai_gi_im_piau_tng_huan(
                siann=siann,
                un=un,
                tiau=tiau,
                piau_im_huat=siong_pinn_piau_im,
                piau_im_object=piau_im_object,
            )
            top_content = converted_1

        # 準備右邊標音內容
        right_content = ""
        if "右" in han_ji_piau_im_hong_sik and tai_gi_im_piau:
            right_content = tai_gi_im_piau_tng_huan(
                siann=siann,
                un=un,
                tiau=tiau,
                piau_im_huat=zian_pinn_piau_im,
                piau_im_object=piau_im_object,
            )

        # 組合 HTML
        # <ruby>
        #   <rb>漢字</rb>
        #   <rp>(</rp><rt>上邊</rt><rp>)</rp>
        #   <rtc>右邊</rtc>
        # </ruby>

        has_top = bool(top_content and top_content.strip())
        has_right = bool(right_content and right_content.strip())

        if not has_top and not has_right:
            if han_ji.strip() == "":
                content_lines.append("  <span>　</span>")
            else:
                content_lines.append(f"  <span>{han_ji}</span>")
        else:
            ruby_parts = [f"<ruby><rb>{han_ji}</rb>"]

            if has_top:
                ruby_parts.append(f"<rp>(</rp><rt>{top_content}</rt><rp>)</rp>")

            if has_right:
                ruby_parts.append(f"<rtc>{right_content}</rtc>")

            ruby_parts.append("</ruby>")
            content_lines.append("".join(ruby_parts))

    end_paragraph_if_needed()
    content_lines.append("</div>")

    content_html = "\n".join(content_lines)

    web_page_main_file_name = Path(output_path).stem
    full_image_url = "https://alanjui.github.io/Piau-Im/assets/images/king_tian.png"

    html_template = f"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{title}</title>
    <meta content='https://alanjui.github.io/Piau-Im/{web_page_main_file_name}.html' property='og:url' />
    <meta content='{title}' property='og:title' />
    <meta content='{title}' property='og:description' />
    <meta content='{full_image_url}' property='og:image' />
    <link rel="stylesheet" href="assets/styles/styles.css">
</head>
<body>
    <main class="page">
        <article class="article_content">
        {content_html}
        </article>
    </main>
</body>
</html>
"""

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)

    try:
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(html_template)
        print(f"成功輸出 HTML 至: {output_path}")
    except Exception as e:
        print(f"寫入檔案失敗: {e}")

    return EXIT_CODE_SUCCESS


def process(wb, args) -> int:
    """
    為【漢字】之【漢字標音】，以批次作業方式，完成各種標音方法標注。

    Args:
        wb: Excel Workbook 物件

    Returns:
        處理結果代碼
    """
    # --------------------------------------------------------------------------
    # 作業初始化
    # --------------------------------------------------------------------------
    logging_process_step("<=========== 作業開始！==========>")

    try:
        # 初始化 process config
        program = Program(wb, args, hanji_piau_im_sheet_name="網頁匯入")

        # # 建立儲存格處理器
        # if args.new:
        #     xls_cell = CellProcessor(
        #         program=program,
        #         new_jin_kang_piau_im_ji_khoo_sheet=True,
        #         new_piau_im_ji_khoo_sheet=True,
        #         new_khuat_ji_piau_sheet=True,
        #     )
        # else:
        #     xls_cell = CellProcessor(
        #         program=program,
        #         new_jin_kang_piau_im_ji_khoo_sheet=False,
        #         new_piau_im_ji_khoo_sheet=False,
        #         new_khuat_ji_piau_sheet=False,
        #     )
    except Exception as e:
        logging_exc_error(msg="初始化作業，發生執行異常！", error=e)
        return EXIT_CODE_PROCESS_FAILURE

    # ------------------------------------------------------------------------------
    # 處理作業
    # ------------------------------------------------------------------------------
    try:
        output_file = args.output_file
        export_excel_to_html(program, output_file)
    except Exception as e:
        logging_exception(
            msg=f"程式：{program.program_name} ，執行時發生異常問題！",
            error=e,
        )
        raise

    # ------------------------------------------------------------------------------
    # 處理作業結束
    # ------------------------------------------------------------------------------
    print("=" * 80)
    logging_process_step("<=========== 作業結束！==========>")
    return EXIT_CODE_SUCCESS


# =========================================================================
# 主程式
# =========================================================================
def main(args) -> int:
    """主程式"""
    # =========================================================================
    # (0) 程式初始化
    # =========================================================================
    # 取得專案根目錄。
    current_file_path = Path(__file__).resolve()
    project_root = current_file_path.parent
    # 取得程式名稱
    program_name = current_file_path.stem

    # =========================================================================
    # (1) 開始執行程式
    # =========================================================================
    logging_process_step(f"《========== 程式開始執行：{program_name} ==========》")
    logging_process_step(f"專案根目錄為: {project_root}")

    # =========================================================================
    # (2) 設定【作用中活頁簿】：偵測及獲取 Excel 已開啟之活頁簿檔案。
    # =========================================================================
    try:
        # 取得 Excel 活頁簿
        wb = None
        wb = xw.apps.active.books.active
    except Exception as e:
        logging.error(f"無法找到作用中的 Excel 工作簿: {e}")
        return EXIT_CODE_NO_FILE

    # 若無法取得【作用中活頁簿】，則因無法繼續作業，故返回【作業異常終止代碼】結束。
    if not wb:
        logging_exc_error(msg="無法取得 Excel 活頁簿！", error=None)
        return EXIT_CODE_NO_FILE

    # =========================================================================
    # (3) 執行【處理作業】
    # =========================================================================
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

    # =========================================================================
    # (4) 儲存檔案
    # =========================================================================
    try:
        # 儲存檔案
        if not Program.save_workbook_as_new_file(wb=wb):
            return EXIT_CODE_SAVE_FAILURE  # 作業異當終止：無法儲存檔案
    except Exception as e:
        logging_exception(msg="儲存檔案失敗！", error=e)
        return EXIT_CODE_SAVE_FAILURE  # 作業異當終止：無法儲存檔案

    # =========================================================================
    # (5) 結束程式
    # =========================================================================
    print("\n")
    print("=" * 80)
    logging_process_step(f"《========== 程式終止執行：{program_name} ==========》")
    return EXIT_CODE_SUCCESS


# =========================================================================
# 單元測試
# =========================================================================
def test_01():
    pass


# =========================================================================
# 程式入口
# =========================================================================
if __name__ == "__main__":
    import argparse
    import sys

    # 解析命令行參數
    parser = argparse.ArgumentParser(
        description="a940_自Excel轉製html檔.py",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用範例：
  python a940.py                                # 執行一般模式
  python a940.py --output_file <output_file>    # 建立新的字庫工作表
  python a940.py --test                         # 執行測試模式
""",
    )
    parser.add_argument(
        "--test",
        action="store_true",
        help="執行測試模式",
    )

    # 預設輸出到 docs 目錄
    output_file = str(Path("docs") / "output_from_excel.html")
    parser.add_argument(
        "output_file",
        nargs="?",
        default=output_file,
        help="輸出 HTML 檔案的路徑 (預設: docs/output_from_excel.html)",
    )
    args = parser.parse_args()

    if args.test:
        # 執行測試
        test_01()
    else:
        # 從 Excel 呼叫
        # exit_code = export_excel_to_html(args, output_file)
        exit_code = main(args)
        if exit_code != EXIT_CODE_SUCCESS:
            print(f"程式異常終止，返回代碼：{exit_code}")
            sys.exit(exit_code)
