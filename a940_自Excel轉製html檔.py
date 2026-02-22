"""
a940_自Excel轉製html檔.py

功能：
    參考 a400_製作標音網頁.py 之作法，將 Excel 檔中的【漢字標音】（即：雅俗通十五音）
    轉輸出成類似風格的 HTML 檔。

    預期 Excel 格式：
    Column A: 漢字 / 標點符號 / 換行符號(\n)
    Column B: 漢字標音 (若為空，則視為標點符號，不加 ruby)

    輸出：
    docs/ [檔名].html
"""

import sqlite3
import sys
from pathlib import Path

import xlwings as xw

from mod_excel_access import get_value_by_name

# 嘗試載入 mod_標音
try:
    from mod_標音 import PiauIm, split_tai_gi_im_piau
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


def convert_15yin(raw, system, p_im, cursor_15yin):
    if not p_im or not cursor_15yin or len(raw) != 3:
        return raw
    try:
        t = TONE_MAP.get(raw[1])
        if not t:
            return raw

        # Query: 堅(韻 Rhyme) + 五(調 Tone) + 曾(聲 Initial)
        # DB Columns: 字韻, 聲調, 切音, 漢字標音
        cursor_15yin.execute(
            "SELECT 漢字標音 FROM 漢字表 WHERE 字韻=? AND 聲調=? AND 切音=? LIMIT 1",
            (raw[0], t, raw[2]),
        )
        row = cursor_15yin.fetchone()
        if not row:
            return raw

        tg = row[0]  # zian5
        if not split_tai_gi_im_piau:
            return tg

        parts = split_tai_gi_im_piau(tg)
        s, u, ti = parts[0], parts[1], parts[2]

        # 針對「英」聲母 (q) 進行處理
        # 如果切分後的韻母 u 以 'q' 開頭 (例如 'qiu', 'qu')，
        # 表示這是以 q 代表零聲母的情況
        if not s and u.startswith("q"):
            s = "Ø"  # 對應 DB 中的零聲母 Key (大寫 Ø)
            u = u[1:]  # 去掉 q，剩下真正的韻母 (e.g. 'iu')

        # 如果 s 是空字串，也要轉成 'Ø'
        if not s:
            s = "Ø"

        # 確保 s 是大寫 Ø (因為 DB 裡是大寫，但 split 可能轉小寫了)
        if s == "ø":
            s = "Ø"

        # 使用 Ø 查詢對應表，如果系統是輸出台羅拼音等，通常 Ø 對應的就是空字串
        # 但如果是直接輸出【台語音標】(即 key 值)，可能會印出 Ø
        # 所以先檢查轉換系統是否是【台語音標】
        conv = p_im.han_ji_piau_im_tng_huan(system, s, u, ti)

        # 如果轉換結果包含 Ø，把它換成空字串 (這是使用者的要求)
        if conv and "Ø" in conv:
            conv = conv.replace("Ø", "")

        return conv if conv else tg

        conv = p_im.han_ji_piau_im_tng_huan(system, s, u, ti)
        return conv if conv else tg
    except Exception as e:
        print(f"Error converting {raw}: {e}")
        return raw


def export_excel_to_html(output_path):
    # 連接 Excel
    try:
        wb = xw.books.active

        # (1) 預設使用【網頁匯入】工作表
        try:
            sheet = wb.sheets["網頁匯入"]
        except Exception:
            sheet = wb.sheets.active
            print(f"找無【網頁匯入】，使用: {sheet.name}")

        # 嘗試取得網頁標題
        try:
            title = get_value_by_name(wb, "TITLE")
            if title is None:
                title = sheet.name
        except Exception:
            title = sheet.name

        # 取得標音方式設定
        # 邏輯參考 a400，讀取設定值
        piau_im_hong_sik = get_value_by_name(wb, "標音方式")
        siong_pinn_piau_im = get_value_by_name(wb, "上邊標音")
        zian_pinn_piau_im = get_value_by_name(wb, "右邊標音")

        # 去除前後空白
        if piau_im_hong_sik:
            piau_im_hong_sik = str(piau_im_hong_sik).strip()
        if siong_pinn_piau_im:
            siong_pinn_piau_im = str(siong_pinn_piau_im).strip()
        if zian_pinn_piau_im:
            zian_pinn_piau_im = str(zian_pinn_piau_im).strip()

        # 若未設定，給予預設值
        if piau_im_hong_sik is None:
            piau_im_hong_sik = "上邊"
        if siong_pinn_piau_im is None:
            siong_pinn_piau_im = "台語音標"
        if zian_pinn_piau_im is None:
            zian_pinn_piau_im = ""

    except Exception as e:
        print("無法連接到 Excel。請確認 Excel 已開啟且有活動工作簿。")
        print(f"錯誤訊息: {e}")
        return

    # 連接資料庫與初始化 PiauIm
    piau_im_proc = None
    conn_15yin = None
    conn_piau_im = None
    cursor_15yin = None

    if PiauIm:
        try:
            # 1. 連接 Han_Ji_Piau_Im.db 用於 PiauIm 初始化 (轉換邏輯)
            script_dir = Path(__file__).resolve().parent
            db_piau_im_path = (script_dir / "Han_Ji_Piau_Im.db").resolve()
            if not db_piau_im_path.exists():
                print(
                    f"警告：找不到漢字標音字典檔 {db_piau_im_path}，無法進行標音轉換。"
                )
            else:
                conn_piau_im = sqlite3.connect(str(db_piau_im_path))
                # 初始化 PiauIm 並載入轉換字典，指定 "雅俗通"
                piau_im_proc = PiauIm(
                    han_ji_khoo="雅俗通", cursor=conn_piau_im.cursor()
                )

            # 2. 連接 雅俗通十五音字典.db 用於查詢原碼 (堅五曾 -> TLPA)
            db_15yin_path = (script_dir / "雅俗通十五音字典.db").resolve()
            if not db_15yin_path.exists():
                print(f"警告：找不到十五音字典檔 {db_15yin_path}，無法進行標音轉換。")
            else:
                conn_15yin = sqlite3.connect(str(db_15yin_path))
                cursor_15yin = conn_15yin.cursor()

        except Exception as e:
            print(f"資料庫初始化失敗: {e}")

    # 讀取資料
    # 從 A2 開始讀取，並嘗試讀取到 C 欄 (以防有雙排標音需求)
    try:
        last_row = sheet.range("A" + str(sheet.cells.last_cell.row)).end("up").row
        if last_row < 2:
            print("Excel 無資料 (至少需要有一列資料)。")
            return

        # 讀取 A, B, C 欄
        # A: 漢字
        # B: 標音1 (通常為上邊)
        # C: 標音2 (通常為右邊)
        # 注意：若只用 B 欄，則視設定決定放上邊 (rt) 或右邊 (rtc)
        data = sheet.range(f"A2:C{last_row}").value
    except Exception as e:
        print(f"讀取 Excel 資料失敗: {e}")
        return

    # 建構內容 HTML
    content_lines = []

    # 初始區塊
    content_lines.append('<div class="fifteen_yin">')

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
            # 單欄情況(雖然這裡是 A:C)
            row = [row]

        # 補足長度
        while len(row) < 3:
            row.append(None)

        han_ji = row[0]
        piau_im_1 = row[1]
        piau_im_2 = row[2]

        if han_ji is None:
            han_ji = ""
        if piau_im_1 is None:
            piau_im_1 = ""
        if piau_im_2 is None:
            piau_im_2 = ""

        han_ji = str(han_ji)
        piau_im_1 = str(piau_im_1)
        piau_im_2 = str(piau_im_2)

        # 換行符號處理
        if han_ji == "\\n" or han_ji == "\\r\\n" or han_ji == "\n" or han_ji == "\r\n":
            end_paragraph_if_needed()
            start_paragraph_if_needed()
            continue

        # 轉換標音
        # 根據設定 (siong_pinn_piau_im, zian_pinn_piau_im) 轉換內容
        # 預設邏輯：
        # - B欄 (piau_im_1) 為主要標音來源 (通常是十五音代碼，如 "堅五曾")
        # - 若有 C欄 (piau_im_2)，則視為第二標音來源 (若未留空)

        # 確保 system 參數在使用前已經過 strip() 清理

        # 準備上邊標音內容
        top_content = ""
        if "上" in piau_im_hong_sik:
            # 若為雙排標音且 C 欄有值，通常 B=上, C=右? 或是 B=源, 轉成上/右?
            # 依據使用者需求 (2)：依照 env 設定進行十五音轉換
            # 假設 B 欄是原始碼，轉換後填入對應位置

            # 嘗試轉換 B 欄內容
            converted_1 = convert_15yin(
                piau_im_1, siong_pinn_piau_im, piau_im_proc, cursor_15yin
            )
            top_content = converted_1

        # 準備右邊標音內容
        right_content = ""
        if "右" in piau_im_hong_sik:
            # 如果是單獨 "右邊"，B 欄是來源
            if piau_im_hong_sik == "右邊":
                right_content = convert_15yin(
                    piau_im_1, zian_pinn_piau_im, piau_im_proc, cursor_15yin
                )
            else:
                # "上及右" 或其他
                # 若 C 欄有值，優先使用 C 欄轉換
                if piau_im_2 and piau_im_2.strip():
                    right_content = convert_15yin(
                        piau_im_2, zian_pinn_piau_im, piau_im_proc, cursor_15yin
                    )
                else:
                    # 若 C 欄無值，是否重複使用 B 欄轉換？
                    # 根據慣例，若只要右邊顯示某種拼音 (例如十五音原碼)，可再次轉換 B
                    right_content = convert_15yin(
                        piau_im_1, zian_pinn_piau_im, piau_im_proc, cursor_15yin
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


if __name__ == "__main__":
    if len(sys.argv) > 1:
        output_file = sys.argv[1]
    else:
        # 預設輸出到 docs 目錄
        output_file = str(Path("docs") / "output_from_excel.html")

    export_excel_to_html(output_file)
