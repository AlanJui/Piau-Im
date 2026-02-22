"""
a940_自Excel轉製html檔.py

功能：
    讀取 Excel 工作表中的資料，將其轉換為帶有 <ruby> 標籤的 HTML 格式。
    預期 Excel 格式：
    Column A: 漢字 / 標點符號 / 換行符號(\n)
    Column B: 漢字標音 (若為空，則視為標點符號，不加 ruby)

    當 Column A 為換行符號 (\n) 時，表示段落結束，將輸出 </p><p>。

使用方式：
    python a940_自Excel轉製html檔.py [output_html_file_path]
"""

import sys
from pathlib import Path

import xlwings as xw


def export_excel_to_html(output_path):
    # 連接 Excel
    try:
        wb = xw.books.active
        sheet = wb.sheets.active  # 假設使用者停留在要輸出的工作表
    except Exception as e:
        print("無法連接到 Excel。請確認 Excel 已開啟且有活動工作簿。")
        print(f"錯誤訊息: {e}")
        return

    # 讀取資料
    # 假設這是有標頭的，從 A2 開始讀取
    # 找到最後一行
    last_row = sheet.range("A" + str(sheet.cells.last_cell.row)).end("up").row
    if last_row < 2:
        print("Excel 無資料 (至少需要有一列資料)。")
        return

    # 讀取 A 和 B 欄
    data = sheet.range(f"A2:B{last_row}").value

    # 建構 HTML 內容
    html_lines = []
    html_lines.append("<!DOCTYPE html>")
    html_lines.append("<html>")
    html_lines.append("<head>")
    html_lines.append('<meta charset="utf-8">')
    html_lines.append("<style>")
    html_lines.append(
        '  body { font-family: "Microsoft JhengHei", serif; font-size: 1.2em; line-height: 2.0; }'
    )
    html_lines.append("  ruby { font-size: 1.5em; margin-right: 0.2em; }")
    html_lines.append("  rt { font-size: 0.5em; }")
    html_lines.append("</style>")
    html_lines.append("</head>")
    html_lines.append("<body>")
    html_lines.append('<div class="content-box">')

    # 初始段落
    current_line = []
    in_paragraph = False

    # 確保輸出目錄存在
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)

    def start_paragraph_if_needed():
        nonlocal in_paragraph
        if not in_paragraph:
            html_lines.append("<p>")
            in_paragraph = True

    def end_paragraph_if_needed():
        nonlocal in_paragraph
        if in_paragraph:
            html_lines.append("</p>")
            in_paragraph = False

    # 強制開始第一個段落
    start_paragraph_if_needed()

    for row in data:
        # row 可能為 None (某些情況下) 或是 list
        if row is None:
            continue

        han_ji = row[0]
        piau_im = row[1]

        # 處理 None
        if han_ji is None:
            han_ji = ""
        if piau_im is None:
            piau_im = ""

        # 轉字串
        han_ji = str(han_ji)
        piau_im = str(piau_im)

        # 檢查是否為換行符號
        # Excel 的 Alt+Enter 通常是 \n，但在 xlwings 讀出來可能是 \n
        # 題目說: 當 【A欄】儲存格為 =CHAR(10) '\n'
        if han_ji == "\n" or han_ji == "\r\n":
            end_paragraph_if_needed()
            start_paragraph_if_needed()
            continue

        if piau_im.strip():
            # 有標音 -> <ruby>
            html_chunk = f"<ruby>{han_ji}<rt>{piau_im}</rt></ruby>"
            start_paragraph_if_needed()  # 確保在段落內
            html_lines.append("  " + html_chunk)
        else:
            # 無標音 -> 純文字
            # 可能是標點符號或純漢字
            start_paragraph_if_needed()  # 確保在段落內
            html_lines.append("  " + han_ji)

    # 結束最後一個段落
    end_paragraph_if_needed()

    html_lines.append("</div>")
    html_lines.append("</body>")
    html_lines.append("</html>")

    # 寫入檔案
    try:
        with open(output_path, "w", encoding="utf-8") as f:
            f.write("\n".join(html_lines))
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
