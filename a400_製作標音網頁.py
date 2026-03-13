"""
a400_製作標音網頁.py V0.2.2.11

修改紀錄：
v0.2.2.9 2026-2-25: 自動産生【文章標題】及【作者姓名】的 Ruby Tag。
v0.2.2.10 2026-2-27: 變更 _process_sheet() 計算 start_row 的邏輯，確保從【文章標題】及【作者姓名】之後開始讀取內容。
v0.2.2.11 2026-3-13:
  - 修復標題與作者重複出現的問題。
  - 修復檔名與 meta 標籤中出現 "None" 字串的問題。
  - 恢復完整的進度監控輸出。
"""

import os
import sys
from pathlib import Path
import xlwings as xw

from mod_excel_access import get_value_by_name
from mod_logging import (
    init_logging,
    logging_exc_error,
    logging_exception,
    logging_process_step,
    logging_warning,
)
from mod_帶調符音標 import is_han_ji, kam_si_u_tiau_hu, tng_im_piau, tng_tiau_ho
from mod_標音 import is_punctuation, split_tai_gi_im_piau
from mod_程式 import ExcelCell, Program

EXIT_CODE_SUCCESS = 0
EXIT_CODE_NO_FILE = 1
EXIT_CODE_PROCESS_FAILURE = 10

init_logging()

class CellProcessor(ExcelCell):
    def __init__(self, program: Program):
        super().__init__(program=program)
        wb = self.program.wb
        self.title = get_value_by_name(wb=wb, name="TITLE")
        self.image_url = get_value_by_name(wb=wb, name="IMAGE_URL")
        self.output_dir = "docs"
        self.total_chars_per_line = get_value_by_name(wb=wb, name="網頁每列字數")
        self.total_chars_per_line = 0 if self.total_chars_per_line is None else int(self.total_chars_per_line)
        self.han_ji_piau_im_format = get_value_by_name(wb=wb, name="網頁格式")
        self.piau_im_hong_sik = get_value_by_name(wb=wb, name="標音方式")
        self.siong_pinn_piau_im = get_value_by_name(wb=wb, name="上邊標音")
        self.zian_pinn_piau_im = get_value_by_name(wb=wb, name="右邊標音")

        self.zu_im_huat_list = {
            "SNI": ["fifteen_yin", "rt", "十五音切語"],
            "TPS": ["Piau_Im", "rt", "方音符號注音"],
            "MPS2": ["Piau_Im", "rt", "注音二式"],
            "POJ": ["pin_yin", "rt", "白話字拼音"],
            "TL": ["pin_yin", "rt", "台羅拼音"],
            "BP": ["pin_yin", "rt", "閩拼標音"],
            "TLPA_Plus": ["pin_yin", "rt", "台羅改良式"],
            "DBL": ["Siang_Pai", "rtc", "雙排注音"],
            "雅俗通": ["fifteen_yin", "rt", "十五音切語"],
            "無預設": ["Siang_Pai", "rtc", "雙排注音"],
        }

    def _build_ruby_tag(self, han_ji: str, siong_piau_im: str, zian_piau_im: str) -> str:
        if siong_piau_im != "" and zian_piau_im == "":
            return f"<ruby>\n  {han_ji}\n  <rt>{siong_piau_im}</rt>\n</ruby>\n"
        elif siong_piau_im == "" and zian_piau_im != "":
            return f"<ruby>\n  {han_ji}\n  <rtc>{zian_piau_im}</rtc>\n</ruby>\n"
        elif siong_piau_im != "" and zian_piau_im != "":
            return f"<ruby>\n  {han_ji}\n  <rt>{siong_piau_im}</rt>\n  <rtc>{zian_piau_im}</rtc>\n</ruby>\n"
        else:
            return f"<span>{han_ji}</span>\n"

    def generate_ruby_tag(self, han_ji: str, tai_gi_im_piau: str) -> tuple:
        if not tai_gi_im_piau or not str(tai_gi_im_piau).strip():
            return f"  <span>{han_ji}</span>\n", "", ""
        try:
            zu_im_list = split_tai_gi_im_piau(tai_gi_im_piau)
        except:
            return f"  <span>{han_ji}</span>\n", "", ""

        siann_bu = zu_im_list[0] if zu_im_list[0] else "ø"
        siong, zian = "", ""
        if self.han_ji_piau_im_format == "無預設":
            if "上" in str(self.piau_im_hong_sik):
                siong = self.program.piau_im.han_ji_piau_im_tng_huan(self.siong_pinn_piau_im, siann_bu, zu_im_list[1], zu_im_list[2])
            if "右" in str(self.piau_im_hong_sik):
                zian = self.program.piau_im.han_ji_piau_im_tng_huan(self.zian_pinn_piau_im, siann_bu, zu_im_list[1], zu_im_list[2])
        else:
            if self.han_ji_piau_im_format in ["POJ", "TL", "BP", "TLPA_Plus", "SNI", "雅俗通"]:
                siong = self.program.piau_im.han_ji_piau_im_tng_huan(self.siong_pinn_piau_im, siann_bu, zu_im_list[1], zu_im_list[2])
            elif self.han_ji_piau_im_format == "TPS":
                zian = self.program.piau_im.han_ji_piau_im_tng_huan(self.zian_pinn_piau_im, siann_bu, zu_im_list[1], zu_im_list[2])
            elif self.han_ji_piau_im_format == "DBL":
                siong = self.program.piau_im.han_ji_piau_im_tng_huan(self.siong_pinn_piau_im, siann_bu, zu_im_list[1], zu_im_list[2])
                zian = self.program.piau_im.han_ji_piau_im_tng_huan(self.zian_pinn_piau_im, siann_bu, zu_im_list[1], zu_im_list[2])

        ruby_tag = self._build_ruby_tag(han_ji, siong, zian)
        if tai_gi_im_piau and "<ruby" in ruby_tag:
            ruby_tag = ruby_tag.replace("<ruby", f'<ruby data-tlpa="{tai_gi_im_piau}"')
        return ruby_tag, siong, zian

    def generate_title_and_author_with_ruby(self) -> tuple:
        program = self.program
        sheet = program.wb.sheets["漢字注音"]
        start_row = program.line_start_row + program.han_ji_row_offset
        han_ji_list, tlpa_list = [], []
        row = start_row
        found_line_end = False
        while row < 1000:
            for col in range(program.start_col, program.end_col):
                h = sheet.range((row, col)).value
                t = sheet.range((row - 1, col)).value
                if h in ["φ", "\\n", "\n"]: 
                    found_line_end = True
                    break
                han_ji_list.append(str(h) if h else "")
                tlpa_list.append(str(t).strip() if t else "")
            # 重要：移到下一行標音行
            row += program.ROWS_PER_LINE
            if found_line_end: break
        
        html_segments = []
        for i, han_ji in enumerate(han_ji_list):
            if not han_ji.strip(): continue
            tag, siong, zian = self.generate_ruby_tag(han_ji, tlpa_list[i])
            html_segments.append(tag.rstrip() + "\n")
            print(f"標題處理: {han_ji} [{tlpa_list[i]}] ==》 上：{siong} / 右：{zian}")

        title_html, author_html, split_index = "", "", -1
        for i, seg in enumerate(html_segments):
            if "》" in seg: split_index = i; break
        
        if split_index != -1:
            t_parts = html_segments[:split_index+1]
            a_parts = html_segments[split_index+1:]
            title_html = "".join([p.replace("<span>", '<span class="title_mark">') if "《" in p or "》" in p else p for p in t_parts])
            author_html = "".join(a_parts)
        else:
            title_html = "".join(html_segments)

        return f"<p class='title'>{title_html}</p>", f"<p class='author'>{author_html}</p>" if author_html else "", row

    def _process_sheet(self, sheet) -> str:
        title_ruby, author_ruby, next_start_row = self.generate_title_and_author_with_ruby()
        pai_ban = self.zu_im_huat_list.get(self.han_ji_piau_im_format, ["pin_yin"])[0]
        write_buffer = f"<div class='{pai_ban}'>\n{title_ruby}\n{author_ruby}\n<p>\n"
        
        program = self.program
        char_count = 0
        end_row = program.line_end_row + program.han_ji_row_offset

        for row in range(next_start_row, end_row, program.ROWS_PER_LINE):
            try:
                sheet.range((row, program.start_col)).select()
            except:
                pass
                
            for col in range(program.start_col, program.end_col):
                val = sheet.range((row, col)).value
                addr = f"{xw.utils.col_name(col)}{row}"
                
                if val == "φ": 
                    char_count += 1
                    print(f"{char_count}. {addr} ==> 《文章終止》\n" + "="*80)
                    return write_buffer + "</p></div>"
                
                if val in ["\n", "\\n"]: 
                    char_count += 1
                    print(f"{char_count}. {addr} ==> 《換換行》\n" + "-"*80)
                    write_buffer += "</p><p>\n"
                    char_count = 0
                    break
                
                str_val = str(val).strip() if val else ""
                if is_punctuation(str_val):
                    msg = f"{str_val}【標點符號】"
                    write_buffer += f"<span>{str_val}</span>\n"
                elif not str_val:
                    msg = "【空白】"
                    write_buffer += "<span>　</span>\n"
                elif not is_han_ji(val):
                    msg = f"{str_val}【其他字元】"
                    write_buffer += f"<span>{str_val}</span>\n"
                else:
                    tlpa = sheet.range((row - 1, col)).value
                    tlpa = str(tlpa).strip() if tlpa else ""
                    if not tlpa:
                        msg = f"{str_val}【無音標】"
                        write_buffer += f"<span>{str_val}</span>\n"
                    else:
                        if kam_si_u_tiau_hu(tlpa): tlpa = tng_tiau_ho(tng_im_piau(tlpa))
                        tag, siong, zian = self.generate_ruby_tag(str_val, tlpa)
                        write_buffer += tag
                        msg = f"{str_val} [{tlpa}] ==》 上：{siong} / 右：{zian}"
                
                char_count += 1
                print(f"{char_count}. {addr} ==> {msg}")
                if self.total_chars_per_line and char_count >= self.total_chars_per_line:
                    write_buffer += "<br/>\n"; char_count = 0; print("《人工斷行》")
        return write_buffer + "</p></div>"

def _create_html_file(program, output_path, head_extra="", title="", content=""):
    web_page_stem = Path(output_path).stem
    img_url = str(program.image_url or "").strip()
    if img_url == "None" or not img_url:
        img_url = "king_tian.png" # 預設圖片
    full_img = img_url if img_url.startswith("http") else f"https://alanjui.github.io/Piau-Im/assets/images/{img_url}"
    
    template = f"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{title}</title>
    <meta content='https://alanjui.github.io/Piau-Im/{web_page_stem}.html' property='og:url' />
    {head_extra}
    <link rel="stylesheet" href="./assets/styles/styles.css">
    <script type="text/javascript" src="./assets/javascripts/phonetic_switcher.js"></script>
</head>
<body>
    <main class="page">
        <article class="article_content">
            <div style='text-align: center'><img src='{full_img}' width='800' /></div>
            {content}
        </article>
    </main>
    <a href="index.html" class="floating-home-btn">🏠</a>
</body>
</html>"""
    with open(output_path, "w", encoding="utf-8") as f: f.write(template)

def process(wb, args) -> int:
    logging_process_step("<=========== 作業開始！==========>")
    try:
        program = Program(wb=wb, args=args, hanji_piau_im_sheet_name="漢字注音")
        xls_cell = CellProcessor(program=program)
        sheet = wb.sheets["漢字注音"]
        sheet.activate()
        print("開始製作【漢字注音】網頁！")
        html_content = xls_cell._process_sheet(sheet)

        # 生成輸出檔案名稱 (處理 None)
        piau_im_huat = program.piau_im_huat
        ue_im_lui_piat = program.ue_im_lui_piat
        han_ji_piau_im_format = program.han_ji_piau_im_format
        siong = str(program.siong_pinn_piau_im) if program.siong_pinn_piau_im and str(program.siong_pinn_piau_im) != "None" else ""
        zian = str(program.zian_pinn_piau_im) if program.zian_pinn_piau_im and str(program.zian_pinn_piau_im) != "None" else ""

        if han_ji_piau_im_format == "無預設":
            im_piau = piau_im_huat
        else:
            if siong and zian: im_piau = f"{siong}＋{zian}"
            elif siong: im_piau = siong
            elif zian: im_piau = zian
            else: im_piau = piau_im_huat

        output_file = f"《{program.title}》【{ue_im_lui_piat}】{im_piau}.html"
        output_path = os.path.join("docs", output_file)
        os.makedirs("docs", exist_ok=True)
        
        # 處理 Meta 標籤中的 None
        meta_keys = ["TITLE", "IMAGE_URL", "網頁格式", "上邊標音", "右邊標音"]
        meta_list = []
        for k in meta_keys:
            v = get_value_by_name(wb, k)
            if v is None or str(v) == "None": v = ""
            meta_list.append(f'<meta name="{k}" content="{v}" />')
        head_extra = "\n    ".join(meta_list)
        
        _create_html_file(program, output_path, head_extra, program.title, html_content)
        program.save_workbook_as_new_file(wb=wb)
        logging_process_step("<=========== 作業結束！==========>")
        return EXIT_CODE_SUCCESS
    except Exception as e:
        logging_exc_error("處理作業異常", e)
        return EXIT_CODE_PROCESS_FAILURE

def main(args):
    try:
        wb = xw.apps.active.books.active
        return process(wb, args)
    except Exception as e:
        logging_exception("無法找到作用中的 Excel 工作簿！", e)
        return EXIT_CODE_NO_FILE

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--test", action="store_true")
    parser.add_argument("--new", action="store_true")
    args = parser.parse_args()
    sys.exit(main(args))
