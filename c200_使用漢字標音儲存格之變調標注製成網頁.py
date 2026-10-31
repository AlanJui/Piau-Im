"""
c200_使用漢字標音儲存格之變調標注製成網頁.py v0.1.0

依 a400 的網頁骨架產出 HTML，標音來源改為：

1. 先執行 c100，將【台語音標】轉成【本調／變調】並寫入【漢字標音】儲存格。
2. 再執行本程式：
   - 【漢字標音】儲存格內容直接寫入 <ruby> 的 <rt>（不再依【上邊標音】轉換）。
   - 由【台語音標】轉成標準【方音符號】，寫入 <rtc>。
   - 標點、空白、非漢字以 <span> 輸出，不進入 <rt>。

<ruby> 仍附 data-tlpa、data-pian-tiau，供網頁標音切換器使用。
頁面 meta【使用變調標音格式】不寫成「是」，避免切換器在載入時覆寫已寫好的 <rt>。
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
)
from mod_帶調符音標 import is_han_ji, kam_si_u_tiau_hu, tng_im_piau, tng_tiau_ho
from mod_標音 import is_punctuation, split_tai_gi_im_piau
from mod_程式 import ExcelCell, Program

EXIT_CODE_SUCCESS = 0
EXIT_CODE_NO_FILE = 1
EXIT_CODE_PROCESS_FAILURE = 10

# 變調標示底色（淺灰 #E8E8E8），與 c100 相同
BIAN_TIAU_FILL_COLOR = (232, 232, 232)

# 雙排：上端為漢字標音（<rt>），右端為方音符號（<rtc>）
PAI_BAN_CLASS = "Siang_Pai"
HONG_IM_HU_HO = "方音符號"

init_logging()


def _clean_named_value(value) -> str:
    """命名範圍空值或字串 None 視為空白。"""
    text = str(value or "").strip()
    return "" if text == "None" else text


class CellProcessor(ExcelCell):
    def __init__(self, program: Program):
        super().__init__(program=program)
        wb = self.program.wb
        self.title = get_value_by_name(wb=wb, name="TITLE")
        self.image_url = get_value_by_name(wb=wb, name="IMAGE_URL")
        self.output_dir = "docs"
        self.total_chars_per_line = get_value_by_name(wb=wb, name="網頁每列字數")
        self.total_chars_per_line = (
            0 if self.total_chars_per_line is None else int(self.total_chars_per_line)
        )

    def _build_ruby_tag(
        self, han_ji: str, siong_piau_im: str, zian_piau_im: str
    ) -> str:
        tab = "\t"
        left_margin = f"{tab * 5}"
        if siong_piau_im != "" and zian_piau_im == "":
            return f"{left_margin}<ruby>{han_ji}<rt>{siong_piau_im}</rt></ruby>\n"
        if siong_piau_im == "" and zian_piau_im != "":
            return f"{left_margin}<ruby>{han_ji}<rtc>{zian_piau_im}</rtc></ruby>\n"
        if siong_piau_im != "" and zian_piau_im != "":
            return (
                f"{left_margin}<ruby>{han_ji}"
                f"<rt>{siong_piau_im}</rt><rtc>{zian_piau_im}</rtc></ruby>\n"
            )
        return f"{left_margin}<span>{han_ji}</span>\n"

    @staticmethod
    def _is_pian_tiau_cell(cell) -> bool:
        """漢字儲存格是否塗成變調標示淺灰底色。"""
        return cell.color == BIAN_TIAU_FILL_COLOR

    @staticmethod
    def _attach_ruby_data_attrs(
        ruby_tag: str, tai_gi_im_piau: str = "", pian_tiau: bool = False
    ) -> str:
        """為 <ruby> 附加 data-tlpa、data-pian-tiau 屬性。"""
        if "<ruby" not in ruby_tag:
            return ruby_tag
        attrs = []
        if tai_gi_im_piau and str(tai_gi_im_piau).strip():
            attrs.append(f'data-tlpa="{str(tai_gi_im_piau).strip()}"')
        attrs.append(f'data-pian-tiau="{"true" if pian_tiau else "false"}"')
        return ruby_tag.replace("<ruby", "<ruby " + " ".join(attrs), 1)

    def _normalize_tlpa(self, tai_gi_im_piau: str) -> str:
        """帶調符音標轉成帶調號的台語音標，供 data-tlpa 與方音符號轉換。"""
        im_piau = str(tai_gi_im_piau or "").strip()
        if not im_piau:
            return ""
        if kam_si_u_tiau_hu(im_piau):
            im_piau = tng_tiau_ho(tng_im_piau(im_piau))
        return im_piau

    def _convert_hong_im_hu_ho(self, tai_gi_im_piau: str) -> str:
        """將台語音標轉成標準方音符號（調符格式），寫入 <rtc>。"""
        im_piau = self._normalize_tlpa(tai_gi_im_piau)
        if not im_piau:
            return ""
        try:
            siann_bu, un_bu, tiau_ho = split_tai_gi_im_piau(im_piau)
        except Exception as error:
            logging_exception(
                msg=f"拆解台語音標失敗，無法產生方音符號：{im_piau}",
                error=error,
            )
            return ""
        if not un_bu:
            return ""
        siann_bu = siann_bu if siann_bu else "ø"
        try:
            hong_im = self.program.piau_im.han_ji_piau_im_tng_huan(
                HONG_IM_HU_HO, siann_bu, un_bu, str(tiau_ho)
            )
        except Exception as error:
            logging_exception(
                msg=f"台語音標轉方音符號失敗：{im_piau}",
                error=error,
            )
            return ""
        return str(hong_im or "").strip()

    def _read_han_ji_piau_im(self, han_ji_piau_im: str) -> str:
        """取【漢字標音】儲存格文字。標點不寫入 <rt>。"""
        text = _clean_named_value(han_ji_piau_im)
        if not text or is_punctuation(text):
            return ""
        return text

    def generate_ruby_tag(
        self,
        han_ji: str,
        tai_gi_im_piau: str,
        han_ji_piau_im: str = "",
        pian_tiau: bool = False,
    ) -> tuple:
        """
        產製 <ruby>。

        <rt> 直接採用【漢字標音】儲存格（c100 已寫入的本調／變調）。
        <rtc> 為由台語音標轉成的方音符號。
        """
        tab = "\t"
        left_margin = f"{tab * 5}"
        han_ji_text = "" if han_ji is None else str(han_ji)

        if is_punctuation(han_ji_text) or not is_han_ji(han_ji_text):
            return f"{left_margin}<span>{han_ji_text}</span>\n", "", ""

        tlpa = self._normalize_tlpa(tai_gi_im_piau)
        siong = self._read_han_ji_piau_im(han_ji_piau_im)
        zian = self._convert_hong_im_hu_ho(tlpa)

        ruby_tag = self._attach_ruby_data_attrs(
            self._build_ruby_tag(han_ji_text, siong, zian),
            tai_gi_im_piau=tlpa,
            pian_tiau=pian_tiau,
        )
        return ruby_tag, siong, zian

    def generate_title_and_author_with_ruby(self) -> tuple:
        """自文首《》拆出標題與作者，標音同樣取自漢字標音儲存格。"""
        program = self.program
        sheet = program.wb.sheets["漢字注音"]
        start_row = program.line_start_row + program.han_ji_row_offset

        first_han_ji = sheet.range((start_row, program.start_col)).value
        if first_han_ji != "《":
            return "", "", start_row

        han_ji_list = []
        tlpa_list = []
        han_ji_piau_im_list = []
        pian_tiau_list = []
        row = start_row
        found_line_end = False
        while row < 1000:
            for col in range(program.start_col, program.end_col + 1):
                han_ji_cell = sheet.range((row, col))
                han_ji = han_ji_cell.value
                tai_gi = sheet.range((row - 1, col)).value
                piau_im = sheet.range((row + 1, col)).value
                if han_ji in ["φ", "\\n", "\n"]:
                    found_line_end = True
                    break
                han_ji_list.append(str(han_ji) if han_ji else "")
                tlpa_list.append(str(tai_gi).strip() if tai_gi else "")
                han_ji_piau_im_list.append(str(piau_im).strip() if piau_im else "")
                pian_tiau_list.append(self._is_pian_tiau_cell(han_ji_cell))
            row += program.ROWS_PER_LINE
            if found_line_end:
                break

        html_segments = []
        tab = "\t"
        left_margin = f"{tab * 5}"
        for index, han_ji in enumerate(han_ji_list):
            if not han_ji.strip():
                continue
            if is_punctuation(han_ji) or not is_han_ji(han_ji):
                tag = f"{left_margin}<span>{han_ji}</span>\n"
                siong, zian = "", ""
            else:
                tag, siong, zian = self.generate_ruby_tag(
                    han_ji,
                    tlpa_list[index],
                    han_ji_piau_im=han_ji_piau_im_list[index],
                    pian_tiau=pian_tiau_list[index],
                )
            html_segments.append(tag.rstrip() + "\n")
            print(f"標題處理: {han_ji} [{tlpa_list[index]}] ==》 上：{siong} / 右：{zian}")

        title_html, author_html = "", ""
        split_index = -1
        for index, segment in enumerate(html_segments):
            if "》" in segment:
                split_index = index
                break

        if split_index != -1:
            title_parts = html_segments[: split_index + 1]
            author_parts = html_segments[split_index + 1 :]
            title_html = "".join(
                [
                    (
                        part.replace("<span>", '<span class="title_mark">')
                        if "《" in part or "》" in part
                        else part
                    )
                    for part in title_parts
                ]
            )
            combined_after = "".join(han_ji_list[split_index + 1 :])
            if "：" in combined_after or ":" in combined_after:
                author_html = "".join(author_parts)
            else:
                author_html = ""
        else:
            title_html = "".join(html_segments)

        return title_html, author_html if author_html else "", row

    def _process_sheet(self, sheet) -> str:
        tab = "\t"
        left_margin = f"{tab * 3}"
        title_ruby, author_ruby, next_start_row = (
            self.generate_title_and_author_with_ruby()
        )
        write_buffer = f"""
{left_margin}<div class='{PAI_BAN_CLASS}'>
{left_margin}\t<p class='title'>
{title_ruby}{left_margin}\t</p>
{left_margin}\t<p class='author'>
{author_ruby}{left_margin}\t</p>
{left_margin}\t<p>
"""
        program = self.program
        char_count = 0
        end_row = program.line_end_row + program.han_ji_row_offset
        closing = f"{tab * 4}</p>\n{tab * 3}</div>"

        for row in range(next_start_row, end_row, program.ROWS_PER_LINE):
            try:
                sheet.range((row, program.start_col)).select()
            except Exception:
                pass

            for col in range(program.start_col, program.end_col + 1):
                han_ji_cell = sheet.range((row, col))
                val = han_ji_cell.value
                addr = f"{xw.utils.col_name(col)}{row}"

                if val == "φ":
                    char_count += 1
                    print(f"{char_count}. {addr} ==> 《文章終止》\n" + "=" * 80)
                    return write_buffer + closing

                if val in ["\n", "\\n"]:
                    char_count += 1
                    print(f"{char_count}. {addr} ==> 《換行》\n" + "-" * 80)
                    write_buffer += f"{tab * 4}</p><p>\n"
                    char_count = 0
                    break

                str_val = str(val).strip() if val else ""
                if is_punctuation(str_val):
                    msg = f"{str_val}【標點符號】"
                    write_buffer += f"{tab * 5}<span>{str_val}</span>\n"
                elif not str_val:
                    msg = "【空白】"
                    write_buffer += f"{tab * 5}<span>　</span>\n"
                elif not is_han_ji(val):
                    msg = f"{str_val}【其他字元】"
                    write_buffer += f"{tab * 5}<span>{str_val}</span>\n"
                else:
                    tlpa = sheet.range((row - 1, col)).value
                    tlpa = str(tlpa).strip() if tlpa else ""
                    han_ji_piau_im = sheet.range((row + 1, col)).value
                    han_ji_piau_im = str(han_ji_piau_im).strip() if han_ji_piau_im else ""
                    pian_tiau = self._is_pian_tiau_cell(han_ji_cell)
                    tag, siong, zian = self.generate_ruby_tag(
                        str_val,
                        tlpa,
                        han_ji_piau_im=han_ji_piau_im,
                        pian_tiau=pian_tiau,
                    )
                    write_buffer += tag
                    pian_flag = "變調" if pian_tiau else "本調"
                    msg = (
                        f"{str_val} [{tlpa}] ==》 上：{siong} / 右：{zian}"
                        f"（{pian_flag}）"
                    )

                char_count += 1
                print(f"{char_count}. {addr} ==> {msg}")
                if (
                    self.total_chars_per_line
                    and char_count >= self.total_chars_per_line
                ):
                    write_buffer += "</p><p>\n"
                    char_count = 0
                    print("《人工斷行》")

        return write_buffer + closing


def _create_html_file(program, output_path, head_extra="", title="", content=""):
    web_page_stem = Path(output_path).stem
    img_url = _clean_named_value(getattr(program, "image_url", ""))
    if not img_url:
        img_url = "king_tian.png"

    full_img = img_url if img_url.startswith("http") else f"./assets/images/{img_url}"
    page_title = _clean_named_value(title)

    template = f"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{page_title}</title>
    <meta content='https://alanjui.github.io/Piau-Im/{web_page_stem}.html' property='og:url' />
    {head_extra}
    <link rel="stylesheet" href="./assets/styles/styles.css">
    <script type="text/javascript" src="./assets/javascripts/phonetic_switcher.js"></script>
</head>
<body>
    <main class="page">
        <article class="article_content">
            <div style='text-align: center'>
                <img src='{full_img}' width='800' />
            </div>
            {content}
        </article>
    </main>
    <a href="index.html" class="floating-home-btn">🏠</a>
</body>
</html>"""
    with open(output_path, "w", encoding="utf-8") as html_file:
        html_file.write(template)


def _build_head_extra(wb, piau_im_huat: str) -> str:
    """
    工具列初始值對齊本頁：上端為 c100 所用標音方法，右端為方音符號。

    【使用變調標音格式】留空，載入時不覆寫 <rt>。
    """
    meta_values = {
        "TITLE": get_value_by_name(wb, "TITLE"),
        "IMAGE_URL": get_value_by_name(wb, "IMAGE_URL"),
        "網頁格式": PAI_BAN_CLASS,
        "標音方法": piau_im_huat,
        "上邊標音": piau_im_huat,
        "右邊標音": HONG_IM_HU_HO,
        "使用變調標音格式": "",
    }
    meta_list = []
    for key, value in meta_values.items():
        text = _clean_named_value(value)
        meta_list.append(f'<meta name="{key}" content="{text}" />')
    return "\n    ".join(meta_list)


def process(wb, args) -> int:
    logging_process_step("<=========== 作業開始！==========>")
    try:
        program = Program(wb=wb, args=args, hanji_piau_im_sheet_name="漢字注音")
        xls_cell = CellProcessor(program=program)
        sheet = wb.sheets["漢字注音"]
        sheet.activate()
        print("開始以【漢字標音儲存格】與【方音符號】製作網頁！")
        logging_process_step(
            "【漢字標音】儲存格 → <rt>；台語音標轉【方音符號】→ <rtc>。"
        )
        html_content = xls_cell._process_sheet(sheet)

        title = _clean_named_value(program.title)
        ue_im_lui_piat = _clean_named_value(program.ue_im_lui_piat)
        piau_im_huat = _clean_named_value(program.piau_im_huat) or "變調標注"
        output_file = f"《{title}》【{ue_im_lui_piat}】{piau_im_huat}變調標注＋方音符號.html"
        output_path = os.path.join("docs", output_file)
        os.makedirs("docs", exist_ok=True)

        head_extra = _build_head_extra(wb, piau_im_huat)
        _create_html_file(program, output_path, head_extra, title, html_content)
        logging_process_step(f"已寫入網頁：{output_path}")

        program.save_workbook_as_new_file(wb=wb)
        logging_process_step("<=========== 作業結束！==========>")
        return EXIT_CODE_SUCCESS
    except Exception as error:
        logging_exc_error("處理作業異常", error)
        return EXIT_CODE_PROCESS_FAILURE


def main(args):
    try:
        wb = xw.apps.active.books.active
        return process(wb, args)
    except Exception as error:
        logging_exception("無法找到作用中的 Excel 工作簿！", error)
        return EXIT_CODE_NO_FILE


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description="以【漢字標音】儲存格的變調標注，與方音符號，製成標音網頁",
    )
    parser.add_argument("--new", action="store_true")
    args = parser.parse_args()
    sys.exit(main(args))
