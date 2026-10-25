"""
a400_製作標音網頁.py V0.2.2.17

修改紀錄：
v0.2.2.9 2026-2-25: 自動産生【文章標題】及【作者姓名】的 Ruby Tag。
v0.2.2.10 2026-2-27: 變更 _process_sheet() 計算 start_row 的邏輯，確保從【文章標題】及【作者姓名】之後開始讀取內容。
v0.2.2.11 2026-3-13:
  - 修復標題與作者重複出現的問題。
  - 修復檔名與 meta 標籤中出現 "None" 字串的問題。
  - 恢復完整的進度監控輸出。
v0.2.2.12 2026-3-20: 改善 html 輸出之格式，確保【內縮】與【換行】的正確顯示。
v0.2.2.13 2026-10-25:
  - 若 env【使用變調標音格式】=「是」，略過【上邊標音】轉換，
    直接將【漢字標音】儲存格內容置入 <rt>。
v0.2.2.14 2026-10-25:
  - 標點符號（如《》、。、)）不得置入 <rt>，一律以 <span> 輸出。
v0.2.2.15 2026-10-25:
  - 使用變調標音格式時：略過【上邊標音】，但【右邊標音】（如方音符號）
    只要有設定值，仍一律轉換為 <rtc>。
v0.2.2.16 2026-10-25:
  - <ruby> 新增 data-pian-tiau="true|false"：依漢字儲存格淺灰底色標示是否變調。
v0.2.2.17 2026-10-25:
  - 改依「使用變調標音格式」流程圖：標音方式含「上」時，
    依【上邊標音】將台語音標轉成本調／變調格式作 <rt>；
    含「右」時仍依【右邊標音】作 <rtc>。
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

init_logging()


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
        self.han_ji_piau_im_format = get_value_by_name(wb=wb, name="網頁格式")
        self.piau_im_hong_sik = get_value_by_name(wb=wb, name="標音方式")
        self.siong_pinn_piau_im = get_value_by_name(wb=wb, name="上邊標音")
        self.zian_pinn_piau_im = get_value_by_name(wb=wb, name="右邊標音")
        # 【使用變調標音格式】=「是」：依【上邊標音】將台語音標轉成本調／變調格式作 <rt>
        iong_bian_tiau = get_value_by_name(wb=wb, name="使用變調標音格式")
        self.iong_bian_tiau_piau_im_format = str(iong_bian_tiau or "").strip() == "是"

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

    def _build_ruby_tag(
        self, han_ji: str, siong_piau_im: str, zian_piau_im: str
    ) -> str:
        TAB = "\t"  # t = char(9) 取名 tab
        LEFT_MARGIN = f"{TAB*5}"  # 左邊距，根據實際需要調整
        if siong_piau_im != "" and zian_piau_im == "":
            return f"{LEFT_MARGIN}<ruby>{han_ji}<rt>{siong_piau_im}</rt></ruby>\n"
        elif siong_piau_im == "" and zian_piau_im != "":
            return f"{LEFT_MARGIN}<ruby>{han_ji}<rtc>{zian_piau_im}</rtc></ruby>\n"
        elif siong_piau_im != "" and zian_piau_im != "":
            return f"{LEFT_MARGIN}<ruby>{han_ji}<rt>{siong_piau_im}</rt><rtc>{zian_piau_im}</rtc></ruby>\n"
        else:
            return f"{LEFT_MARGIN}<span>{han_ji}</span>\n"

    def _has_zian_pinn_piau_im(self) -> bool:
        """【右邊標音】是否有有效設定值。"""
        zian_huat = str(self.zian_pinn_piau_im or "").strip()
        return bool(zian_huat) and zian_huat != "None"

    def _hong_sik_has_siong(self) -> bool:
        """標音方式是否含「上」（上／上及右）。"""
        return "上" in str(self.piau_im_hong_sik or "")

    def _hong_sik_has_zian(self) -> bool:
        """標音方式是否含「右」（右／上及右）。"""
        return "右" in str(self.piau_im_hong_sik or "")

    def _convert_zian_piau_im(self, siann_bu: str, un_bu: str, tiau_ho: str) -> str:
        """依【右邊標音】設定，將台語音標聲韻調轉成右邊標音文字。"""
        if not self._has_zian_pinn_piau_im():
            return ""
        return self.program.piau_im.han_ji_piau_im_tng_huan(
            self.zian_pinn_piau_im, siann_bu, un_bu, tiau_ho
        )

    def _convert_siong_bun_or_bian_tiau(
        self, tai_gi_im_piau: str, use_bian_tiau: bool
    ) -> str:
        """
        依【上邊標音】將台語音標轉成【本調】或【變調】標音格式（供 <rt>）。
        """
        siong_huat = str(self.siong_pinn_piau_im or "").strip()
        if not siong_huat or siong_huat == "None":
            return ""
        # 延遲匯入，避免單元測試假模組與 c100 互相牽連
        from c100_漢字標音儲存格改用變調標示格式 import (
            convert_to_bun_or_bian_tiau_format,
        )

        return convert_to_bun_or_bian_tiau_format(
            piau_im=self.program.piau_im,
            piau_im_huat=siong_huat,
            tai_gi_im_piau=tai_gi_im_piau,
            use_bian_tiau=use_bian_tiau,
        )

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
        # HTML data 屬性可用字串 "true"/"false"
        attrs.append(f'data-pian-tiau="{"true" if pian_tiau else "false"}"')
        return ruby_tag.replace("<ruby", "<ruby " + " ".join(attrs), 1)

    def generate_ruby_tag(
        self,
        han_ji: str,
        tai_gi_im_piau: str,
        han_ji_piau_im: str = "",  # 保留相容；變調格式模式改由台語音標＋上邊標音轉換
        pian_tiau: bool = False,
    ) -> tuple:
        """
        產製 <ruby>。

        使用變調標音格式＝是（且標音方式含「上」）時：
        - 淺灰底 → data-pian-tiau=true，<rt>＝依【上邊標音】轉成變調格式
        - 非淺灰 → data-pian-tiau=false，<rt>＝依【上邊標音】轉成本調格式
        標音方式含「右」時：<rtc> 仍依【右邊標音】標準轉換。
        使用變調標音格式＝否：延用原處理邏輯。
        """
        TAB = "\t"  # t = char(9) 取名 tab
        LEFT_MARGIN = f"{TAB*5}"  # 左邊距，根據實際需要調整
        _ = han_ji_piau_im  # 變調格式模式不再直接取用【漢字標音】儲存格

        # 標點／非漢字：不得置入 <rt>，一律以 <span> 輸出
        if is_punctuation(han_ji) or not is_han_ji(han_ji):
            return f"{LEFT_MARGIN}<span>{han_ji}</span>\n", "", ""

        if not tai_gi_im_piau or not str(tai_gi_im_piau).strip():
            return f"{LEFT_MARGIN}<span>{han_ji}</span>\n", "", ""

        try:
            zu_im_list = split_tai_gi_im_piau(tai_gi_im_piau)
        except Exception:
            return f"{LEFT_MARGIN}<span>{han_ji}</span>\n", "", ""

        siann_bu = zu_im_list[0] if zu_im_list[0] else "ø"
        un_bu = zu_im_list[1]
        tiau_ho = zu_im_list[2]
        siong, zian = "", ""

        if self.iong_bian_tiau_piau_im_format:
            # ---------- 使用變調標音格式之處理邏輯 ----------
            # <rt>：標音方式＝上 或 上及右
            if self._hong_sik_has_siong():
                siong = self._convert_siong_bun_or_bian_tiau(
                    tai_gi_im_piau=tai_gi_im_piau,
                    use_bian_tiau=pian_tiau,
                )
            # <rtc>：標音方式＝右 或 上及右（流程圖中「製作 <rt>」為筆誤，實為 <rtc>）
            if self._hong_sik_has_zian():
                zian = self._convert_zian_piau_im(siann_bu, un_bu, tiau_ho)
        elif self.han_ji_piau_im_format == "無預設":
            # ---------- 原處理邏輯 ----------
            if self._hong_sik_has_siong():
                siong = self.program.piau_im.han_ji_piau_im_tng_huan(
                    self.siong_pinn_piau_im, siann_bu, un_bu, tiau_ho
                )
            if self._hong_sik_has_zian():
                zian = self._convert_zian_piau_im(siann_bu, un_bu, tiau_ho)
        else:
            if self.han_ji_piau_im_format in [
                "POJ",
                "TL",
                "BP",
                "TLPA_Plus",
                "SNI",
                "雅俗通",
            ]:
                siong = self.program.piau_im.han_ji_piau_im_tng_huan(
                    self.siong_pinn_piau_im, siann_bu, un_bu, tiau_ho
                )
            elif self.han_ji_piau_im_format == "TPS":
                zian = self._convert_zian_piau_im(siann_bu, un_bu, tiau_ho)
            elif self.han_ji_piau_im_format == "DBL":
                siong = self.program.piau_im.han_ji_piau_im_tng_huan(
                    self.siong_pinn_piau_im, siann_bu, un_bu, tiau_ho
                )
                zian = self._convert_zian_piau_im(siann_bu, un_bu, tiau_ho)

        ruby_tag = self._attach_ruby_data_attrs(
            self._build_ruby_tag(han_ji, siong, zian),
            tai_gi_im_piau=tai_gi_im_piau,
            pian_tiau=pian_tiau,
        )
        return ruby_tag, siong, zian

    def generate_title_and_author_with_ruby(self) -> tuple:
        """智慧判定標題與作者"""
        program = self.program
        sheet = program.wb.sheets["漢字注音"]
        start_row = program.line_start_row + program.han_ji_row_offset

        # 1. 讀取第一行第一個漢字
        first_han_ji = sheet.range((start_row, program.start_col)).value
        if first_han_ji != "《":
            # 如果開頭不是《，判定為無標題文章，直接回傳空，並讓指標維持在起始行
            return "", "", start_row

        # 2. 既然有《，開始提取
        han_ji_list, tlpa_list, pian_tiau_list = [], [], []
        row = start_row
        found_line_end = False
        while row < 1000:
            for col in range(program.start_col, program.end_col + 1):
                han_ji_cell = sheet.range((row, col))
                h = han_ji_cell.value
                t = sheet.range((row - 1, col)).value
                if h in ["φ", "\\n", "\n"]:
                    found_line_end = True
                    break
                han_ji_list.append(str(h) if h else "")
                tlpa_list.append(str(t).strip() if t else "")
                pian_tiau_list.append(self._is_pian_tiau_cell(han_ji_cell))
            row += program.ROWS_PER_LINE
            if found_line_end:
                break

        html_segments = []
        TAB = "\t"
        LEFT_MARGIN = f"{TAB*5}"
        for i, han_ji in enumerate(han_ji_list):
            if not han_ji.strip():
                continue
            # 標題／作者列的標點（《》、。、)、：等）只用 <span>，不產 <rt>
            if is_punctuation(han_ji) or not is_han_ji(han_ji):
                tag = f"{LEFT_MARGIN}<span>{han_ji}</span>\n"
                siong, zian = "", ""
            else:
                tag, siong, zian = self.generate_ruby_tag(
                    han_ji,
                    tlpa_list[i],
                    pian_tiau=pian_tiau_list[i],
                )
            html_segments.append(tag.rstrip() + "\n")
            print(f"標題處理: {han_ji} [{tlpa_list[i]}] ==》 上：{siong} / 右：{zian}")

        # 3. 根據標點符號拆分
        title_html, author_html, split_index = "", "", -1
        for i, seg in enumerate(html_segments):
            if "》" in seg:
                split_index = i
                break

        if split_index != -1:
            t_parts = html_segments[: split_index + 1]
            a_parts = html_segments[split_index + 1 :]
            title_html = "".join(
                [
                    (
                        p.replace("<span>", f'<span class="title_mark">')
                        if "《" in p or "》" in p
                        else p
                    )
                    for p in t_parts
                ]
            )
            # 檢查後續是否包含作者標記
            combined_after = "".join(han_ji_list[split_index + 1 :])
            if "：" in combined_after or ":" in combined_after:
                author_html = "".join(a_parts)
            else:
                # 若無作者標記，則作者留空，且將指標重設回標題後的一行
                author_html = ""
                # 注意：這裡 row 已經在提取完畢後指向正確位置
        else:
            title_html = "".join(html_segments)

        return (
            f"{title_html}",
            f"{author_html}" if author_html else "",
            row,
        )

    def _process_sheet(self, sheet) -> str:
        TAB = "\t"  # t = char(9) 取名 tab
        LEFT_MARGIN = f"{TAB*3}"  # 左邊距，根據實際需要調整
        title_ruby, author_ruby, next_start_row = (
            self.generate_title_and_author_with_ruby()
        )
        pai_ban = self.zu_im_huat_list.get(self.han_ji_piau_im_format, ["pin_yin"])[0]
        write_buffer = f"""
{LEFT_MARGIN}<div class='{pai_ban}'>
{LEFT_MARGIN}\t<p class='title'>
{title_ruby}{LEFT_MARGIN}\t</p>
{LEFT_MARGIN}\t<p class='author'>
{author_ruby}{LEFT_MARGIN}\t</p>
{LEFT_MARGIN}\t<p>
"""
        program = self.program
        char_count = 0
        end_row = program.line_end_row + program.han_ji_row_offset

        for row in range(next_start_row, end_row, program.ROWS_PER_LINE):
            try:
                sheet.range((row, program.start_col)).select()
            except Exception:
                pass

            for col in range(program.start_col, program.end_col+1):
                val = sheet.range((row, col)).value
                addr = f"{xw.utils.col_name(col)}{row}"

                if val == "φ":
                    char_count += 1
                    print(f"{char_count}. {addr} ==> 《文章終止》\n" + "=" * 80)
                    return write_buffer + f"{TAB*4}</p>\n{TAB*3}</div>"

                if val in ["\n", "\\n"]:
                    char_count += 1
                    print(f"{char_count}. {addr} ==> 《換換行》\n" + "-" * 80)
                    write_buffer += f"{TAB*4}</p><p>\n"
                    char_count = 0
                    break

                str_val = str(val).strip() if val else ""
                if is_punctuation(str_val):
                    msg = f"{str_val}【標點符號】"
                    write_buffer += f"{TAB*5}<span>{str_val}</span>\n"
                elif not str_val:
                    msg = "【空白】"
                    write_buffer += f"{TAB*5}<span>　</span>\n"
                elif not is_han_ji(val):
                    msg = f"{str_val}【其他字元】"
                    write_buffer += f"{TAB*5}<span>{str_val}</span>\n"
                else:
                    han_ji_cell = sheet.range((row, col))
                    tlpa = sheet.range((row - 1, col)).value
                    tlpa = str(tlpa).strip() if tlpa else ""
                    pian_tiau = self._is_pian_tiau_cell(han_ji_cell)
                    if not tlpa:
                        msg = f"{str_val}【無音標】"
                        write_buffer += f"{TAB*5}<span>{str_val}</span>\n"
                    else:
                        if kam_si_u_tiau_hu(tlpa):
                            tlpa = tng_tiau_ho(tng_im_piau(tlpa))
                        tag, siong, zian = self.generate_ruby_tag(
                            str_val, tlpa, pian_tiau=pian_tiau
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
                    # write_buffer += "<br/>\n"
                    write_buffer += "</p><p>\n"
                    char_count = 0
                    print("《人工斷行》")

        # write_buffer += f"\n{TAB*3}</p></div>"
        # return write_buffer


def _create_html_file(program, output_path, head_extra="", title="", content=""):
    web_page_stem = Path(output_path).stem
    img_url = str(program.image_url or "").strip()
    if img_url == "None" or not img_url:
        img_url = "king_tian.png"  # 預設圖片

    # 修正：恢復相對路徑，確保 Local 端可瀏覽
    full_img = img_url if img_url.startswith("http") else f"./assets/images/{img_url}"

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
            <div style='text-align: center'>
                <img src='{full_img}' width='800' />
            </div>
            {content}
        </article>
    </main>
    <a href="index.html" class="floating-home-btn">🏠</a>
</body>
</html>"""
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(template)


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
        siong = (
            str(program.siong_pinn_piau_im)
            if program.siong_pinn_piau_im and str(program.siong_pinn_piau_im) != "None"
            else ""
        )
        zian = (
            str(program.zian_pinn_piau_im)
            if program.zian_pinn_piau_im and str(program.zian_pinn_piau_im) != "None"
            else ""
        )

        if han_ji_piau_im_format == "無預設":
            im_piau = piau_im_huat
        else:
            if siong and zian:
                im_piau = f"{siong}＋{zian}"
            elif siong:
                im_piau = siong
            elif zian:
                im_piau = zian
            else:
                im_piau = piau_im_huat

        output_file = f"《{program.title}》【{ue_im_lui_piat}】{im_piau}.html"
        output_path = os.path.join("docs", output_file)
        os.makedirs("docs", exist_ok=True)

        # 處理 Meta 標籤中的 None
        meta_keys = [
            "TITLE",
            "IMAGE_URL",
            "網頁格式",
            "上邊標音",
            "右邊標音",
            "使用變調標音格式",
        ]
        meta_list = []
        for k in meta_keys:
            v = get_value_by_name(wb, k)
            if v is None or str(v) == "None":
                v = ""
            meta_list.append(f'<meta name="{k}" content="{v}" />')
        head_extra = "\n    ".join(meta_list)

        if xls_cell.iong_bian_tiau_piau_im_format:
            logging_process_step(
                "【使用變調標音格式】=是：<rt> 依【上邊標音】轉成本調／變調格式；"
                "<rtc> 依【右邊標音】（標音方式含「右」時）。"
            )

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
