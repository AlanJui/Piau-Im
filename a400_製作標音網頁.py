"""
a400_製作標音網頁.py V0.2.2.8
"""

# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import logging
import os
from pathlib import Path

# 載入第三方套件
import xlwings as xw

# 載入自訂模組
from mod_excel_access import get_value_by_name
from mod_logging import (
    init_logging,
    logging_exc_error,  # noqa: F401
    logging_exception,  # noqa: F401
    logging_process_step,  # noqa: F401
    logging_warning,  # noqa: F401
)
from mod_帶調符音標 import is_han_ji, kam_si_u_tiau_hu, tng_im_piau, tng_tiau_ho
from mod_標音 import is_punctuation, split_tai_gi_im_piau
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

# =========================================================================
# 設定日誌
# =========================================================================
init_logging()


# =========================================================================
# 自訂 ExcelCell 子類別：覆蓋特定方法以實現萌典查詢功能
# =========================================================================
class CellProcessor(ExcelCell):
    """
    個人字典查詢專用的儲存格處理器
    繼承自 ExcelCell
    覆蓋以下方法以實現個人字典查詢功能：
    - _process_sheet(): 處理整個工作表
    """

    def __init__(
        self,
        program: Program,
        new_jin_kang_piau_im_ji_khoo_sheet: bool = False,
        new_piau_im_ji_khoo_sheet: bool = False,
        new_khuat_ji_piau_sheet: bool = False,
    ):
        # 調用父類別（MengDianExcelCell）的建構子
        super().__init__(
            program=program,
            new_jin_kang_piau_im_ji_khoo_sheet=new_jin_kang_piau_im_ji_khoo_sheet,
            new_piau_im_ji_khoo_sheet=new_piau_im_ji_khoo_sheet,
            new_khuat_ji_piau_sheet=new_khuat_ji_piau_sheet,
        )

        # 網頁相關
        wb = self.program.wb
        self.title = get_value_by_name(wb=wb, name="TITLE")
        self.image_url = get_value_by_name(wb=wb, name="IMAGE_URL")
        self.output_dir = "docs"
        self.web_title = str(self.title) if self.title else "網頁標題"
        self.total_chars_per_line = get_value_by_name(wb=wb, name="網頁每列字數")
        self.total_chars_per_line = (
            0 if self.total_chars_per_line is None else int(self.total_chars_per_line)
        )

        # 標音相關
        self.han_ji_piau_im_format = get_value_by_name(wb=wb, name="網頁格式")
        self.piau_im_hong_sik = get_value_by_name(wb=wb, name="標音方式")
        self.siong_pinn_piau_im = get_value_by_name(wb=wb, name="上邊標音")
        self.zian_pinn_piau_im = get_value_by_name(wb=wb, name="右邊標音")

        # 聲調符號對映表
        self.zu_im_huat_list = {
            "SNI": ["fifteen_yin", "rt", "十五音切語"],
            "TPS": ["Piau_Im", "rt", "方音符號注音"],
            "MPS2": ["Piau_Im", "rt", "注音二式"],
            "POJ": ["pin_yin", "rt", "白話字拼音"],
            "TL": ["pin_yin", "rt", "台羅拼音"],
            "BP": ["pin_yin", "rt", "閩拼標音"],
            "TLPA_Plus": ["pin_yin", "rt", "台羅改良式"],
            "DBL": ["Siang_Pai", "rtc", "雙排注音"],
            "無預設": ["Siang_Pai", "rtc", "雙排注音"],
        }

    # =================================================================
    # 輔助方法
    # =================================================================
    def _build_full_ruby_tag(
        self, han_ji: str, siong_piau_im: str, zian_piau_im: str
    ) -> str:
        """構建 Ruby 標籤"""
        if siong_piau_im != "" and zian_piau_im == "":
            html_str = (
                "<ruby>\n"
                "  <rb>%s</rb>\n"  # 漢字
                "  <rp>(</rp>\n"  # 括號左
                "  <rt>%s</rt>\n"  # 上方標音
                "  <rp>)</rp>\n"  # 括號右
                "</ruby>\n"
            )
            return html_str % (han_ji, siong_piau_im)
        elif siong_piau_im == "" and zian_piau_im != "":
            html_str = (
                "<ruby>\n"
                "  <rb>%s</rb>\n"  # 漢字
                "  <rp>(</rp>\n"  # 括號左
                "  <rtc>%s</rtc>\n"  # 右方標音
                "  <rp>)</rp>\n"  # 括號右
                "</ruby>\n"
            )
            return html_str % (han_ji, zian_piau_im)
        elif siong_piau_im != "" and zian_piau_im != "":
            html_str = (
                "<ruby>\n"
                "  <rb>%s</rb>\n"  # 漢字
                "  <rp>(</rp>\n"  # 括號左
                "  <rt>%s</rt>\n"  # 上方標音
                "  <rp>)</rp>\n"  # 括號右
                "  <rp>(</rp>\n"  # 括號左
                "  <rtc>%s</rtc>\n"  # 右方標音
                "  <rp>)</rp>\n"  # 括號右
                "</ruby>\n"
            )
            return html_str % (han_ji, siong_piau_im, zian_piau_im)
        else:
            return f"<span>{han_ji}</span>\n"

    def _build_ruby_tag(
        self, han_ji: str, siong_piau_im: str, zian_piau_im: str
    ) -> str:
        """構建 Ruby 標籤"""
        if siong_piau_im != "" and zian_piau_im == "":
            # 只有上方標音：羅馬拼音、白話字、台羅、閩拼等
            html_str = (
                "<ruby>\n"
                "  %s\n"  # 漢字
                "  <rt>%s</rt>\n"  # 上方標音
                "</ruby>\n"
            )
            return html_str % (han_ji, siong_piau_im)
        elif siong_piau_im == "" and zian_piau_im != "":
            # 只有右方標音：方音符號（TPS）
            html_str = (
                "<ruby>\n"
                "  %s\n"  # 漢字
                "  <rtc>%s</rtc>\n"  # 右方標音
                "</ruby>\n"
            )
            return html_str % (han_ji, zian_piau_im)
        elif siong_piau_im != "" and zian_piau_im != "":
            # 同時有上方及右方標音：雙排注音（DBL）
            html_str = (
                "<ruby>\n"
                "  %s\n"  # 漢字
                "  <rt>%s</rt>\n"  # 上方標音
                "  <rtc>%s</rtc>\n"  # 右方標音
                "</ruby>\n"
            )
            return html_str % (han_ji, siong_piau_im, zian_piau_im)
        else:
            return f"<span>{han_ji}</span>\n"

    def generate_ruby_tag(self, han_ji: str, tai_gi_im_piau: str) -> tuple:
        """
        生成 Ruby 標籤

        Args:
            han_ji: 漢字
            tai_gi_im_piau: 台語音標

        Returns:
            (ruby_tag, siong_piau_im, zian_piau_im)
        """
        # 檢查台語音標是否為空或 None
        if not tai_gi_im_piau or (
            isinstance(tai_gi_im_piau, str) and tai_gi_im_piau.strip() == ""
        ):
            # 如果沒有音標，返回不帶標音的 span
            ruby_tag = f"  <span>{han_ji}</span>\n"
            return ruby_tag, "", ""

        # 解構【台語音標】=【聲母】+【韻母】+【調號】
        try:
            zu_im_list = split_tai_gi_im_piau(tai_gi_im_piau)
        except (IndexError, AttributeError) as e:
            # 如果解構失敗，返回不帶標音的 span
            print(f"警告：無法解構音標 '{tai_gi_im_piau}' for 漢字 '{han_ji}': {e}")
            ruby_tag = f"  <span>{han_ji}</span>\n"
            return ruby_tag, "", ""

        # 零聲母處理
        if zu_im_list[0] == "" or zu_im_list[0] is None:
            siann_bu = "ø"  # 無聲母: ø
        else:
            siann_bu = zu_im_list[0]

        siong_piau_im = ""
        zian_piau_im = ""

        # 根據【網頁格式】，決定【漢字】之上方或右方，是否該顯示【標音】
        if self.han_ji_piau_im_format == "無預設":
            # 根據【標音方式】決定漢字之上方及右方，是否需要放置標音
            if self.piau_im_hong_sik == "上及右":
                siong_piau_im = self.program.piau_im.han_ji_piau_im_tng_huan(
                    piau_im_huat=self.siong_pinn_piau_im,
                    siann_bu=siann_bu,
                    un_bu=zu_im_list[1],
                    tiau_ho=zu_im_list[2],
                )
                zian_piau_im = self.program.piau_im.han_ji_piau_im_tng_huan(
                    piau_im_huat=self.zian_pinn_piau_im,
                    siann_bu=siann_bu,
                    un_bu=zu_im_list[1],
                    tiau_ho=zu_im_list[2],
                )
            elif self.piau_im_hong_sik == "上":
                siong_piau_im = self.program.piau_im.han_ji_piau_im_tng_huan(
                    piau_im_huat=self.siong_pinn_piau_im,
                    siann_bu=siann_bu,
                    un_bu=zu_im_list[1],
                    tiau_ho=zu_im_list[2],
                )
            elif self.piau_im_hong_sik == "右":
                zian_piau_im = self.program.piau_im.han_ji_piau_im_tng_huan(
                    piau_im_huat=self.zian_pinn_piau_im,
                    siann_bu=siann_bu,
                    un_bu=zu_im_list[1],
                    tiau_ho=zu_im_list[2],
                )
        else:
            # 按指定網頁格式設定標音位置
            if self.han_ji_piau_im_format in ["POJ", "TL", "BP", "TLPA_Plus", "SNI"]:
                siong_piau_im = self.program.piau_im.han_ji_piau_im_tng_huan(
                    piau_im_huat=self.siong_pinn_piau_im,
                    siann_bu=siann_bu,
                    un_bu=zu_im_list[1],
                    tiau_ho=zu_im_list[2],
                )
            elif self.han_ji_piau_im_format == "TPS":
                zian_piau_im = self.program.piau_im.han_ji_piau_im_tng_huan(
                    piau_im_huat=self.zian_pinn_piau_im,
                    siann_bu=siann_bu,
                    un_bu=zu_im_list[1],
                    tiau_ho=zu_im_list[2],
                )
            elif self.han_ji_piau_im_format == "DBL":
                siong_piau_im = self.program.piau_im.han_ji_piau_im_tng_huan(
                    piau_im_huat=self.siong_pinn_piau_im,
                    siann_bu=siann_bu,
                    un_bu=zu_im_list[1],
                    tiau_ho=zu_im_list[2],
                )
                zian_piau_im = self.program.piau_im.han_ji_piau_im_tng_huan(
                    piau_im_huat=self.zian_pinn_piau_im,
                    siann_bu=siann_bu,
                    un_bu=zu_im_list[1],
                    tiau_ho=zu_im_list[2],
                )

        # 根據標音方式，設定 Ruby Tag
        ruby_tag = self._build_ruby_tag(han_ji, siong_piau_im, zian_piau_im)

        return ruby_tag, siong_piau_im, zian_piau_im

    def generate_title_with_ruby(self) -> str:
        """
        取得文章標題並加注【台語音標】

        Returns:
            含有 Ruby 標籤的標題 HTML
        """
        wb = self.program.wb
        sheet = wb.sheets["漢字注音"]

        title_chars = ""
        tlpa_im_piau_list = []
        title_with_ruby = ""
        row, col = 5, 4

        if sheet.range((row, col)).value == "《":
            while True:
                cell_val = sheet.range((row, col)).value
                if cell_val is None:
                    break
                title_chars += cell_val
                if cell_val != "《" and cell_val != "》":
                    tlpa_im_piau_list.append(sheet.range((row - 1, col)).value)
                if cell_val == "》":
                    break
                col += 1
                if col > 18:  # 超出一列最大值（R欄=18），換行至下一列
                    row += 1
                    col = 4

        if title_chars:
            # 去除《與》符號，只傳入標題文字本體加注 ruby
            title_han_ji = title_chars.replace("《", "").replace("》", "")
            i = 0
            for han_ji in title_han_ji:
                tai_gi_im_piau = ""
                # han_ji_piau_im = ""
                siong_piau_im = ""
                zian_piau_im = ""

                if han_ji.strip() == "":
                    i += 1
                    continue
                elif han_ji == "\n":
                    # 若讀到換行字元，則直接輸出換行標籤
                    tag = "<br/>\n"
                    title_with_ruby += tag
                elif not is_han_ji(han_ji):
                    tag = f"<span>{han_ji}</span>"
                    title_with_ruby += tag
                else:
                    # 取得對應的台語音標
                    tai_gi_im_piau = (
                        tlpa_im_piau_list[i] if i < len(tlpa_im_piau_list) else ""
                    )
                    # 將【漢字】及【上方】/【右方】標音合併成一個 Ruby Tag
                    ruby_tag, siong_piau_im, zian_piau_im = self.generate_ruby_tag(
                        han_ji=han_ji, tai_gi_im_piau=tai_gi_im_piau
                    )
                    title_with_ruby += ruby_tag
                    msg = f"{han_ji} [{tai_gi_im_piau}] ==》 上方標音：{siong_piau_im} / 右方標音：{zian_piau_im}"
                    print(msg)
                i += 1

        return title_with_ruby

    # =================================================================
    # 覆蓋父類別的方法
    # =================================================================
    def _process_sheet(self, sheet) -> str:
        """
        處理工作表內容並生成 HTML 網頁內容

        Args:
            sheet: Excel 工作表物件

        Returns:
            HTML 網頁內容字串
        """
        write_buffer = ""

        # 輸出放置圖片的 HTML Tag
        div_image = (
            "<div class='separator' style='clear: both'>\n"
            "  <a href='圖片' style='display: block; padding: 1em 0; text-align: center'>\n"
            "    <img alt='%s' border='0' width='800' \n"
            "      src='%s' />\n"
            "  </a>\n"
            "</div>\n"
        )
        # if self.image_url:
        #     write_buffer += div_image % (self.web_title, self.image_url)
        image_url = self.program.image_url.strip()
        if image_url.lower().startswith(("http://", "https://")):
            full_image_url = image_url
        else:
            full_image_url = f"./assets/images/{image_url}"
        write_buffer += div_image % (self.web_title, full_image_url)

        # 輸出【文章】Div tag 及【文章標題】Ruby Tag
        title_with_ruby = self.generate_title_with_ruby()
        pai_ban_iong_huat = self.zu_im_huat_list[self.han_ji_piau_im_format][0]

        if title_with_ruby:
            div_tag = (
                "<div class='%s'>\n"
                "  <p class='title'>\n"
                "    <span>《</span>\n"
                "    %s\n"
                "    <span>》</span>\n"
                "  </p>\n"
            )
        else:
            div_tag = "<div class='%s'>\n" "  <p class='title'>\n" "    %s\n" "  </p>\n"
        write_buffer += div_tag % (pai_ban_iong_huat, title_with_ruby)
        write_buffer += "<p>\n"

        # 逐列處理工作表內容
        program = self.program
        End_Of_File = False
        char_count = 0
        total_chars_per_line = self.total_chars_per_line  # 網頁每列字數
        if total_chars_per_line == 0:
            total_chars_per_line = 0

        # 工作表起始列號 = 範圍起始列號 + 漢字列偏移量 = 3 + 2 = 5
        start_row = program.line_start_row + program.han_ji_row_offset
        end_row = program.line_end_row + program.han_ji_row_offset

        for row in range(start_row, end_row, program.ROWS_PER_LINE):
            if title_with_ruby and row == start_row:
                # 已經處理過標題列，跳過
                continue
            sheet.range((row, program.start_col)).select()

            for col in range(program.start_col, program.end_col):
                cell_value = sheet.range((row, col)).value

                if cell_value == "φ":  # 讀到【結尾標示】
                    End_Of_File = True
                    msg = "《文章終止》"
                    print(
                        f"{char_count}. {xw.utils.col_name(col)}{row} = ({row}, {col}) ==> {msg}"
                    )
                    break

                elif cell_value == "\n":  # 讀到【換行標示】
                    write_buffer += "</p><p>\n"
                    msg = "《換行》"
                    char_count = 0
                    print(
                        f"{char_count}. {xw.utils.col_name(col)}{row} = ({row}, {col}) ==> {msg}"
                    )
                    break

                else:
                    # 處理儲存格內容（先檢查標點符號，再檢查漢字）
                    str_value = str(cell_value).strip() if cell_value else ""

                    # 先檢查是否為標點符號或其他非漢字字元
                    if is_punctuation(str_value):
                        msg = f"{str_value}【標點符號】"
                        ruby_tag = f"  <span>{str_value}</span>\n"
                        write_buffer += ruby_tag
                        char_count += 1
                        print(
                            f"{char_count}. {xw.utils.col_name(col)}{row} = ({row}, {col}) ==> {msg}"
                        )
                    elif str_value == "":
                        msg = "【空白】"
                        ruby_tag = "  <span>　</span>\n"
                        write_buffer += ruby_tag
                        char_count += 1
                        print(
                            f"{char_count}. {xw.utils.col_name(col)}{row} = ({row}, {col}) ==> {msg}"
                        )
                    elif not is_han_ji(cell_value):
                        msg = f"{str_value}【其他字元】"
                        ruby_tag = f"  <span>{str_value}</span>\n"
                        write_buffer += ruby_tag
                        char_count += 1
                        print(
                            f"{char_count}. {xw.utils.col_name(col)}{row} = ({row}, {col}) ==> {msg}"
                        )
                    else:
                        # 處理漢字
                        han_ji = cell_value.strip()
                        tai_gi_im_piau = sheet.range((row - 1, col)).value
                        tai_gi_im_piau = (
                            str(tai_gi_im_piau).strip()
                            if tai_gi_im_piau is not None
                            else ""
                        )

                        # 檢查音標是否為空
                        if not tai_gi_im_piau or tai_gi_im_piau == "":
                            print(
                                f"警告：漢字 '{han_ji}' 在 ({row}, {col}) 沒有台語音標"
                            )
                            ruby_tag = f"  <span>{han_ji}</span>\n"
                            write_buffer += ruby_tag
                            char_count += 1
                            print(
                                f"{char_count}. {xw.utils.col_name(col)}{row} = ({row}, {col}) ==> {han_ji}【無音標】"
                            )
                            continue

                        # 【去調符】轉換為【帶調號之台語音標】
                        if kam_si_u_tiau_hu(tai_gi_im_piau):
                            tai_gi_im_piau = tng_im_piau(tai_gi_im_piau)
                            tlpa_im_piau = tng_tiau_ho(tai_gi_im_piau)
                        else:
                            tlpa_im_piau = tai_gi_im_piau

                        ruby_tag, siong_piau_im, zian_piau_im = self.generate_ruby_tag(
                            han_ji=han_ji,
                            tai_gi_im_piau=tlpa_im_piau,
                        )
                        write_buffer += ruby_tag
                        msg = f"{han_ji} [{tlpa_im_piau}] ==》 上方標音：{siong_piau_im} / 右方標音：{zian_piau_im}"
                        char_count += 1
                        print(
                            f"{char_count}. {xw.utils.col_name(col)}{row} = ({row}, {col}) ==> {msg}"
                        )

                # 檢查是否需要插入換行標籤
                if total_chars_per_line != 0 and char_count >= total_chars_per_line:
                    write_buffer += "<br/>\n"
                    char_count = 0
                    print("《人工斷行》")

            if End_Of_File:
                break

        write_buffer += "</p></div>"
        return write_buffer


# =========================================================================
# 主要處理函數
# =========================================================================
def _create_html_file(
    program: Program,
    output_path: str,
    content: str,
    title: str = "您的標題",
    head_extra: str = "",
):
    """
        創建 HTML 檔案
    <meta content='https://alanjui.github.io/Piau-Im//《深慮論》【文讀音】閩拼調符.html' property='og:url' />
    <meta content='《深慮論》【文讀音：閩拼+方音符號】' property='og:title' />
    <meta content='《深慮論》明朝：方孝孺' property='og:description' />

    file_path = output_path
    parent_directory = Path(r"{file_path}").parent
    file_name = Path(r"{file_path}").name
    main_file_name = Path(r"{file_path}").stem
    file_extension = Path(r"{file_path}").suffix
    """

    # 取得網頁主檔案名稱（不含路徑及副檔名）
    web_page_main_file_name = Path(output_path).stem
    # 取得 Excel 檔案名稱（不含路徑及副檔名）
    excel_file_stem = program.excel_file_stem

    # 取得圖片 URL
    # 判斷 image_url 是否為完整 URL (http/https 開頭)
    image_url = str(program.image_url or "").strip()
    if image_url.lower().startswith(("http://", "https://")):
        full_image_url = image_url
    else:
        full_image_url = f"https://alanjui.github.io/Piau-Im/assets/images/{image_url}"

    template = f"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{title}</title>
    <meta content='https://alanjui.github.io/Piau-Im/{web_page_main_file_name}.html' property='og:url' />
    <meta content='{excel_file_stem}' property='og:title' />
    <meta content='{title}' property='og:description' />
    <meta content='{full_image_url}' property='og:image' />
    {head_extra}
    <link rel="stylesheet" href="assets/styles/styles.css">
</head>
<body>
    <main class="page">
        <article class="article_content">
        {content}
        </article>
    </main>
</body>
</html>
    """
    with open(output_path, "w", encoding="utf-8") as file:
        file.write(template)
    print(f"\n輸出網頁檔案：{output_path}")


# =========================================================================
# 作業程序
# =========================================================================
def process(wb, args) -> int:
    """
    查詢漢字讀音並標注

    Args:
        wb: Excel Workbook 物件
        args: 命令列參數

    Returns:
        處理結果代碼
    """
    # --------------------------------------------------------------------------
    # 作業初始化
    # --------------------------------------------------------------------------
    logging_process_step("<=========== 作業開始！==========>")

    try:
        # --------------------------------------------------------------------------
        # 初始化 Program 配置
        # --------------------------------------------------------------------------
        program = Program(wb=wb, args=args, hanji_piau_im_sheet_name="漢字注音")

        # 建立萌典專用的儲存格處理器（繼承自 ExcelCell）
        xls_cell = CellProcessor(
            program=program,
            new_jin_kang_piau_im_ji_khoo_sheet=(
                args.new if hasattr(args, "new") else False
            ),
            new_piau_im_ji_khoo_sheet=args.new if hasattr(args, "new") else False,
            new_khuat_ji_piau_sheet=args.new if hasattr(args, "new") else False,
        )
    except Exception as e:
        logging_exc_error(msg="處理作業異常！", error=e)
        return EXIT_CODE_PROCESS_FAILURE

    # --------------------------------------------------------------------------
    # 作業處理中
    # --------------------------------------------------------------------------
    try:
        # 處理工作表
        # sheet_name = program.hanji_piau_im_sheet_name
        # sheet = wb.sheets[sheet_name]
        sheet = wb.sheets["漢字注音"]
        sheet.activate()

        # 產生 HTML 網頁內容
        print("開始製作【漢字注音】網頁！")
        html_content = xls_cell._process_sheet(sheet=sheet)

        # 生成輸出檔案名稱
        piau_im_huat = program.piau_im_huat
        ue_im_lui_piat = program.ue_im_lui_piat
        han_ji_piau_im_format = program.han_ji_piau_im_format
        if han_ji_piau_im_format == "無預設":
            im_piau = piau_im_huat
        elif han_ji_piau_im_format == "上":
            im_piau = program.siong_pinn_piau_im
        elif han_ji_piau_im_format == "右":
            im_piau = program.zian_pinn_piau_im
        else:
            im_piau = f"{program.siong_pinn_piau_im}＋{program.zian_pinn_piau_im}"

        title = program.title
        output_file = f"《{title}》【{ue_im_lui_piat}】{im_piau}.html"
        output_path = os.path.join(xls_cell.output_dir, output_file)

        # 生成標準 Excel 檔案名稱
        new_excel_file_name = Program.generate_new_excel_file_name(wb=wb)
        if not new_excel_file_name:
            logging_exc_error(
                msg="無法依【檔案命名標準】生成另存新檔之 Excel 檔案名稱!"
            )
            return EXIT_CODE_PROCESS_FAILURE
        else:
            print(
                f"作業結束時，將以 Excel 檔案名稱：【{new_excel_file_name}】另存新檔。"
            )

        # 確保輸出目錄存在
        os.makedirs(xls_cell.output_dir, exist_ok=True)

        # 取得 env 工作表的設定值並組合 meta 標籤字串
        env_keys = [
            "FILE_NAME",
            "TITLE",
            "IMAGE_URL",
            "OUTPUT_PATH",
            "章節序號",
            "顯示注音輸入",
            "每頁總列數",
            "每列總字數",
            "語音類型",
            "漢字庫",
            "標音方法",
            "網頁格式",
            "標音方式",
            "上邊標音",
            "右邊標音",
            "網頁每列字數",
        ]
        head_extra = "\n"
        for key in env_keys:
            value = get_value_by_name(wb, key)
            head_extra += f'    <meta name="{key}" content="{value}" />\n'

        # 輸出到網頁檔案
        _create_html_file(
            program, output_path, html_content, xls_cell.web_title, head_extra
        )

        logging_process_step("【漢字注音】工作表轉製網頁作業完畢！")

    except Exception as e:
        logging_exc_error(msg="處理作業異常！", error=e)
        return EXIT_CODE_PROCESS_FAILURE

    # --------------------------------------------------------------------------
    # 處理作業結束
    # --------------------------------------------------------------------------
    logging_process_step("<=========== 作業結束！==========>")
    return EXIT_CODE_SUCCESS


# =========================================================================
# 主程式
# =========================================================================
def main(args) -> int:
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
    # 取得【作用中活頁簿】
    wb = None
    try:
        wb = xw.apps.active.books.active  # 取得 Excel 作用中的活頁簿檔案
    except Exception as e:
        msg = "無法找到作用中的 Excel 工作簿！"
        logging_exception(msg=msg, error=e)
        return EXIT_CODE_NO_FILE

    # 若無法取得【作用中活頁簿】，則因無法繼續作業，故返回【作業異常終止代碼】結束。
    if not wb:
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
    logging_process_step(f"《========== 程式終止執行：{program_name} ==========》")
    return EXIT_CODE_SUCCESS


# =============================================================================
# 測試程式
# =============================================================================
def test_01():
    """
    測試程式主要作業流程
    """
    print("\n\n")
    print("=" * 100)
    print("執行測試：test_01()")
    # 執行主要作業流程
    return EXIT_CODE_SUCCESS


# =============================================================================
# 程式作業模式切換
# =============================================================================
if __name__ == "__main__":
    import argparse
    import sys

    # 解析命令行參數
    parser = argparse.ArgumentParser(
        description="缺字表修正後續作業程式",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用範例：
  python a000.py          # 執行一般模式
  python a000.py -new     # 建立新的字庫工作表
  python a000.py -test    # 執行測試模式
""",
    )
    parser.add_argument(
        "--test",
        action="store_true",
        help="執行測試模式",
    )
    parser.add_argument(
        "--new",
        action="store_true",
        help="建立新的標音字庫工作表",
    )
    args = parser.parse_args()

    if args.test:
        # 執行測試
        test_01()
    else:
        # 從 Excel 呼叫
        exit_code = main(args)
        if exit_code != EXIT_CODE_SUCCESS:
            print(f"程式異常終止，返回代碼：{exit_code}")
            sys.exit(exit_code)
