# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import argparse
import logging
import os
import re
from pathlib import Path
from typing import Tuple

# 載入第三方套件
import xlwings as xw
from dotenv import load_dotenv

# 載入自訂模組
from mod_BP_tng_huan import convert_bp_im_piau_to_zu_im
from mod_BP_tng_huan_ping_im import convert_TLPA_to_BP
from mod_ca_ji_tian import HanJiTian  # 新的查字典模組
from mod_excel_access import (
    clear_han_ji_kap_piau_im,
    delete_sheet_by_name,
    get_value_by_name,
    reset_cells_format_in_sheet,
)
from mod_file_access import save_as_new_file
from mod_logging import (
    init_logging,
    logging_exc_error,
    logging_exception,
    logging_process_step,
    logging_warning,  # noqa: F401
)
from mod_字庫 import JiKhooDict
from mod_帶調符音標 import is_han_ji, kam_si_u_tiau_hu, read_text_with_han_ji, tng_im_piau, tng_tiau_ho
from mod_標音 import (
    PiauIm,
    ca_ji_tng_piau_im,
    convert_tl_with_tiau_hu_to_tlpa,
    is_punctuation,
    split_hong_im_hu_ho,
    split_tai_gi_im_piau,
    tlpa_tng_han_ji_piau_im,
)
from mod_程式 import ExcelCell, Program

# =========================================================================
# 載入環境變數
# =========================================================================
load_dotenv()

DB_HO_LOK_UE = os.getenv('DB_HO_LOK_UE', 'Ho_Lok_Ue.db')
DB_KONG_UN = os.getenv('DB_KONG_UN', 'Kong_Un.db')

# =========================================================================
# 常數定義
# =========================================================================
# 定義 Exit Code
EXIT_CODE_SUCCESS = 0
EXIT_CODE_NO_FILE = 1
EXIT_CODE_INVALID_INPUT = 2
EXIT_CODE_SAVE_FAILURE = 3
EXIT_CODE_PROCESS_FAILURE = 10
EXIT_CODE_UNKNOWN_ERROR = 99

# 標音位置：-1 ==> 自動標音；-2 ==> 人工標音
PIAU_IM_ROW = -1

# =========================================================================
# 設定日誌
# =========================================================================
init_logging()

# =========================================================================
# 主要處理函數
# =========================================================================
class WebPageConfig:
    """網頁製作配置資料類別"""

    def __init__(self, wb):
        # Excel 相關
        self.TOTAL_LINES = int(wb.names['每頁總列數'].refers_to_range.value)
        self.CHARS_PER_ROW = int(wb.names['每列總字數'].refers_to_range.value)
        self.ROWS_PER_LINE = 4
        self.start_row = 5
        self.start_col = 4
        self.end_row = self.start_row + (self.TOTAL_LINES * self.ROWS_PER_LINE)
        self.end_col = self.start_col + self.CHARS_PER_ROW

        # 標音相關
        self.han_ji_khoo_name = get_value_by_name(wb=wb, name='漢字庫')
        self.ue_im_lui_piat = get_value_by_name(wb=wb, name='語音類型')
        self.piau_im_huat = get_value_by_name(wb=wb, name='標音方法')
        self.piau_im_format = get_value_by_name(wb=wb, name='網頁格式')
        self.piau_im_hong_sik = get_value_by_name(wb=wb, name='標音方式')
        self.siong_pinn_piau_im = get_value_by_name(wb=wb, name='上邊標音')
        self.zian_pinn_piau_im = get_value_by_name(wb=wb, name='右邊標音')

        # 網頁相關
        self.title = get_value_by_name(wb=wb, name='TITLE')
        self.image_url = get_value_by_name(wb=wb, name='IMAGE_URL')
        self.output_dir = 'docs'
        self.web_title = str(self.title) if self.title else '網頁標題'

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


class WebPageGenerator:
    """網頁生成器"""

    def __init__(self, config: WebPageConfig, piau_im: PiauIm, ji_tian: HanJiTian):
        self.config = config
        self.piau_im = piau_im
        self.ji_tian = ji_tian

    def generate_ruby_tag(self, wb, han_ji: str, tai_gi_im_piau: str) -> tuple:
        """
        生成 Ruby 標籤

        Args:
            han_ji: 漢字
            tai_gi_im_piau: 台語音標

        Returns:
            (ruby_tag, siong_piau_im, zian_piau_im)
        """
        zu_im_list = split_tai_gi_im_piau(tai_gi_im_piau)

        # 零聲母處理
        if zu_im_list[0] == "" or zu_im_list[0] is None:
            siann_bu = "ø"  # 無聲母: ø
        else:
            siann_bu = zu_im_list[0]

        siong_piau_im = ""
        zian_piau_im = ""

        # 根據【網頁格式】，決定【漢字】之上方或右方，是否該顯示【標音】
        if self.config.piau_im_format == "無預設":
            # 根據【標音方式】決定漢字之上方及右方，是否需要放置標音
            if self.config.piau_im_hong_sik == "上及右":
                siong_piau_im = self.piau_im.han_ji_piau_im_tng_huan(
                    piau_im_huat=self.config.siong_pinn_piau_im,
                    siann_bu=siann_bu,
                    un_bu=zu_im_list[1],
                    tiau_ho=zu_im_list[2],
                )
                zian_piau_im = self.piau_im.han_ji_piau_im_tng_huan(
                    piau_im_huat=self.config.zian_pinn_piau_im,
                    siann_bu=siann_bu,
                    un_bu=zu_im_list[1],
                    tiau_ho=zu_im_list[2],
                )
            elif self.config.piau_im_hong_sik == "上":
                siong_piau_im = self.piau_im.han_ji_piau_im_tng_huan(
                    piau_im_huat=self.config.siong_pinn_piau_im,
                    siann_bu=siann_bu,
                    un_bu=zu_im_list[1],
                    tiau_ho=zu_im_list[2],
                )
            elif self.config.piau_im_hong_sik == "右":
                zian_piau_im = self.piau_im.han_ji_piau_im_tng_huan(
                    piau_im_huat=self.config.zian_pinn_piau_im,
                    siann_bu=siann_bu,
                    un_bu=zu_im_list[1],
                    tiau_ho=zu_im_list[2],
                )
        else:
            # 按指定網頁格式設定標音位置
            if self.config.piau_im_format in ["POJ", "TL", "BP", "TLPA_Plus", "SNI"]:
                siong_piau_im = self.piau_im.han_ji_piau_im_tng_huan(
                    piau_im_huat=self.config.siong_pinn_piau_im,
                    siann_bu=siann_bu,
                    un_bu=zu_im_list[1],
                    tiau_ho=zu_im_list[2],
                )
            elif self.config.piau_im_format == "TPS":
                zian_piau_im = self.piau_im.han_ji_piau_im_tng_huan(
                    piau_im_huat=self.config.zian_pinn_piau_im,
                    siann_bu=siann_bu,
                    un_bu=zu_im_list[1],
                    tiau_ho=zu_im_list[2],
                )
            elif self.config.piau_im_format == "DBL":
                siong_piau_im = self.piau_im.han_ji_piau_im_tng_huan(
                    piau_im_huat=self.config.siong_pinn_piau_im,
                    siann_bu=siann_bu,
                    un_bu=zu_im_list[1],
                    tiau_ho=zu_im_list[2],
                )
                zian_piau_im = self.piau_im.han_ji_piau_im_tng_huan(
                    piau_im_huat=self.config.zian_pinn_piau_im,
                    siann_bu=siann_bu,
                    un_bu=zu_im_list[1],
                    tiau_ho=zu_im_list[2],
                )

        # 根據標音方式，設定 Ruby Tag
        ruby_tag = self._build_ruby_tag(han_ji, siong_piau_im, zian_piau_im)

        return ruby_tag, siong_piau_im, zian_piau_im

    def _build_ruby_tag(
        self, han_ji: str, siong_piau_im: str, zian_piau_im: str
    ) -> str:
        """構建 Ruby 標籤"""
        if siong_piau_im != "" and zian_piau_im == "":
            html_str = (
                "<ruby>\n"
                "  <rb>%s</rb>\n"   # 漢字
                "  <rp>(</rp>\n"    # 括號左
                "  <rt>%s</rt>\n"   # 上方標音
                "  <rp>)</rp>\n"    # 括號右
                "</ruby>\n"
            )
            return html_str % (han_ji, siong_piau_im)
        elif siong_piau_im == "" and zian_piau_im != "":
            html_str = (
                "<ruby>\n"
                "  <rb>%s</rb>\n"   # 漢字
                "  <rp>(</rp>\n"    # 括號左
                "  <rtc>%s</rtc>\n" # 右方標音
                "  <rp>)</rp>\n"    # 括號右
                "</ruby>\n"
            )
            return html_str % (han_ji, zian_piau_im)
        elif siong_piau_im != "" and zian_piau_im != "":
            html_str = (
                "<ruby>\n"
                "  <rb>%s</rb>\n"   # 漢字
                "  <rp>(</rp>\n"    # 括號左
                "  <rt>%s</rt>\n"   # 上方標音
                "  <rp>)</rp>\n"    # 括號右
                "  <rp>(</rp>\n"    # 括號左
                "  <rtc>%s</rtc>\n" # 右方標音
                "  <rp>)</rp>\n"    # 括號右
                "</ruby>\n"
            )
            return html_str % (han_ji, siong_piau_im, zian_piau_im)
        else:
            return f"<span>{han_ji}</span>\n"

    def generate_title_with_ruby(self, sheet, wb) -> str:
        """
        取得文章標題並加注【台語音標】

        Returns:
            含有 Ruby 標籤的標題 HTML
        """
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
                    tlpa_im_piau_list.append(sheet.range((row-1, col)).value)
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
                han_ji_piau_im = ""
                siong_piau_im = ""
                zian_piau_im = ""

                if han_ji.strip() == "":
                    i += 1
                    continue
                elif han_ji == '\n':
                    # 若讀到換行字元，則直接輸出換行標籤
                    tag = "<br/>\n"
                    title_with_ruby += tag
                elif not is_han_ji(han_ji):
                    tag = f"<span>{han_ji}</span>"
                    title_with_ruby += tag
                else:
                    # 取得對應的台語音標
                    tai_gi_im_piau = tlpa_im_piau_list[i] if i < len(tlpa_im_piau_list) else ""
                    # 將【漢字】及【上方】/【右方】標音合併成一個 Ruby Tag
                    ruby_tag, siong_piau_im, zian_piau_im = self.generate_ruby_tag(
                        wb=wb, han_ji=han_ji, tai_gi_im_piau=tai_gi_im_piau
                    )
                    title_with_ruby += ruby_tag
                    msg = f"{han_ji} [{tai_gi_im_piau}] ==》 上方標音：{siong_piau_im} / 右方標音：{zian_piau_im}"
                    print(msg)
                i += 1

        return title_with_ruby

    def generate_web_page(self, sheet, wb) -> str:
        """
        製作完整網頁內容

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
        if self.config.image_url:
            write_buffer += div_image % (self.config.web_title, self.config.image_url)

        # 輸出【文章】Div tag 及【文章標題】Ruby Tag
        title_with_ruby = self.generate_title_with_ruby(sheet, wb)
        pai_ban_iong_huat = self.config.zu_im_huat_list[self.config.piau_im_format][0]

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
            div_tag = (
                "<div class='%s'>\n"
                "  <p class='title'>\n"
                "    %s\n"
                "  </p>\n"
            )
        write_buffer += div_tag % (pai_ban_iong_huat, title_with_ruby)

        # 逐列處理工作表內容
        End_Of_File = False
        char_count = 0
        total_chars_per_line = int(wb.names['網頁每列字數'].refers_to_range.value)
        if total_chars_per_line == 0:
            total_chars_per_line = 0

        for row in range(self.config.start_row, self.config.end_row, self.config.ROWS_PER_LINE):
            if title_with_ruby and row == self.config.start_row:
                # 已經處理過標題列，跳過
                continue
            sheet.range((row, 1)).select()

            for col in range(self.config.start_col, self.config.end_col):
                cell_value = sheet.range((row, col)).value

                if cell_value == 'φ':  # 讀到【結尾標示】
                    End_Of_File = True
                    msg = "《文章終止》"
                    print(f"{char_count}. {xw.utils.col_name(col)}{row} = ({row}, {col}) ==> {msg}")
                    break

                elif cell_value == '\n':  # 讀到【換行標示】
                    write_buffer += "</p><p>\n"
                    msg = "《換行》"
                    char_count = 0
                    print(f"{char_count}. {xw.utils.col_name(col)}{row} = ({row}, {col}) ==> {msg}")
                    break

                elif not is_han_ji(cell_value):
                    # 處理非漢字內容
                    str_value = str(cell_value).strip() if cell_value else ""

                    if is_punctuation(str_value):
                        msg = f"{str_value}【標點符號】"
                        ruby_tag = f"  <span>{str_value}</span>\n"
                    elif str_value == "":
                        msg = "【空白】"
                        ruby_tag = "  <span>　</span>\n"
                    else:
                        msg = f"{str_value}【其他字元】"
                        ruby_tag = f"  <span>{str_value}</span>\n"

                    write_buffer += ruby_tag
                    char_count += 1
                    print(f"{char_count}. {xw.utils.col_name(col)}{row} = ({row}, {col}) ==> {msg}")

                else:
                    # 處理漢字
                    han_ji = cell_value.strip()
                    tai_gi_im_piau = sheet.range((row + PIAU_IM_ROW, col)).value
                    tai_gi_im_piau = tai_gi_im_piau if tai_gi_im_piau is not None else ""

                    # 【去調符】轉換為【帶調號之台語音標】
                    if kam_si_u_tiau_hu(tai_gi_im_piau):
                        tai_gi_im_piau = tng_im_piau(tai_gi_im_piau)
                        tlpa_im_piau = tng_tiau_ho(tai_gi_im_piau)
                    else:
                        tlpa_im_piau = tai_gi_im_piau

                    ruby_tag, siong_piau_im, zian_piau_im = self.generate_ruby_tag(
                        wb=wb,
                        han_ji=han_ji,
                        tai_gi_im_piau=tlpa_im_piau,
                    )
                    write_buffer += ruby_tag
                    msg = f"{han_ji} [{tlpa_im_piau}] ==》 上方標音：{siong_piau_im} / 右方標音：{zian_piau_im}"
                    char_count += 1
                    print(f"{char_count}. {xw.utils.col_name(col)}{row} = ({row}, {col}) ==> {msg}")

                # 檢查是否需要插入換行標籤
                if total_chars_per_line != 0 and char_count >= total_chars_per_line:
                    write_buffer += "<br/>\n"
                    char_count = 0
                    print('《人工斷行》')

            if End_Of_File:
                break

        write_buffer += "</p></div>"
        return write_buffer


# =========================================================================
# 主要處理函數
# =========================================================================
def generate_web_page(wb, sheet_name: str = '漢字注音') -> int:
    """
    製作標音網頁

    Args:
        wb: Excel Workbook 物件
        sheet_name: 工作表名稱

    Returns:
        處理結果代碼
    """
    try:
        # 初始化配置
        config = WebPageConfig(wb)

        # 初始化標音物件
        piau_im = PiauIm(config.han_ji_khoo_name)

        # 決定使用的資料庫
        db_name = DB_HO_LOK_UE if config.han_ji_khoo_name == '河洛話' else DB_KONG_UN

        # 初始化字典物件
        ji_tian = HanJiTian(db_name)

        # 建立網頁生成器
        generator = WebPageGenerator(config, piau_im, ji_tian)

        # 選擇工作表
        sheet = wb.sheets[sheet_name]
        sheet.activate()
        sheet.range('A1').select()

        # 產生 HTML 網頁內容
        print("開始製作【漢字注音】網頁！")
        html_content = generator.generate_web_page(sheet, wb)

        # 生成輸出檔案名稱
        hue_im = config.ue_im_lui_piat
        piau_im_huat = config.piau_im_huat
        if config.piau_im_format == "無預設":
            im_piau = piau_im_huat
        elif config.piau_im_format == "上":
            im_piau = config.siong_pinn_piau_im
        elif config.piau_im_format == "右":
            im_piau = config.zian_pinn_piau_im
        else:
            im_piau = f"{config.siong_pinn_piau_im}＋{config.zian_pinn_piau_im}"

        output_file = f"《{config.title}》【{hue_im}】{im_piau}.html"
        output_path = os.path.join(config.output_dir, output_file)

        # 確保輸出目錄存在
        os.makedirs(config.output_dir, exist_ok=True)

        # 取得 env 工作表的設定值並組合 meta 標籤字串
        env_keys = [
            "FILE_NAME", "TITLE", "IMAGE_URL", "OUTPUT_PATH", "章節序號",
            "顯示注音輸入", "每頁總列數", "每列總字數", "語音類型",
            "漢字庫", "標音方法", "網頁格式", "標音方式", "上邊標音", "右邊標音", "網頁每列字數"
        ]
        head_extra = ""
        for key in env_keys:
            value = get_value_by_name(wb, key)
            head_extra += f'    <meta name="{key}" content="{value}" />\n'

        # 輸出到網頁檔案
        _create_html_file(output_path, html_content, config.web_title, head_extra)

        logging_process_step("【漢字注音】工作表轉製網頁作業完畢！")
        return EXIT_CODE_SUCCESS

    except Exception as e:
        logging.exception("製作標音網頁作業發生例外！")
        raise


def _create_html_file(output_path: str, content: str, title: str = '您的標題', head_extra: str = ""):
    """創建 HTML 檔案"""
    template = f"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <title>{title}</title>
    <meta charset="UTF-8">
    {head_extra}
    <link rel="stylesheet" href="assets/styles/styles2.css">
</head>
<body>
    {content}
</body>
</html>
    """
    with open(output_path, 'w', encoding='utf-8') as file:
        file.write(template)
    print(f"\n輸出網頁檔案：{output_path}")


# =========================================================================
# 作業程序
# =========================================================================
def process(wb) -> int:
    """執行主要處理工作"""
    logging_process_step("<----------- 作業開始！---------->")

    # 將【漢字注音】工作表中的標音漢字，轉成 HTML 網頁檔案
    status_code = generate_web_page(
        wb=wb,
        sheet_name='漢字注音',
    )

    if status_code != EXIT_CODE_SUCCESS:
        logging.error("標音漢字轉換為 HTML 網頁檔案失敗！")
        return status_code

    # 要求畫面回到【漢字注音】工作表
    wb.sheets['漢字注音'].activate()

    logging_process_step("<----------- 作業結束！---------->")
    return EXIT_CODE_SUCCESS


# =============================================================================
# 程式主流程
# =============================================================================
def main() -> int:
    """主程式入口"""
    # 程式初始化
    current_file_path = Path(__file__).resolve()
    project_root = current_file_path.parent
    program_name = current_file_path.stem

    logging_process_step(f"《========== 程式開始執行：{program_name} ==========》")
    logging_process_step(f"專案根目錄為: {project_root}")

    # 設定【作用中活頁簿】
    wb = None
    try:
        wb = xw.apps.active.books.active
    except Exception as e:
        logging.error(f"無法找到作用中的 Excel 工作簿: {e}", exc_info=True)
        return EXIT_CODE_NO_FILE

    if not wb:
        return EXIT_CODE_NO_FILE

    # 執行【處理作業】
    try:
        result_code = process(wb)
        if result_code != EXIT_CODE_SUCCESS:
            msg = f"程式異常終止：{program_name}"
            logging_exc_error(msg=msg, error=None)
            return EXIT_CODE_PROCESS_FAILURE

    except Exception as e:
        msg = f"程式異常終止：{program_name}"
        logging_exc_error(msg=msg, error=e)
        return EXIT_CODE_UNKNOWN_ERROR

    finally:
        # 儲存檔案
        try:
            wb.sheets['漢字注音'].activate()
            file_path = save_as_new_file(wb=wb)
            if not file_path:
                logging_exc_error(msg="儲存檔案失敗！", error=None)
                return EXIT_CODE_SAVE_FAILURE
            else:
                logging_process_step(f"儲存檔案至路徑：{file_path}")
        except Exception as e:
            logging_exc_error(msg="儲存檔案失敗！", error=e)
            return EXIT_CODE_SAVE_FAILURE

    # 結束程式
    logging_process_step(f"《========== 程式終止執行：{program_name} ==========》")
    return EXIT_CODE_SUCCESS


if __name__ == "__main__":
    import sys

    exit_code = main()
    sys.exit(exit_code)
