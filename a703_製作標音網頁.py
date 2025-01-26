# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import logging
import os
import sys
from pathlib import Path

# 載入第三方套件
import xlwings as xw
from dotenv import load_dotenv

# 載入自訂模組
from mod_file_access import save_as_new_file
from mod_標音 import split_tai_gi_im_piau  # 分解台語音標
from mod_標音 import PiauIm, is_punctuation

# from p730_Tng_Sing_Bang_Iah import tng_sing_bang_iah

# =========================================================================
# 載入環境變數
# =========================================================================
load_dotenv()

# 預設檔案名稱從環境變數讀取
DB_HO_LOK_UE = os.getenv('DB_HO_LOK_UE', 'Ho_Lok_Ue.db')
DB_KONG_UN = os.getenv('DB_KONG_UN', 'Kong_Un.db')

# =========================================================================
# 設定日誌
# =========================================================================
logging.basicConfig(
    filename='process_log.txt',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def logging_process_step(msg):
    print(msg)
    logging.info(msg)

# =========================================================================
# 常數定義
# =========================================================================
# 定義 Exit Code
EXIT_CODE_SUCCESS = 0  # 成功
EXIT_CODE_NO_FILE = 1  # 無法找到檔案
EXIT_CODE_INVALID_INPUT = 2  # 輸入錯誤
EXIT_CODE_PROCESS_FAILURE = 3  # 過程失敗
EXIT_CODE_UNKNOWN_ERROR = 99  # 未知錯誤


# =========================================================================
# 程式區域函式
# =========================================================================
def create_html_file(output_path, content, title='您的標題'):
    template = f"""
<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <title>{title}</title>
    <meta charset="UTF-8">
    <link rel="stylesheet" href="assets/styles/styles.css">
</head>
<body>
    {content}
</body>
</html>
    """

    # Write to HTML file
    with open(output_path, 'w', encoding='utf-8') as file:
        file.write(template)

    # 顯示輸出之網頁檔案及其存放目錄路徑
    print(f"\n輸出網頁檔案：{output_path}")


def put_picture(wb, source_sheet_name):
    html_str = ""

    title = wb.sheets["env"].range("TITLE").value
    # web_page_title = f"《{title}》【{source_sheet_name}】\n"
    web_page_title = f"《{title}》\n"
    image_url = wb.sheets["env"].range("IMAGE_URL").value

    # ruff: noqa: E501
    div_tag = (
        "<div class='separator' style='clear: both'>\n"
        "  <a href='圖片' style='display: block; padding: 1em 0; text-align: center'>\n"
        "    <img alt='%s' border='0' width='400' data-original-height='630' data-original-width='1200'\n"
        "      src='%s' />\n"
        "  </a>\n"
        "</div>\n"
    )
    # 寫入文章附圖
    # html_str = f"《{title}》【{source_sheet_name}】\n"
    html_str = f"{title}\n"
    # html_str += div_tag % (title, image_url)
    html_str += (div_tag % (title, image_url) + "\n")
    return html_str


def tng_uann_piau_im(piau_im, zu_im_huat, siann_bu, un_bu, tiau_ho):
    """根據指定的標音方法，轉換台語音標之羅馬拚音字母"""
    if zu_im_huat == "十五音":
        return piau_im.SNI_piau_im(siann_bu, un_bu, tiau_ho)
    elif zu_im_huat == "雅俗通":
        return piau_im.NST_piau_im(siann_bu, un_bu, tiau_ho)
    elif zu_im_huat == "白話字":
        return piau_im.POJ_piau_im(siann_bu, un_bu, tiau_ho)
    elif zu_im_huat == "台羅拼音":
        return piau_im.TL_piau_im(siann_bu, un_bu, tiau_ho)
    elif zu_im_huat == "閩拼方案":
        return piau_im.BP_piau_im(siann_bu, un_bu, tiau_ho)
    elif zu_im_huat == "方音符號":
        return piau_im.TPS_piau_im(siann_bu, un_bu, tiau_ho)
    elif zu_im_huat == "台語音標":
        siann = piau_im.Siann_Bu_Dict[siann_bu]["台語音標"] or ""
        un = piau_im.Un_Bu_Dict[un_bu]["台語音標"]
        return f"{siann}{un}{tiau_ho}"
    return ""


def concat_ruby_tag(wb, piau_im, han_ji, tai_gi_im_piau):
    """將漢字、台語音標及台語注音符號，合併成一個 Ruby Tag"""
    zu_im_list = split_tai_gi_im_piau(tai_gi_im_piau)
    if zu_im_list[0] == "" or zu_im_list[0] == None:
        siann_bu = "ø"
    else:
        siann_bu = zu_im_list[0]

    style = wb.names['網頁格式'].refers_to_range.value
    piau_im_hong_sik = wb.names['標音方式'].refers_to_range.value
    siong_pinn_piau_im = wb.names['上邊標音'].refers_to_range.value
    zian_pinn_piau_im = wb.names['右邊標音'].refers_to_range.value

    ruby_tag = ""
    siong_piau_im = ""
    zian_piau_im = ""

    # 根據【網頁格式】，決定【漢字】之上方或右方，是否該顯示【標音】
    if style == "無預設":
        # 若【網頁格式】設定為【無預設】，則根據【標音方式】決定漢字之上方及右方，是否需要放置標音
        if piau_im_hong_sik == "上及右":
            # 漢字上方顯示【上邊標音】，下方顯示【下邊標音】
            siong_piau_im = tng_uann_piau_im(
                piau_im=piau_im,    # 注音法物件
                zu_im_huat=siong_pinn_piau_im,
                siann_bu=siann_bu,
                un_bu=zu_im_list[1],
                tiau_ho=zu_im_list[2]
            )
            zian_piau_im = tng_uann_piau_im(
                piau_im=piau_im,    # 注音法物件
                zu_im_huat=zian_pinn_piau_im,
                siann_bu=siann_bu,
                un_bu=zu_im_list[1],
                tiau_ho=zu_im_list[2]
            )
        elif piau_im_hong_sik == "上":
            siong_piau_im = tng_uann_piau_im(
                piau_im=piau_im,    # 注音法物件
                zu_im_huat=siong_pinn_piau_im,
                siann_bu=siann_bu,
                un_bu=zu_im_list[1],
                tiau_ho=zu_im_list[2]
            )
        elif piau_im_hong_sik == "右":
            zian_piau_im = tng_uann_piau_im(
                piau_im=piau_im,    # 注音法物件
                zu_im_huat=zian_pinn_piau_im,
                siann_bu=siann_bu,
                un_bu=zu_im_list[1],
                tiau_ho=zu_im_list[2]
            )
    else:
        if style == "POJ" or style == "TL" or style == "BP" or style == "TLPA_Plus":
            # 羅馬拼音字母標音法，將標音置於漢字上方
            siong_piau_im = tng_uann_piau_im(
                piau_im=piau_im,    # 注音法物件
                zu_im_huat=siong_pinn_piau_im,
                siann_bu=siann_bu,
                un_bu=zu_im_list[1],
                tiau_ho=zu_im_list[2]
            )
        elif style == "SNI":
            # 十五音反切法，將標音置於漢字上方
            siong_piau_im = tng_uann_piau_im(
                piau_im=piau_im,    # 注音法物件
                zu_im_huat=siong_pinn_piau_im,
                siann_bu=siann_bu,
                un_bu=zu_im_list[1],
                tiau_ho=zu_im_list[2]
            )
        elif style == "TPS":
            # 注音符號標音法，將標音置於漢字右方
            zian_piau_im = tng_uann_piau_im(
                piau_im=piau_im,    # 注音法物件
                zu_im_huat=zian_pinn_piau_im,
                siann_bu=siann_bu,
                un_bu=zu_im_list[1],
                tiau_ho=zu_im_list[2]
            )
        elif style == "DBL":
            # 漢字上方顯示台語音標，下方顯示台語注音符號
            siong_piau_im = tng_uann_piau_im(
                piau_im=piau_im,    # 注音法物件
                zu_im_huat=siong_pinn_piau_im,
                siann_bu=siann_bu,
                un_bu=zu_im_list[1],
                tiau_ho=zu_im_list[2]
            )
            zian_piau_im = tng_uann_piau_im(
                piau_im=piau_im,    # 注音法物件
                zu_im_huat=zian_pinn_piau_im,
                siann_bu=siann_bu,
                un_bu=zu_im_list[1],
                tiau_ho=zu_im_list[2]
            )

    # 根據標音方式，設定 Ruby Tag
    if siong_piau_im != "" and zian_piau_im == "":
        # 將標音置於漢字上方
        ruby_tag = f"  <ruby><rb>{han_ji}</rb><rp>(</rp><rt>{siong_piau_im}</rt><rp>)</rp></ruby>\n"
    elif siong_piau_im == "" and zian_piau_im != "":
        # 將標音置於漢字右方
        ruby_tag = f"  <ruby><rb>{han_ji}</rb><rp>(</rp><rtc>{zian_piau_im}</rtc><rp>)</rp></ruby>\n"
    elif siong_piau_im != "" and zian_piau_im != "":
        # 將標音置於漢字上方及右方
        ruby_tag = f"  <ruby><rb>{han_ji}</rb><rt>{siong_piau_im}</rt><rp>(</rp><rtc>{zian_piau_im}</rtc><rp>)</rp></ruby>\n"

    return ruby_tag


# =========================================================
# 依據指定的【注音方法】，輸出含 Ruby Tags 之 HTML 網頁
# =========================================================
def build_web_page(wb, sheet, source_chars, total_length, page_type='含頁頭', piau_im_huat='方音符號', piau_im=None):
    # ==========================================================
    # 注音法設定和共用變數
    # ==========================================================
    zu_im_huat_list = {
        "SNI": ["fifteen_yin", "rt", "十五音切語"],
        "TPS": ["Piau_Im", "rt", "方音符號注音"],
        "POJ": ["pin_yin", "rt", "白話字拼音"],
        "TL": ["pin_yin", "rt", "台羅拼音"],
        "BP": ["pin_yin", "rt", "閩拼標音"],
        "TLPA_Plus": ["pin_yin", "rt", "台羅改良式"],
        "DBL": ["Siang_Pai", "rtc", "雙排注音"],
        "無預設": ["Siang_Pai", "rtc", "雙排注音"],
    }

    # 選擇工作表
    sheet = wb.sheets['漢字注音']
    sheet.activate()
    write_buffer = ""

    #--------------------------------------------------------------------------
    # 輸出放置圖片的 HTML Tag
    #--------------------------------------------------------------------------
    # 寫入文章附圖
    if page_type == '含頁頭':
        write_buffer += put_picture(wb, sheet.name)

    #--------------------------------------------------------------------------
    # 輸出 <div> tag
    #--------------------------------------------------------------------------
    div_class = zu_im_huat_list[Web_Page_Style][0]
    # html_str = f"<div class='{div_class}'><p>\n"
    # write_buffer += html_str
    write_buffer += f"<div class='{div_class}'><p>\n"

    #--------------------------------------------------------------------------
    # 作業處理：逐列取出漢字，組合成純文字檔
    #--------------------------------------------------------------------------
    # 設定起始及結束的【列】位址（【第5列】、【第9列】、【第13列】等列）
    TOTAL_LINES = int(wb.names['每頁總列數'].refers_to_range.value)
    ROWS_PER_LINE = 4
    start_row = 5
    end_row = start_row + (TOTAL_LINES * ROWS_PER_LINE)
    line = 1    # 處理行號指示器

    # 設定起始及結束的【欄】位址（【D欄=4】到【R欄=18】）
    CHARS_PER_ROW = int(wb.names['每列總字數'].refers_to_range.value)
    start_col = 4
    end_col = start_col + CHARS_PER_ROW

    # 逐列處理作業
    end_of_file = False
    for row in range(start_row, end_row, ROWS_PER_LINE):
        Empty_Cells_Total = 0
        # 設定【作用儲存格】為列首
        sheet.range((row, 1)).select()

        # 逐欄取出儲存格內容
        for col in range(start_col, end_col):
            ruby_tag = ""

            cell_value = sheet.range((row, col)).value
            if cell_value == 'φ':       # 讀到【結尾標示】
                end_of_file = True
                print(f"({row}, {xw.utils.col_name(col)}) = 《文章終止》")
                break
            elif cell_value == '\n':    # 讀到【換行標示】
                print(f"({row}, {xw.utils.col_name(col)}) = 《換行》")
                # 若遇到換行字元，退出迴圈
                break
            elif cell_value == None or cell_value == '':    # 讀到【空白】
                print(f"({row}, {xw.utils.col_name(col)}) = 《空格》")
                Empty_Cells_Total += 1
                if Empty_Cells_Total >= 2:
                    EOF = True
                    break
                else:
                    continue
            else:                       # 讀到：漢字或標點符號
                # 當 han_ji 是標點符號時，不需要注音
                if is_punctuation(cell_value):
                    # ruby_tag = f"  {cell_value}\n"
                    ruby_tag = f"  <span>{cell_value}</span>\n"
                    msg = f"({row}, {col_name}) = {cell_value}"
                else:
                    han_ji = cell_value  # 取得漢字
                    # 取得漢字的【台語音標】
                    tai_gi_im_piau = sheet.range((row - 1, col)).value  # 取得漢字的台語音標
                    # 當儲存格寫入之資料為 None 情況時之處理作法：給予空字串
                    tai_gi_im_piau = tai_gi_im_piau if tai_gi_im_piau is not None else ""
                    # 將已注音之漢字加入【漢字注音表】
                    ruby_tag = concat_ruby_tag(
                        wb=wb,
                        piau_im=piau_im,    # 注音法物件
                        han_ji=han_ji,
                        tai_gi_im_piau=tai_gi_im_piau
                    )
                    col_name = xw.utils.col_name(col)   # 取得欄位名稱
                    msg =f"({row}, {col_name}) = {han_ji} [{tai_gi_im_piau}]"

            write_buffer += ruby_tag
            print(msg)

        # =========================================================
        # 換行處理：(1)每處理完 15 字後，換下一行 ；(2) 讀到【換行標示】
        # =========================================================

        # 已到【行尾】作業處理
        # print(f"({row}, {col_name}) = 《行尾》")

        # 讀到【換行標示】，需要結束目前【段落】，並開始新的【段落】
        if cell_value == '\n':
            write_buffer += f"</p><p>\n"

        line += 1
        if end_of_file or line > TOTAL_LINES: break

    # 返回網頁輸出暫存區
    write_buffer += "</p></div>"
    return write_buffer


def tng_sing_bang_iah(wb, sheet_name='漢字注音', han_ji_source='V3', page_type='含頁頭'):
    global source_sheet  # 宣告 source_sheet 為全域變數
    global source_sheet_name  # 宣告 source_sheet_name 為全域變數
    global total_length  # 宣告 total_length 為全域變數
    global Web_Page_Style

    # -------------------------------------------------------------------------
    # 連接指定資料庫
    # -------------------------------------------------------------------------
    han_ji_khoo = wb.names['漢字庫'].refers_to_range.value
    Web_Page_Style = wb.names['網頁格式'].refers_to_range.value
    piau_im = PiauIm(han_ji_khoo)

    # -------------------------------------------------------------------------
    # 選擇指定的工作表
    # -------------------------------------------------------------------------
    sheet = wb.sheets[sheet_name]   # 選擇工作表
    sheet.activate()               # 將「漢字注音」工作表設為作用中工作表
    sheet.range('A1').select()     # 將 A1 儲存格設為作用儲存格
    source_sheet_name = sheet.name

    han_ji_piau_im_huat = wb.names['標音方法'].refers_to_range.value

    # -----------------------------------------------------
    # 產生 HTML 網頁用文字檔
    # -----------------------------------------------------
    title = wb.names['TITLE'].refers_to_range.value
    web_page_title = f"{title}"

    # 確保 output 子目錄存在
    output_dir = 'docs'
    output_file = f"{title}_{han_ji_piau_im_huat}.html"
    output_path = os.path.join(output_dir, output_file)

    # 開啟文字檔，準備寫入網頁內容
    f = open(output_path, 'w', encoding='utf-8')

    # 取得 V3 儲存格的字串
    source_chars = sheet.range(han_ji_source).value
    if source_chars:
        # 計算字串的總長度
        total_length = len(source_chars)

        # ==========================================================
        # 自「漢字注音表」，製作各種注音法之 HTML 網頁
        # ==========================================================
        print(f"開始製作【漢字注音】網頁！")
        html_content = build_web_page(
            wb=wb,
            sheet=sheet,
            source_chars=source_chars,
            total_length=total_length,
            page_type=page_type,
            piau_im_huat=han_ji_piau_im_huat,
            piau_im= piau_im
        )

        # 輸出到網頁檔案
        create_html_file(output_path, html_content, web_page_title)
        print(f"【漢字注音】網頁製作完畢！")

    return 0


# =========================================================================
# 作業程序
# =========================================================================
def process(wb):
    # ---------------------------------------------------------------------
    # 將【漢字注音】工作表中的標音漢字，轉成 HTML 網頁檔案。
    # ---------------------------------------------------------------------
    result_code = tng_sing_bang_iah(
        wb=wb,
        sheet_name='漢字注音',
        han_ji_source='V3',
        page_type='含頁頭'
    )
    if result_code != EXIT_CODE_SUCCESS:
        logging.error("標音漢字轉換為 HTML 網頁檔案失敗！")
        return result_code

    # ---------------------------------------------------------------------
    # 作業結尾處理
    # ---------------------------------------------------------------------
    file_path = save_as_new_file(wb=wb)
    if not file_path:
        logging.error("儲存檔案失敗！")
        return EXIT_CODE_PROCESS_FAILURE    # 作業異當終止：無法儲存檔案
    else:
        logging_process_step(f"儲存檔案至路徑：{file_path}")
        return EXIT_CODE_SUCCESS    # 作業正常結束


# =============================================================================
# 程式主流程
# =============================================================================
def main():
    # =========================================================================
    # 開始作業
    # =========================================================================
    logging.info("作業開始")

    # =========================================================================
    # (1) 取得專案根目錄
    # =========================================================================
    current_file_path = Path(__file__).resolve()
    project_root = current_file_path.parent
    print(f"專案根目錄為: {project_root}")
    logging.info(f"專案根目錄為: {project_root}")

    # =========================================================================
    # (2) 嘗試獲取當前作用中的 Excel 工作簿
    # =========================================================================
    wb = None
    try:
        wb = xw.apps.active.books.active
    except Exception as e:
        logging.error(f"無法找到作用中的 Excel 工作簿: {e}", exc_info=True)
        print("無法找到作用中的 Excel 工作簿")
        return EXIT_CODE_NO_FILE

    if not wb:
        print("無法作業，原因可能為：(1) 未指定輸入檔案；(2) 未找到作用中的 Excel 工作簿！")
        logging.error("無法作業，未指定輸入檔案或 Excel 無效。")
        return EXIT_CODE_NO_FILE

    # =========================================================================
    # (3) 處理作業
    # =========================================================================
    try:
        result_code = process(wb)
        if result_code != EXIT_CODE_SUCCESS:
            logging.error("處理作業失敗，過程中出錯！")
            return result_code

    except Exception as e:
        print(f"執行過程中發生未知錯誤: {e}")
        logging.error(f"執行過程中發生未知錯誤: {e}", exc_info=True)
        return EXIT_CODE_UNKNOWN_ERROR

    finally:
        if wb:
            logging_process_step(f"製作【漢字標音】網頁己完成！")

    # =========================================================================
    # 結束作業
    # =========================================================================
    logging.info("作業完成！")
    return EXIT_CODE_SUCCESS


if __name__ == "__main__":
    exit_code = main()
    if exit_code == EXIT_CODE_SUCCESS:
        print("程式正常完成！")
    else:
        print(f"程式異常終止，錯誤代碼為: {exit_code}")
    sys.exit(exit_code)
