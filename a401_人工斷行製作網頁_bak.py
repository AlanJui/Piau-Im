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
    <link rel="stylesheet" href="assets/styles/styles2.css">
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
    elif zu_im_huat == "閩拼調號":
        return piau_im.BP_piau_im(siann_bu, un_bu, tiau_ho)
    elif zu_im_huat == "閩拼調符":
        return piau_im.BP_piau_im_with_tiau_hu(siann_bu, un_bu, tiau_ho)
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
        siann_bu = "ø"  # 無聲母: ø
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
    """
    依據指定的【注音方法】，輸出含 Ruby Tags 之 HTML 網頁，並根據【網頁每列字數】來決定是否手動插入換行 <br> 標籤。
    同時，保留 Console 顯示目前處理狀態，以便 Debug。
    """
    # 取得「網頁每列字數」的設定值
    total_chars_per_line = wb.names['網頁每列字數'].refers_to_range.value
    if total_chars_per_line == "預設":
        total_chars_per_line = None  # 不做人工斷行
    else:
        total_chars_per_line = int(total_chars_per_line)  # 確保為整數

    # 取得「標音方法」
    han_ji_piau_im_huat = wb.names['標音方法'].refers_to_range.value

    # 取得「漢字庫」
    han_ji_khoo = wb.names['漢字庫'].refers_to_range.value
    piau_im = PiauIm(han_ji_khoo)

    # 取得輸出格式
    web_page_style = wb.names['網頁格式'].refers_to_range.value

    # 初始化 HTML 內容
    write_buffer = ""

    # 加入標題圖片
    if page_type == '含頁頭':
        write_buffer += put_picture(wb, sheet.name)

    # 加入 <div> 容器
    write_buffer += "<div class='Siang_Pai'><p>\n"

    # 記錄當前行的字數
    current_line_char_count = 0

    # 逐字處理
    EndOfFile = False
    for row in range(5, sheet.used_range.last_cell.row + 1, 4):  # 逐段處理，每段 4 行
        for col in range(4, sheet.used_range.last_cell.column + 1):
            cell_value = sheet.range((row, col)).value
            if cell_value is None or cell_value.strip() == '':
                continue
            elif cell_value == "φ":
                EndOfFile = True
                print('讀到文章終止符號 φ')
                break

            ruby_tag = ""

            if is_punctuation(cell_value):  # 標點符號
                ruby_tag = f"  <span>{cell_value}</span>\n"
                console_msg = f"({row}, {xw.utils.col_name(col)}) = {cell_value}"
            else:  # 漢字
                tai_gi_im_piau = sheet.range((row - 1, col)).value or ""
                ruby_tag = concat_ruby_tag(
                    wb=wb,
                    piau_im=piau_im,
                    han_ji=cell_value,
                    tai_gi_im_piau=tai_gi_im_piau
                )
                console_msg = f"({row}, {xw.utils.col_name(col)}) = {cell_value} [{tai_gi_im_piau}]"

            # 顯示目前處理進度
            print(console_msg)

            # 計算當前行字數（漢字 + 標點符號）
            current_line_char_count += 1

            # 若有設定【網頁每列字數】，且已達設定字數，則手動換行
            if total_chars_per_line and current_line_char_count >= total_chars_per_line:
                write_buffer += ruby_tag
                write_buffer += "  </br>\n"  # 插入換行
                print(">>> 插入換行 <br>")  # Console 顯示換行點
                current_line_char_count = 0  # 重設行字數
            else:
                write_buffer += ruby_tag

        write_buffer += "</p><p>\n"  # 段落結束，開始新段落
        if EndOfFile:
            break

    # 關閉 HTML 結構
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
