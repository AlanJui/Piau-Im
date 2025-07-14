# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import logging
import os
import sqlite3
import sys
from pathlib import Path

# 載入第三方套件
import xlwings as xw
from dotenv import load_dotenv

# 載入自訂模組
from mod_excel_access import get_value_by_name
from mod_file_access import load_module_function, save_as_new_file
from mod_帶調符音標 import is_han_ji, kam_si_u_tiau_hu, tng_im_piau, tng_tiau_ho

# from mod_廣韻 import tng_uann_piau_im
from mod_標音 import split_tai_gi_im_piau  # 分解台語音標
from mod_標音 import PiauIm, ca_ji_kiat_ko_tng_piau_im, is_punctuation

# =========================================================================
# 常數定義
# =========================================================================
Piau_Im_Row = -1  # 標音位置：-1 ==> 自動標音；-2 ==> 人工標音
# 定義 Exit Code
EXIT_CODE_SUCCESS = 0  # 成功
EXIT_CODE_NO_FILE = 1  # 無法找到檔案
EXIT_CODE_INVALID_INPUT = 2  # 輸入錯誤
EXIT_CODE_SAVE_FAILURE = 3  # 儲存失敗
EXIT_CODE_PROCESS_FAILURE = 10  # 過程失敗
EXIT_CODE_UNKNOWN_ERROR = 99  # 未知錯誤

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
from mod_logging import init_logging, logging_exc_error, logging_process_step

init_logging()

# =========================================================================
# 程式區域函式
# =========================================================================
def create_html_file(output_path, content, title='您的標題', head_extra=""):
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


def title_piau_im(wb, title: str) -> str:
    # ------------------------------------------------------------------------------
    # 為漢字查找讀音，漢字上方填：【台語音標】；漢字下方填使用者指定之【漢字標音】
    # ------------------------------------------------------------------------------
    # 取得【漢字庫】之資料庫查詢所需環境參數
    try:
        han_ji_khoo_name = wb.names['漢字庫'].refers_to_range.value
        ue_im_lui_piat = wb.names['語音類型'].refers_to_range.value  # 取得【語音類型】，判別使用【白話音】或【文讀音】何者。
        db_name = 'Ho_Lok_Ue.db' if han_ji_khoo_name == '河洛話' else 'Kong_Un.db'
        piau_im_huat = wb.names['標音方法'].refers_to_range.value
    except Exception as e:
        logging_exc_error(f"【env】工作表找不到選項之設定值", e)
        raise

    # 決定使用的模組名稱
    if han_ji_khoo_name == '河洛話':
        module_name = 'mod_河洛話'
    else:
        module_name = 'mod_廣韻'
    # 決定使用的函數名稱
    function_name = 'han_ji_ca_piau_im'

    # 載入【漢字庫】查找函數
    if han_ji_khoo_name == "河洛話":
        han_ji_ca_piau_im = load_module_function(module_name, function_name)
    elif han_ji_khoo_name == "廣韻":
        han_ji_ca_piau_im = load_module_function(module_name, function_name)
    else:
        msg = "無法執行漢字標音作業，請確認【env】工作表【語音類型】及【漢字庫】欄位的設定是否正確！"
        logging_exc_error(msg=msg, error=None)
        raise ValueError(msg)

    # 連接指定資料庫
    conn = sqlite3.connect(db_name)
    cursor = conn.cursor()

    # 産生標音物件
    piau_im = PiauIm(han_ji_khoo_name)

    #---------------------------------------------------------------
    # 將標題的文字加注【台語音標】
    #---------------------------------------------------------------
    u_piau_im_title = ""
    for han_ji in title:
        tai_gi_im_piau = ""
        han_ji_piau_im = ""

        if han_ji.strip() == "":
            continue
        elif han_ji == '\n':
            # 若讀到換行字元，則直接輸出換行標籤
            tag = "<br/>\n"
            u_piau_im_title += tag
        elif not is_han_ji(han_ji):
            tag = f"<span>{han_ji}</span>"
            u_piau_im_title += tag
        else:
            # 自【漢字庫】查找作業
            result = han_ji_ca_piau_im(cursor=cursor,
                                        han_ji=han_ji,
                                        ue_im_lui_piat=ue_im_lui_piat)
            # 若【漢字庫】查無此字，登錄至【缺字表】
            if not result:
                msg = f"【{han_ji}】查無此字！"
                logging_exc_error(msg=msg, error=None)
            else:
                # 依【漢字庫】查找結果，輸出【台語音標】和【漢字標音】
                tai_gi_im_piau, han_ji_piau_im = ca_ji_kiat_ko_tng_piau_im(
                    result=result,
                    han_ji_khoo=han_ji_khoo_name,
                    piau_im=piau_im,
                    piau_im_huat=piau_im_huat
                )
                # 將【漢字】及【上方】/【右方】標音合併成一個 Ruby Tag
                ruby_tag = concat_ruby_tag(
                    wb=wb,
                    piau_im=piau_im,    # 注音法物件
                    han_ji=han_ji,
                    tai_gi_im_piau=tai_gi_im_piau
                )
                u_piau_im_title += ruby_tag
        print(f"{han_ji} = [{tai_gi_im_piau}] / [{han_ji_piau_im}]")

    return u_piau_im_title


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
        "    <img alt='%s' border='0' width='800' \n"
        "      src='%s' />\n"
        "  </a>\n"
        "</div>\n"
    )
    # 寫入文章附圖
    html_str += (div_tag % (title, image_url) + "\n")
    return html_str


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
            siong_piau_im = piau_im.han_ji_piau_im_tng_huan(
                piau_im_huat=siong_pinn_piau_im,
                siann_bu=siann_bu,
                un_bu=zu_im_list[1],
                tiau_ho=zu_im_list[2]
            )
            zian_piau_im = piau_im.han_ji_piau_im_tng_huan(
                piau_im_huat=zian_pinn_piau_im,
                siann_bu=siann_bu,
                un_bu=zu_im_list[1],
                tiau_ho=zu_im_list[2]
            )
        elif piau_im_hong_sik == "上":
            siong_piau_im = piau_im.han_ji_piau_im_tng_huan(
                piau_im_huat=siong_pinn_piau_im,
                siann_bu=siann_bu,
                un_bu=zu_im_list[1],
                tiau_ho=zu_im_list[2]
            )
        elif piau_im_hong_sik == "右":
            zian_piau_im = piau_im.han_ji_piau_im_tng_huan(
                piau_im_huat=zian_pinn_piau_im,
                siann_bu=siann_bu,
                un_bu=zu_im_list[1],
                tiau_ho=zu_im_list[2]
            )
    else:
        if style == "POJ" or style == "TL" or style == "BP" or style == "TLPA_Plus":
            # 羅馬拼音字母標音法，將標音置於漢字上方
            siong_piau_im = piau_im.han_ji_piau_im_tng_huan(
                piau_im_huat=siong_pinn_piau_im,
                siann_bu=siann_bu,
                un_bu=zu_im_list[1],
                tiau_ho=zu_im_list[2]
            )
        elif style == "SNI":
            # 十五音反切法，將標音置於漢字上方
            siong_piau_im = piau_im.han_ji_piau_im_tng_huan(
                piau_im_huat=siong_pinn_piau_im,
                siann_bu=siann_bu,
                un_bu=zu_im_list[1],
                tiau_ho=zu_im_list[2]
            )
        elif style == "TPS":
            # 注音符號標音法，將標音置於漢字右方
            zian_piau_im = piau_im.han_ji_piau_im_tng_huan(
                piau_im_huat=zian_pinn_piau_im,
                siann_bu=siann_bu,
                un_bu=zu_im_list[1],
                tiau_ho=zu_im_list[2]
            )
        elif style == "DBL":
            # 漢字上方顯示台語音標，下方顯示台語注音符號
            siong_piau_im = piau_im.han_ji_piau_im_tng_huan(
                piau_im_huat=siong_pinn_piau_im,
                siann_bu=siann_bu,
                un_bu=zu_im_list[1],
                tiau_ho=zu_im_list[2]
            )
            zian_piau_im = piau_im.han_ji_piau_im_tng_huan(
                piau_im_huat=zian_pinn_piau_im,
                siann_bu=siann_bu,
                un_bu=zu_im_list[1],
                tiau_ho=zu_im_list[2]
            )

    # 根據標音方式，設定 Ruby Tag
    if siong_piau_im != "" and zian_piau_im == "":
        # 將標音置於漢字上方
        # ruby_tag = f"  <ruby><rb>{han_ji}</rb><rp>(</rp><rt>{siong_piau_im}</rt><rp>)</rp></ruby>\n"
        html_str = (
            "<ruby>\n"
            "  <rb>%s</rb>\n"   # 漢字
            "  <rp>(</rp>\n"    # 括號左
            "  <rt>%s</rt>\n"   # 上方標音
            "  <rp>)</rp>\n"    # 括號右
            "</ruby>\n"
        )
        ruby_tag = html_str % (han_ji, siong_piau_im)
    elif siong_piau_im == "" and zian_piau_im != "":
        # 將標音置於漢字右方
        # ruby_tag = f"  <ruby><rb>{han_ji}</rb><rp>(</rp><rtc>{zian_piau_im}</rtc><rp>)</rp></ruby>\n"
        html_str = (
            "<ruby>\n"
            "  <rb>%s</rb>\n"   # 漢字
            "  <rp>(</rp>\n"    # 括號左
            "  <rtc>%s</rtc>\n" # 右方標音
            "  <rp>)</rp>\n"    # 括號右
            "</ruby>\n"
        )
        ruby_tag = html_str % (han_ji, zian_piau_im)
    elif siong_piau_im != "" and zian_piau_im != "":
        # 將標音置於漢字上方及右方
        # ruby_tag = f"  <ruby><rb>{han_ji}</rb><rt>{siong_piau_im}</rt><rp>(</rp><rtc>{zian_piau_im}</rtc><rp>)</rp></ruby>\n"
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
        ruby_tag = html_str % (han_ji, siong_piau_im, zian_piau_im)

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
        html_str = put_picture(wb, sheet.name)
        write_buffer += html_str

    #--------------------------------------------------------------------------
    # 輸出【文章】Div tag 及【文章標題】Ruby Tag
    #--------------------------------------------------------------------------
    # # 取得文章標題並加注【台語音標】
    # title = wb.names['TITLE'].refers_to_range.value
    # title_with_ruby = title_piau_im(wb, title)

    #--------------------------------------------------------------------
    # 取得文章標題：自 (5,4) 開始讀取到遇到 "》" 為止
    #--------------------------------------------------------------------
    title_chars = ""
    title_with_ruby = ""
    title_cells = []  # ⬅️ 新增：用來記錄讀過的儲存格座標
    row, col = 5, 4

    if sheet.range((row, col)).value == "《":
        while True:
            cell_val = sheet.range((row, col)).value
            if cell_val is None:
                break
            title_chars += cell_val
            title_cells.append((row, col))  # ⬅️ 記下每一格位置
            if cell_val == "》":
                break
            col += 1
            if col > 18:  # 超出一列最大值（R欄=18），則換行至下一列
                row += 1
                col = 4

    # 設定文章內容使用之 div tag 及標題 Ruby Tag
    piau_im_format = wb.names['網頁格式'].refers_to_range.value
    pai_ban_iong_huat = zu_im_huat_list[piau_im_format][0]
    if title_chars:
        # 去除《與》符號，只傳入標題文字本體加注 ruby
        title_han_ji = title_chars.replace("《", "").replace("》", "")
        title_with_ruby = title_piau_im(wb, title_han_ji)

        # CSS 排版用法（CSS class 名稱）
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
    html_str = div_tag % (pai_ban_iong_huat, title_with_ruby)
    write_buffer += html_str

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

    # 取得【網頁每列字數】設定值：數值 0 表【預設】
    total_chars_per_line = int(wb.names['網頁每列字數'].refers_to_range.value)
    if total_chars_per_line == 0:
        total_chars_per_line = 0

    # 逐列處理作業
    End_Of_File = False
    char_count = 0  # 用於計算每列的字數
    for row in range(start_row, end_row, ROWS_PER_LINE):
        # 設定【作用儲存格】為列首
        sheet.range((row, 1)).select()

        # 逐欄取出儲存格內容
        for col in range(start_col, end_col):
            if (row, col) in title_cells:
                continue  # ⬅️ 跳過已處理過的標題儲存格

            ruby_tag = ""
            cell_value = sheet.range((row, col)).value
            # 若【儲存格】內存放【整數值】，則轉為【字串】
            if cell_value == 'φ':       # 讀到【結尾標示】
                End_Of_File = True
                msg = f"《文章終止》"
                print(f"{char_count}. {xw.utils.col_name(col)}{row} = ({row}, {col}) ==> {msg}")
                break
            elif cell_value == '\n':    # 讀到【換行標示】
                # 若遇到換行字元，退出迴圈
                ruby_tag = f"</p><p>\n"
                msg = f"《換行》"
                char_count += 1
                print(f"{char_count}. {xw.utils.col_name(col)}{row} = ({row}, {col}) ==> {msg}")
                char_count = 0  # 重置字數計數器
                break
            elif not is_han_ji(cell_value):
                # # 若【儲存格】存放非漢字，則為：【標點符號】、【空白】或【數值】等
                # if isinstance(cell_value, float) and cell_value.is_integer():
                #     cell_value = str(int(cell_value))
                # ruby_tag = f"  <span>{cell_value}</span>\n"
                # msg = f"{cell_value}"
                # char_count += 1
                str_value = str(cell_value).strip()
                # ✅ 若為全形／半形標點符號
                if is_punctuation(str_value):
                    msg = f"{str_value}【標點符號】"
                    ruby_tag = f"  <span>{str_value}</span>\n"
                elif isinstance(str_value, float) and cell_value.is_integer():
                    # str_value = str(int(str_value))
                    msg = f"{str_value}【英/數半形字元】"
                    ruby_tag = f"  <span>{str_value}</span>\n"
                elif str_value == None or str_value == "":  # 若儲存格內無值
                    msg = "【空白】"    # 表【儲存格】未填入任何字/符，不同於【空白】字元
                    # ruby_tag = f"  <span>&nbsp;&nbsp;</span>\n"
                    ruby_tag = f"  <span>　</span>\n"
                char_count += 1
            else:
                # 當【儲存格】存放的是【漢字】時，則需標注漢字標音
                han_ji = cell_value.strip()  # 消去空白字元

                # 取得漢字的【台語音標】
                tai_gi_im_piau = sheet.range((row + Piau_Im_Row, col)).value  # 取得漢字的台語音標
                # 當儲存格寫入之資料為 None 情況時之處理作法：給予空字串
                tai_gi_im_piau = tai_gi_im_piau if tai_gi_im_piau is not None else ""

                # 如果輸入之【音標】為【帶調符音標】，則需確保轉換為【帶調號TLPA音標】
                if kam_si_u_tiau_hu(tai_gi_im_piau):
                    tai_gi_im_piau = tng_im_piau(tai_gi_im_piau)
                    tlpa_im_piau = tng_tiau_ho(tai_gi_im_piau)
                else:
                    tlpa_im_piau = tai_gi_im_piau

                # 將已注音之漢字加入【漢字注音表】
                ruby_tag = concat_ruby_tag(
                    wb=wb,
                    piau_im=piau_im,    # 注音法物件
                    han_ji=han_ji,
                    tai_gi_im_piau=tlpa_im_piau
                )
                # msg =f"({row}, {xw.utils.col_name(col)}) = {han_ji} [{tlpa_im_piau}] <-- {tai_gi_im_piau}"
                msg =f"{han_ji} [{tlpa_im_piau}] /【{tai_gi_im_piau}】"
                char_count += 1

            write_buffer += ruby_tag
            print(f"{char_count}. {xw.utils.col_name(col)}{row} = ({row}, {col}) ==> {msg}")

            # 檢查是否需要插入換行標籤
            if total_chars_per_line != 0 and char_count >= total_chars_per_line:
                write_buffer += "<br/>\n"
                char_count = 0  # 重置字數計數器
                print('《人工斷行》')

        # =========================================================
        # 換行處理：(1)每處理完 15 字後，換下一行 ；(2) 讀到【換行標示】
        # =========================================================
        # 讀到【換行標示】，需要結束目前【段落】，並開始新的【段落】
        if cell_value == '\n':
            write_buffer += f"</p><p>\n"
            char_count = 0  # 重置字數計數器

        line += 1
        if End_Of_File or line > TOTAL_LINES: break

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
    # output_file = f"{title}_{han_ji_piau_im_huat}.html"
    hue_im = wb.names['語音類型'].refers_to_range.value
    piau_im_huat = wb.names['標音方法'].refers_to_range.value
    piau_im_format = wb.names['標音方式'].refers_to_range.value
    if piau_im_format == "無預設":
        im_piau = piau_im_huat
    elif piau_im_format == "上":
        im_piau = wb.names['上邊標音'].refers_to_range.value
    elif piau_im_format == "右":
        im_piau = wb.names['右邊標音'].refers_to_range.value
    else:
        im_piau = f"{wb.names['上邊標音'].refers_to_range.value}＋{wb.names['右邊標音'].refers_to_range.value}"
    # 檢查檔案名稱是否已包含副檔名
    output_file = f"《{title}》【{hue_im}】{im_piau}.html"

    output_dir = 'docs'
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

        # 取得 env 工作表的設定值並組合 meta 標籤字串
        env_keys = ["FILE_NAME", "TITLE", "IMAGE_URL", "OUTPUT_PATH", "章節序號",
                    "顯示注音輸入", "每頁總列數", "每列總字數", "語音類型",
                    "漢字庫", "標音方法", "網頁格式", "標音方式", "上邊標音", "右邊標音", "網頁每列字數"]
        head_extra = ""
        for key in env_keys:
            value = get_value_by_name(wb, key)
            head_extra += f'    <meta name="{key}" content="{value}" />\n'

        # 輸出到網頁檔案：建立 HTML 檔案時傳入 head_extra
        create_html_file(output_path, html_content, web_page_title, head_extra)
        logging_process_step(f"【漢字注音】工作表轉製網頁作業完畢！")

    return EXIT_CODE_SUCCESS


# =========================================================================
# 作業程序
# =========================================================================
def process(wb):
    logging_process_step("<----------- 作業開始！---------->")
    # ---------------------------------------------------------------------
    # 將【漢字注音】工作表中的標音漢字，轉成 HTML 網頁檔案。
    # ---------------------------------------------------------------------
    status_code = tng_sing_bang_iah(
        wb=wb,
        sheet_name='漢字注音',
        han_ji_source='V3',
        page_type='含頁頭'
    )
    if status_code != EXIT_CODE_SUCCESS:
        logging.error("標音漢字轉換為 HTML 網頁檔案失敗！")
        return status_code

    # ---------------------------------------------------------------------
    # 作業結尾處理
    # ---------------------------------------------------------------------
    # 要求畫面回到【漢字注音】工作表
    wb.sheets['漢字注音'].activate()
    # 作業正常結束
    logging_process_step("<----------- 作業結束！---------->")
    return EXIT_CODE_SUCCESS

# =============================================================================
# 程式主流程
# =============================================================================
def main():
    # =========================================================================
    # (0) 程式初始化
    # =========================================================================
    # 取得專案根目錄。
    current_file_path = Path(__file__).resolve()
    project_root = current_file_path.parent
    # 取得程式名稱
    # program_file_name = current_file_path.name
    program_name = current_file_path.stem

    # =========================================================================
    # (1) 開始執行程式
    # =========================================================================
    logging_process_step(f"《========== 程式開始執行：{program_name} ==========》")
    logging_process_step(f"專案根目錄為: {project_root}")

    # =========================================================================
    # (2) 設定【作用中活頁簿】：偵測及獲取 Excel 已開啟之活頁簿檔案。
    # =========================================================================
    wb = None
    # 取得【作用中活頁簿】
    try:
        wb = xw.apps.active.books.active    # 取得 Excel 作用中的活頁簿檔案
    except Exception as e:
        print(f"發生錯誤: {e}")
        logging.error(f"無法找到作用中的 Excel 工作簿: {e}", exc_info=True)
        return EXIT_CODE_NO_FILE

    # 若無法取得【作用中活頁簿】，則因無法繼續作業，故返回【作業異常終止代碼】結束。
    if not wb:
        return EXIT_CODE_NO_FILE

    # =========================================================================
    # (3) 執行【處理作業】
    # =========================================================================
    try:
        result_code = process(wb)
        if result_code != EXIT_CODE_SUCCESS:
            msg = f"程式異常終止：{program_name}"
            logging_exc_error(msg=msg, error=e)
            return EXIT_CODE_PROCESS_FAILURE

    except Exception as e:
        msg = f"程式異常終止：{program_name}"
        logging_exc_error(msg=msg, error=e)
        return EXIT_CODE_UNKNOWN_ERROR

    finally:
        #--------------------------------------------------------------------------
        # 儲存檔案
        #--------------------------------------------------------------------------
        try:
            # 要求畫面回到【漢字注音】工作表
            wb.sheets['漢字注音'].activate()
            # 儲存檔案
            file_path = save_as_new_file(wb=wb)
            if not file_path:
                logging_exc_error(msg="儲存檔案失敗！", error=e)
                return EXIT_CODE_SAVE_FAILURE    # 作業異當終止：無法儲存檔案
            else:
                logging_process_step(f"儲存檔案至路徑：{file_path}")
        except Exception as e:
            logging_exc_error(msg="儲存檔案失敗！", error=e)
            return EXIT_CODE_SAVE_FAILURE    # 作業異當終止：無法儲存檔案

        # if wb:
        #     xw.apps.active.quit()  # 確保 Excel 被釋放資源，避免開啟殘留

    # =========================================================================
    # 結束程式
    # =========================================================================
    logging_process_step(f"《========== 程式終止執行：{program_name} ==========》")
    return EXIT_CODE_SUCCESS    # 作業正常結束


if __name__ == "__main__":
    # 檢查是否有 '-2' 人工標音參數
    if "-2" in sys.argv:
        Piau_Im_Row = -2
    exit_code = main()
    sys.exit(exit_code)