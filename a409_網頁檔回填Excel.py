# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import logging
import os
import sys
from pathlib import Path

import xlwings as xw
from bs4 import BeautifulSoup
from dotenv import load_dotenv

from mod_file_access import save_as_new_file

# =========================================================================
# 常數定義
# =========================================================================
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
def get_value_by_name(wb, name):
    """利用 Excel 名稱取得指定設定值（若有的話）"""
    try:
        return wb.names[name].refers_to_range.value
    except Exception as e:
        print(f"取得名稱 {name} 失敗：{e}")
        return None

def import_html_to_excel(wb, html_file_path):
    """
    讀取 HTML 檔案，並將 head 區段中的 env 資料回填到 Excel 的 env 工作表，
    同時將 body 內以 <ruby> 與 <span> 呈現的資料填入「漢字注音」工作表中，
    其對應規則如下：
      - 漢字：<ruby> 中的 <rb>
      - 台語音標：<ruby> 中的 <rt>
      - 漢字標音：<ruby> 中的 <rtc>（或 <crt>）
      - 標點符號：<span> 的文字
    另外，每讀到一個 <p> 標籤結尾時，於「漢字注音」工作表的對應儲存格填入公式 =CHAR(10)。
    填入動作後會在 Console 輸出進度訊息。
    """
    # -------------------------------
    # 1. 讀取並解析 HTML 檔案
    # -------------------------------
    with open(html_file_path, 'r', encoding='utf-8') as f:
        html_content = f.read()
    soup = BeautifulSoup(html_content, 'html.parser')

    # -------------------------------
    # 2. 回填 env 工作表：利用 head 區段中的 meta 標籤
    # -------------------------------
    env_keys = ["FILE_NAME", "TITLE", "IMAGE_URL", "OUTPUT_PATH", "章節序號",
                "顯示注音輸入", "每頁總列數", "每列總字數", "語音類型",
                "漢字庫", "標音方法", "網頁格式", "標音方式", "上邊標音", "右邊標音", "網頁每列字數"]
    env_data = {}
    head = soup.find('head')
    if head:
        meta_tags = head.find_all('meta')
        for meta in meta_tags:
            if meta.has_attr('name') and meta.has_attr('content'):
                name = meta['name']
                content = meta['content']
                if name in env_keys:
                    env_data[name] = content
    try:
        env_sheet = wb.sheets['env']
    except Exception as e:
        print("找不到 env 工作表！", e)
        env_sheet = None
    for key, value in env_data.items():
        try:
            wb.names[key].refers_to_range.value = value
            print(f"[env] 已更新 '{key}'：{value}")
        except Exception as e:
            print(f"無法更新 env 參數 {key}：{e}")

    # -------------------------------
    # 3. 解析 body 區段的「漢字」資料
    # -------------------------------
    # 假設文章內容包在 <div class="Siang_Pai"> 中
    content_div = soup.find('div', class_='Siang_Pai')
    if not content_div:
        print("未找到 class 為 'Siang_Pai' 的 <div>！")
        return

    # 取得所有 <p> 區塊（文章段落）
    p_blocks = content_div.find_all('p')
    elements = []  # 每個元素以 dict 表示
    for p in p_blocks:
        for child in p.children:
            if child.name == 'ruby':
                # 取出 <rb>、<rt> 及 <rtc>（或 <crt>）的內容
                rb_tag = child.find('rb')
                rt_tag = child.find('rt')
                rtc_tag = child.find('rtc')
                if not rtc_tag:
                    rtc_tag = child.find('crt')
                entry = {
                    'type': 'ruby',
                    'rb': rb_tag.get_text(strip=True) if rb_tag else "",
                    'rt': rt_tag.get_text(strip=True) if rt_tag else "",
                    'rtc': rtc_tag.get_text(strip=True) if rtc_tag else ""
                }
                elements.append(entry)
            elif child.name == 'span':
                entry = {
                    'type': 'span',
                    'text': child.get_text(strip=True)
                }
                elements.append(entry)
            elif child.name == 'br':
                entry = {'type': 'line_break'}
                elements.append(entry)
            elif isinstance(child, str):
                text = child.strip()
                if text:
                    entry = {
                        'type': 'text',
                        'text': text
                    }
                    elements.append(entry)
        # 當讀完整個 <p> 區塊後，新增一個段落結尾標記
        elements.append({'type': 'p_end'})

    # -------------------------------
    # 4. 回填「漢字注音」工作表
    # -------------------------------
    try:
        sheet = wb.sheets['漢字注音']
    except Exception as e:
        print("找不到『漢字注音』工作表！", e)
        return

    # 設定填入區域（此部分應與匯出時的設定一致）
    start_row = 5      # 假設漢字儲存格起始行號為 5
    start_col = 4      # 漢字儲存格起始欄號為 4
    try:
        chars_per_row = int(get_value_by_name(wb, '每列總字數'))
    except Exception:
        chars_per_row = 15  # 若無設定，預設為 15

    rows_per_block = 4  # 匯出時每組使用 4 列：上方儲存台語音標、中央為漢字、下方為漢字標音
    current_row = start_row
    current_col = start_col

    processed_count = 0  # 記錄已處理的元素數量

    for entry in elements:
        if entry['type'] in ('line_break', 'p_end'):
            if entry['type'] == 'p_end':
                # 在目前 cell 填入公式 =CHAR(10)
                sheet.range((current_row, current_col)).formula = "=CHAR(10)"
                print(f"已填入公式 =CHAR(10) 至 cell ({current_row}, {current_col}) [p_end]")
                processed_count += 1
            # 換行：移動到下一個區塊
            current_row += rows_per_block
            current_col = start_col
            print(f"換行到下一區塊：現在起始 cell 為 ({current_row}, {current_col})")
            continue

        if entry['type'] == 'ruby':
            sheet.range((current_row, current_col)).value = entry['rb']
            sheet.range((current_row - 1, current_col)).value = entry['rt']
            sheet.range((current_row + 1, current_col)).value = entry['rtc']
            print(f"已填入 ruby：漢字 '{entry['rb']}', 台語音標 '{entry['rt']}', 漢字標音 '{entry['rtc']}' 至 cell ({current_row}, {current_col})")
            processed_count += 1
        elif entry['type'] in ('span', 'text'):
            sheet.range((current_row, current_col)).value = entry.get('text', '')
            print(f"已填入 {entry['type']}：'{entry.get('text','')}' 至 cell ({current_row}, {current_col})")
            processed_count += 1

        # 移動到下一個欄位
        current_col += 1
        if current_col >= start_col + chars_per_row:
            current_row += rows_per_block
            current_col = start_col
            print(f"自動換行：已達每列總字數，換至下一區塊，現在 cell 為 ({current_row}, {current_col})")

    EndOfText = 'φ'
    sheet.range((current_row, current_col)).value = EndOfText
    print(f"({current_row}, {current_col})：填入【文章終結符號】（{EndOfText}）")
    print(f"回填 Excel 完成，共處理 {processed_count} 個填入動作！")


def process(wb, html_file_path):
    logging_process_step("<----------- 作業開始！---------->")
    if html_file_path is None:
        try:
            title = wb.names['TITLE'].refers_to_range.value
            hue_im = wb.names['語音類型'].refers_to_range.value
            piau_im_huat = wb.names['標音方法'].refers_to_range.value
            piau_im_format = wb.names['標音方式'].refers_to_range.value
        except Exception as e:
            logging_exc_error(msg="無法自【env】工作表，取得【名稱】之設定！", error=e)
            return EXIT_CODE_INVALID_INPUT

        # 檢查檔案名稱是否已包含副檔名
        if piau_im_format == "無預設":
            im_piau = piau_im_huat
        elif piau_im_format == "上":
            im_piau = wb.names['上邊標音'].refers_to_range.value
        elif piau_im_format == "右":
            im_piau = wb.names['右邊標音'].refers_to_range.value
        else:
            im_piau = f"{wb.names['上邊標音'].refers_to_range.value}＋{wb.names['右邊標音'].refers_to_range.value}"
        # 檢查檔案名稱是否已包含副檔名
        file_name = f"《{title}》【{hue_im}】{im_piau}.html"
        file_dir_path = 'docs'
        html_file_path = os.path.join(file_dir_path, file_name)

    logging_process_step(f"開始讀取 HTML 檔案：{html_file_path}")
    import_html_to_excel(wb, html_file_path)

    # ---------------------------------------------------------------------
    # 作業結尾處理
    # ---------------------------------------------------------------------
    # 要求畫面回到【漢字注音】工作表
    wb.sheets['漢字注音'].activate()
    # 作業正常結束
    logging_process_step("<----------- 作業結束！---------->")
    return EXIT_CODE_SUCCESS


def main(html_file_path):
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
        full_file_name = Path(wb.fullname).name
        dir_path = Path(wb.fullname).parent
        main_file_name = Path(full_file_name).stem
        logging_process_step(f"作用中的 Excel 活頁簿：{full_file_name}，路徑：{dir_path}")
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
        result_code = process(wb, html_file_path)
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
    # 利用 sys.argv 取得命令列參數，第一個參數應為 HTML 檔案路徑
    # if len(sys.argv) < 2:
    #     print("請在命令列提供 HTML 檔案路徑參數。")
    #     sys.exit(1)
    # if not os.path.exists(sys.argv[1]):

    html_file_path = None
    if len(sys.argv) < 2:
        html_file_path = None
        logging_process_step(f"未指定 HTML 檔案路徑！")
    else:
        html_file_path = sys.argv[1]
        logging_process_step(f"已指定 HTML 檔案路徑：{sys.argv[1]}")

    exit_code = main(html_file_path)
    sys.exit(exit_code)