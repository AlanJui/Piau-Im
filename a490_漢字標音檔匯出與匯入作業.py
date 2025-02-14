import json
import logging
import os
import sys
from collections import OrderedDict

import xlwings as xw
from dotenv import load_dotenv

from mod_file_access import save_as_new_file
from mod_標音 import is_punctuation

# =========================================================================
# 載入環境變數
# =========================================================================
load_dotenv()

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
EXIT_CODE_SUCCESS = 0  # 成功
EXIT_CODE_FAILURE = 1
EXIT_CODE_NO_FILE = 1  # 無法找到檔案
EXIT_CODE_INVALID_INPUT = 2  # 輸入錯誤
EXIT_CODE_PROCESS_FAILURE = 3  # 過程失敗
EXIT_CODE_UNKNOWN_ERROR = 99  # 未知錯誤

# =========================================================================
# 小工具：Python 值轉成 JSON 字面值文字
# =========================================================================
def py_val_to_json_literal(val):
    """
    將 Python 的 None/字串 轉成 JSON 字面值(不含最外層引號)。
    e.g. None -> null, "cun1" -> "cun1"
    """
    if val is None:
        return "null"
    # 這裡簡單做，實務上如需跳脫雙引號、跳脫字元等，可自行加 escape
    # return f"\"{val}\""
    return f"{val}"

# =========================================================================
# 將輸出資料寫成 JSONC 檔案
# =========================================================================
def save_to_han_ji_piau_im_file(wb, output_data):
    """
    將 output_data (dict) 寫入 jsonc 檔案。
    output_data["body"] 會是一個「陣列(array)」，裡頭每個元素都是一行文字。
    """
    try:
        title = str(wb.names['TITLE'].refers_to_range.value).strip()
    except KeyError:
        title = "__working__"

    output_path = wb.names['OUTPUT_PATH'].refers_to_range.value
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

    output_file_path = os.path.join(
        ".\\{0}".format(output_path),
        f"《{title}》【{hue_im}】{im_piau}.jsonc"
    )

    # 直接用 json.dump 會把 array 裡的每個元素顯示在多行（因為有換行符）
    # 若你想保留 json.dump 的漂亮排版，則會看到 "body": [ "行1", "行2" ] 各行都自動被換行。
    # 這裡只做一般輸出，可看到 "body" 最後是一個 array，每個元素是字串。
    try:
        with open(output_file_path, 'w', encoding='utf-8') as file:
            json.dump(output_data, file, ensure_ascii=False, indent=2)
    except Exception as e:
        logging.error(f"儲存檔案失敗！錯誤訊息：{e}", exc_info=True)
        return EXIT_CODE_PROCESS_FAILURE

    logging_process_step(f"儲存檔案至路徑：{output_file_path}")
    return EXIT_CODE_SUCCESS

# =========================================================================
# 匯出：將 Excel 中的 env 和「漢字注音」工作表內容，轉成「一行一筆」字串寫進 body array
# =========================================================================
def export_to_text_file(wb):
    # 1) 先組出 head 資料
    env_sheet = wb.sheets['env']
    env_data = OrderedDict({
        "格式版本": env_sheet.range("格式版本").value,
        "文件版本": env_sheet.range("文件版本").value,
        "FILE_NAME": env_sheet.range("FILE_NAME").value,
        "TITLE": env_sheet.range("TITLE").value,
        "IMAGE_URL": env_sheet.range("IMAGE_URL").value,
        "OUTPUT_PATH": env_sheet.range("OUTPUT_PATH").value,
        "章節序號": env_sheet.range("章節序號").value,
        "顯示注音輸入": env_sheet.range("顯示注音輸入").value,
        "每頁總列數": env_sheet.range("每頁總列數").value,
        "每列總字數": env_sheet.range("每列總字數").value,
        "語音類型": env_sheet.range("語音類型").value,
        "漢字庫": env_sheet.range("漢字庫").value,
        "標音方法": env_sheet.range("標音方法").value,
        "網頁格式": env_sheet.range("網頁格式").value,
        "標音方式": env_sheet.range("標音方式").value,
        "上邊標音": env_sheet.range("上邊標音").value,
        "右邊標音": env_sheet.range("右邊標音").value,
        "網頁每列字數": env_sheet.range("網頁每列字數").value
    })

    # 2) 先把 head 包成 {"head": { ... }}
    head_json_obj = {"head": env_data}
    head_json_str = json.dumps(head_json_obj, ensure_ascii=False, indent=2)
    head_json_str = head_json_str.rstrip("}")  # 移除結尾大括號，以便手動插入 body

    # 3) 準備收集 body 內容為「純文字」(一行一筆)
    han_ji_sheet = wb.sheets['漢字注音']
    start_row = 5
    end_row = start_row + int(env_data["每頁總列數"]) * 4
    start_col = 4
    end_col = start_col + int(env_data["每列總字數"])

    body_lines = []  # 改為 list 收集，每一行會有正確縮排
    EndOfArticle = False

    def s(x):
        """轉成字串並去除頭尾空白，若空則回傳 None"""
        return None if (x is None or str(x).strip() == "") else str(x).strip()

    for row in range(start_row, end_row, 4):
        for col in range(start_col, end_col):
            han_ji = han_ji_sheet.range((row, col)).value
            tai_gi_im_piau = s(han_ji_sheet.range((row - 1, col)).value)
            jin_kang_piau_im = s(han_ji_sheet.range((row - 2, col)).value)
            han_ji_piau_im = s(han_ji_sheet.range((row + 1, col)).value)

            if han_ji == 'φ':
                line_str = '    "φ": []'
                body_lines.append(line_str)
                EndOfArticle = True
                print(f'({row},{xw.utils.col_name(col)}) => 《文章終止》')
                break
            elif han_ji == '\n':
                line_str = '    "\\n": []'
                body_lines.append(line_str)
                print(f'({row},{xw.utils.col_name(col)}) => 《換行》')
                break
            elif is_punctuation(han_ji):
                line_str = f'    "{han_ji}": []'
                body_lines.append(line_str)
                print(f'({row},{xw.utils.col_name(col)}) => {han_ji}')
            else:
                def val_or_empty(x):
                    return "" if x is None else x

                line_str = (
                    f'    "{han_ji}": [ "{val_or_empty(tai_gi_im_piau)}", '
                    f'"{val_or_empty(jin_kang_piau_im)}", '
                    f'"{val_or_empty(han_ji_piau_im)}" ]'
                )
                body_lines.append(line_str)
                print(f'({row},{xw.utils.col_name(col)}) => {han_ji}：{tai_gi_im_piau}，{jin_kang_piau_im}，{han_ji_piau_im}')

        if EndOfArticle:
            break

    # 4) 取得輸出檔名
    try:
        title = str(wb.names['TITLE'].refers_to_range.value).strip()
    except KeyError:
        title = "__working__"

    output_path = wb.names['OUTPUT_PATH'].refers_to_range.value
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

    output_file_path = os.path.join(
        ".\\{0}".format(output_path),
        f"《{title}》【{hue_im}】{im_piau}.jsonc"
    )

    # 5) 將 head 先寫入，再以 append 方式把 body_lines 寫到尾端
    with open(output_file_path, 'w', encoding='utf-8') as f:
        f.write(head_json_str)  # 寫入 head 的 JSON，但不關閉大括號
        f.write(",\n  \"body\": {\n")  # "body" 開頭，縮排 2 個空白
        f.write(",\n".join(body_lines))  # 寫入 body 每一行，縮排 4 個空白
        f.write("\n  }\n}")  # 關閉 "body" 和最外層大括號

    print(f"已輸出至 {output_file_path}")
    return 0


# =========================================================
# 漢字標音檔匯入作業
# =========================================================
def import_from_text_file(wb, input_path):
    # 讀取漢字標音檔
    with open(input_path, 'r', encoding='utf-8') as file:
        data = json.load(file)

    # 回填到 env 工作表
    env_sheet = wb.sheets['env']
    for key, value in data["head"].items():
        env_sheet.range(key).value = value

    # 回填到 漢字注音 工作表
    han_ji_sheet = wb.sheets['漢字注音']
    start_row = 5
    start_col = 4
    current_row = start_row
    current_col = start_col

    for han_ji, piau_im_data in data["body"].items():
        # 控制台輸出目前處理進度狀態
        console = ""
        if han_ji == 'φ':
            console = '《文章終止》'
        if han_ji == '\n':
            console = '《換行》'
        elif is_punctuation(han_ji):  # 標點符號
            console = f"【 {han_ji} 】"
        elif han_ji is None or han_ji.strip() == '':
            console = "《空白》"
        else:
            console = f"【 {han_ji} 】"
        print(f"({current_row}, {xw.utils.col_name(current_col)}) {console}")

        # 匯入資料回填【漢字注音】工作表
        if han_ji == "φ":  # 文章結束
            han_ji_sheet.range((current_row, current_col)).value = "φ"
            break
        elif han_ji == "\n":  # 換行
            han_ji_sheet.range((current_row, current_col)).value = "=CHAR(10)"
            current_row += 4
            current_col = start_col
        else:
            if is_punctuation(han_ji):  # 標點符號
                han_ji_sheet.range((current_row, current_col)).value = han_ji
            else:
                # 寫入漢字
                han_ji_sheet.range((current_row, current_col)).value = han_ji
                # 寫入台語音標、人工標音、漢字標音
                tai_gi_im_piau, ren_gong_piau_im, han_ji_piau_im = piau_im_data
                han_ji_sheet.range((current_row - 1, current_col)).value = tai_gi_im_piau
                han_ji_sheet.range((current_row - 2, current_col)).value = ren_gong_piau_im
                han_ji_sheet.range((current_row + 1, current_col)).value = han_ji_piau_im

        current_col += 1
        if current_col > start_col + int(env_sheet.range("每列總字數").value):
            current_row += 4
            current_col = start_col

    # 寫入檔案
    try:
        save_as_new_file(wb)
    except Exception as e:
        logging.error(f"儲存檔案失敗！錯誤訊息：{e}", exc_info=True)
        return EXIT_CODE_PROCESS_FAILURE

    print("漢字標音檔匯入完成！")
    return EXIT_CODE_SUCCESS


# =========================================================================
# 主程式
# =========================================================================
def main():
    if len(sys.argv) > 1:
        mode = sys.argv[1]
    else:
        mode = "1"

    try:
        wb = xw.apps.active.books.active
    except Exception as e:
        print("找不到作用中的 Excel 工作簿！", e)
        print("❌ 執行程式前請打開 Excel 檔案！")
        sys.exit(1)

    if mode == "1":
        # 匯出
        return export_to_text_file(wb=wb)
    elif mode == "2":
        # 匯入
        file_path = sys.argv[2]
        return import_from_text_file(wb, file_path)
    else:
        print("❌ 錯誤：請輸入有效模式：1（匯出）；2（匯入）")
        return EXIT_CODE_INVALID_INPUT

if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)
