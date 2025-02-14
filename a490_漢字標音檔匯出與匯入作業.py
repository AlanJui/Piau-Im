import json
import logging
import os
import sys

# 載入第三方套件
import xlwings as xw
from dotenv import load_dotenv

from mod_file_access import save_as_new_file

# 載入自訂模組
from mod_標音 import is_punctuation

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
EXIT_CODE_FAILURE = 1  # 失敗
EXIT_CODE_NO_FILE = 1  # 無法找到檔案
EXIT_CODE_INVALID_INPUT = 2  # 輸入錯誤
EXIT_CODE_PROCESS_FAILURE = 3  # 過程失敗
EXIT_CODE_UNKNOWN_ERROR = 99  # 未知錯誤


# =========================================================================
# 程式主要處理作業程序
# =========================================================================
def save_to_han_ji_piau_im_file(wb, output_data):
    #--------------------------------------------------------------------------------------
    # 寫入檔案
    #--------------------------------------------------------------------------------------
    # 自 env 工作表取得檔案名稱
    try:
        title = str(wb.names['TITLE'].refers_to_range.value).strip()
    except KeyError:
        title = "__working__"

    # 設定檔案輸出路徑，存於【專案根目錄】下的【子目錄】
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
    # 檢查檔案名稱是否已包含副檔名
    output_file_path = os.path.join(
        ".\\{0}".format(output_path),
        f"《{title}》【{hue_im}】{im_piau}.jsonc")

    # 儲存新建立的工作簿
    try:
        # 寫入檔案
        with open(output_file_path, 'w', encoding='utf-8') as file:
            json.dump(output_data, file, ensure_ascii=False, indent=2)
    except Exception as e:
        logging.error(f"儲存檔案失敗！錯誤訊息：{e}", exc_info=True)
        return EXIT_CODE_PROCESS_FAILURE

    logging_process_step(f"儲存檔案至路徑：{output_file_path}")
    return EXIT_CODE_SUCCESS    # 作業正常結束

# =========================================================
# 漢字標音檔匯出作業
# =========================================================
def export_to_text_file(wb):
    # 讀取 env 工作表的設定
    env_sheet = wb.sheets['env']
    env_data = {
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
    }

    # 讀取漢字注音工作表的內容
    han_ji_sheet = wb.sheets['漢字注音']
    start_row = 5
    end_row = start_row + (int(env_data["每頁總列數"]) * 4)
    start_col = 4
    end_col = start_col + int(env_data["每列總字數"])

    EndOfArticle = False
    body_data = {}
    for row in range(start_row, end_row, 4):
        if EndOfArticle: break
        for col in range(start_col, end_col):
            han_ji = han_ji_sheet.range((row, col)).value
            tai_gi_im_piau = han_ji_sheet.range((row - 1, col)).value
            ren_gong_piau_im = han_ji_sheet.range((row - 2, col)).value
            han_ji_piau_im = han_ji_sheet.range((row + 1, col)).value

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
            print(f"({row}, {xw.utils.col_name(col)}) {console}： 台語音標：{tai_gi_im_piau}；人工標音：{ren_gong_piau_im}；漢字標音：{han_ji_piau_im}")

            # 組合 body_data
            if han_ji == 'φ':  # 文章結束
                body_data["φ"] = []
                EndOfArticle = True
                break
            elif han_ji == '\n':  # 換行
                body_data["\n"] = []
                break
            elif han_ji is None or han_ji.strip() == '':  # 空白
                continue
            elif is_punctuation(han_ji):  # 標點符號
                body_data[han_ji] = []
            else:
                body_data[han_ji] = [tai_gi_im_piau, ren_gong_piau_im, han_ji_piau_im]

            if EndOfArticle: break

    # 組合 head 和 body
    output_data = {
        "head": env_data,
        "body": body_data
    }

    # 寫入檔案
    # with open(output_path, 'w', encoding='utf-8') as file:
    #     json.dump(output_data, file, ensure_ascii=False, indent=2)
    return_code = save_to_han_ji_piau_im_file(wb, output_data)
    if return_code == EXIT_CODE_SUCCESS:
        print("漢字標音檔匯出完成！")
        return EXIT_CODE_SUCCESS
    else:
        return EXIT_CODE_PROCESS_FAILURE


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
        if han_ji == "φ":  # 文章結束
            han_ji_sheet.range((current_row, current_col)).value = "φ"
            break
        elif han_ji == "\n":  # 換行
            current_row += 4
            current_col = start_col
        else:
            # 寫入漢字
            han_ji_sheet.range((current_row, current_col)).value = han_ji
            # 寫入台語音標、人工標音、漢字標音
            tai_gi_im_piau, ren_gong_piau_im, han_ji_piau_im = piau_im_data
            han_ji_sheet.range((current_row - 1, current_col)).value = tai_gi_im_piau
            han_ji_sheet.range((current_row - 2, current_col)).value = ren_gong_piau_im
            han_ji_sheet.range((current_row + 1, current_col)).value = han_ji_piau_im
            current_col += 1

    # 寫入檔案
    try:
        save_as_new_file(wb)
    except Exception as e:
        logging.error(f"儲存檔案失敗！錯誤訊息：{e}", exc_info=True)
        return EXIT_CODE_PROCESS_FAILURE

    print("漢字標音檔匯入完成！")
    return EXIT_CODE_SUCCESS

# =========================================================
# 主程式
# =========================================================
def main():
    if len(sys.argv) > 1:
        mode = sys.argv[1]
    else:
        mode = "1"

    # 打開 Excel 檔案
    # wb = xw.Book('Tai_Gi_Zu_Im_Bun.xlsx')
    try:
        wb = xw.apps.active.books.active
    except Exception as e:
        print("找不到作用中的 Excel 工作簿！", e)
        print("❌ 執行程式前請打開 Excel 檔案！")
        sys.exit(1)

    if mode == "1":
        # 匯出作業
        # return export_to_text_file(wb, 'output.jsonc')
        return export_to_text_file(wb=wb)
    elif mode == "2":
        # 匯入作業
        # file_path = os.path.join(os.getcwd(), 'output.jsonc')
        file_path = sys.argv[2]
        return import_from_text_file(wb, file_path)
    else:
        print("❌ 錯誤：請輸入有效模式：1（匯出）；2（匯入）")
        return EXIT_CODE_INVALID_INPUT


if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)