# =========================================================================
# a520_匯出成【漢字+漢字標音】文字檔.py
#
# 功能說明：
# 將【漢字注音】工作表之內的【漢字】及【漢字標音】匯出，製成純文字檔。
# 輸出格式：每一行文句佔兩列（第一列：漢字；第二列：漢字標音），行與行之間以空行分隔，如：
#
#     歸去來兮！田園將蕪胡不歸？
#     Gui ki lai e! Tian uan ziong u ho put gui?
#
# 【漢字注音】工作表之儲存格結構（每行文句佔 4 列）：
#     第 1 列：人工標音
#     第 2 列：台語音標
#     第 3 列：漢字（起始列號 5，之後每隔 4 列為下一行）
#     第 4 列：漢字標音（本程式取用之標音）
#
# 參考範例：a510_匯出成純文字檔.py
# =========================================================================

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

# =========================================================================
# 載入環境變數
# =========================================================================
load_dotenv()

# =========================================================================
# 設定日誌
# =========================================================================
logging.basicConfig(
    filename="process_log.txt",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
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

# 【漢字】儲存格若為下列全形標點，於【標音行】以對應之半形標點取代，
# 並緊接於前一個標音之後（其後補一個空白），使標音行如一般拼音文句。
PUNCTUATION_MAP = {
    "，": ",",
    "。": ".",
    "！": "!",
    "？": "?",
    "；": ";",
    "：": ":",
    "、": ",",
    "「": "“",
    "」": "”",
    "『": "‘",
    "』": "’",
    "（": "(",
    "）": ")",
    "《": "«",
    "》": "»",
    "〈": "‹",
    "〉": "›",
    "—": "—",
    "…": "…",
}

# 讀到下列句末標點後，下一個標音之字首改為大寫
SENTENCE_END_PUNCTUATIONS = {".", "!", "?"}


# =========================================================================
# Local Function
# =========================================================================
def dump_txt_file(file_path):
    """
    在螢幕 Dump 純文字檔內容。
    """
    print("\n【文字檔內容】：")
    print("========================================\n")
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            content = f.read()
            print(content)
    except FileNotFoundError:
        print(f"無法找到檔案：{file_path}")


class PiauImLineBuilder:
    """組建【標音行】：管理標音間之空白、標點緊接、句首大寫等細節。"""

    def __init__(self):
        self.tokens: list[str] = []  # 已組入之標音／標點
        self.capitalize_next = True  # 行首（或句末標點後）之標音，字首轉大寫

    def add_piau_im(self, piau_im: str):
        piau_im = str(piau_im).strip()
        if not piau_im:
            return
        if self.capitalize_next:
            piau_im = piau_im[0].upper() + piau_im[1:]
            self.capitalize_next = False
        self.tokens.append(piau_im)

    def add_punctuation(self, han_ji_punct: str):
        punct = PUNCTUATION_MAP.get(han_ji_punct)
        if punct is None:
            return
        if self.tokens:
            # 標點緊接於前一個標音之後（不留空白）
            self.tokens[-1] += punct
        else:
            self.tokens.append(punct)
        if punct in SENTENCE_END_PUNCTUATIONS:
            self.capitalize_next = True

    def build(self) -> str:
        return " ".join(self.tokens)


# =========================================================================
# 本程式主要處理作業程序
# =========================================================================
def process(wb):
    """
    自【漢字注音】工作表逐行取出【漢字】及【漢字標音】，組成純文字檔。
    """
    # --------------------------------------------------------------------------
    # 作業初始化
    # --------------------------------------------------------------------------
    logging_process_step("<----------- 作業開始！---------->")

    # 選擇工作表
    sheet = wb.sheets["漢字注音"]
    sheet.activate()

    # --------------------------------------------------------------------------
    # 自【env】設定工作表，取得處理作業所需參數
    # --------------------------------------------------------------------------
    # 設定起始及結束的【列】位址（【第5列】、【第9列】、【第13列】等列）
    TOTAL_LINES = int(wb.names["每頁總列數"].refers_to_range.value)
    ROWS_PER_LINE = 4
    start_row = 5
    end_row = start_row + (TOTAL_LINES * ROWS_PER_LINE)
    line = 1

    # 設定起始及結束的【欄】位址（【D欄=4】起）
    CHARS_PER_ROW = int(wb.names["每列總字數"].refers_to_range.value)
    start_col = 4
    end_col = start_col + CHARS_PER_ROW

    # --------------------------------------------------------------------------
    # 作業處理：逐行取出【漢字】與【漢字標音】，組合成純文字
    # --------------------------------------------------------------------------
    logging_process_step("開始【處理作業】...")
    output_lines: list[str] = []  # 每一元素為一行文句（漢字行＋標音行），或空字串（空行）
    EOF = False

    # 逐行處理作業（每行文句佔 4 列；【漢字】列於 row，【漢字標音】列於 row + 1）
    for row in range(start_row, end_row, ROWS_PER_LINE):
        # 若已到【結尾】或【超過總行數】，則跳出迴圈
        if EOF or line > TOTAL_LINES:
            break

        han_ji_line = ""  # 漢字行
        piau_im_line = PiauImLineBuilder()  # 標音行
        Two_Empty_Cells = 0

        # 逐欄取字處理
        for col in range(start_col, end_col):
            han_ji = sheet.range((row, col)).value
            piau_im = sheet.range((row + 1, col)).value

            if han_ji == "φ":  # 讀到【結尾標示】
                EOF = True
                msg = "【文字終結】"
            elif han_ji in ("\n", "<br/>"):  # 讀到【換行標示】
                msg = "【換行】"
            elif han_ji is None:  # 讀到【缺空】（儲存格未填任何字/符）
                if Two_Empty_Cells == 0:
                    Two_Empty_Cells += 1
                elif Two_Empty_Cells == 1:
                    EOF = True
                msg = "【缺空】"
            elif str(han_ji) in PUNCTUATION_MAP:  # 讀到：標點符號
                han_ji_line += str(han_ji)
                piau_im_line.add_punctuation(str(han_ji))
                msg = str(han_ji)
            else:  # 讀到：漢字
                han_ji_line += str(han_ji)
                if piau_im is not None:
                    piau_im_line.add_piau_im(piau_im)
                msg = f"{han_ji} [{piau_im}]"

            # 顯示處理進度
            col_name = xw.utils.col_name(col)  # 取得欄位名稱
            print(f"({row}, {col_name}) = {msg}")

            # 若讀到【換行】或【文字終結】，跳出逐欄取字迴圈
            if msg == "【換行】" or EOF:
                break

        # 組合本行輸出：漢字行＋標音行；空行（無漢字）則輸出空字串
        if han_ji_line:
            output_lines.append(f"{han_ji_line}\n{piau_im_line.build()}")
        else:
            output_lines.append("")

        print("\n")
        line += 1

    # 移除檔尾多餘之空行
    while output_lines and output_lines[-1] == "":
        output_lines.pop()

    # --------------------------------------------------------------------------
    # 將所有文句寫入文字檔：每行文句（漢字行＋標音行）之間以一空行分隔
    # --------------------------------------------------------------------------
    output_dir_path = wb.names["OUTPUT_PATH"].refers_to_range.value
    output_file = f"{Path(wb.name).stem}【漢字標讀音】.txt"
    output_file_path = os.path.join(output_dir_path, output_file)
    with open(output_file_path, "w", encoding="utf-8") as f:
        f.write("\n\n".join(output_lines) + "\n")
    logging_process_step(f"已成功將【漢字+漢字標音】輸出至檔案：{output_file_path}")

    # 螢幕 Dump 檔案內容
    dump_txt_file(output_file_path)

    # 作業結束前處理
    logging_process_step("完成【處理作業】...")
    return EXIT_CODE_SUCCESS


# =========================================================================
# 程式主要作業流程
# =========================================================================
def main():
    # =========================================================================
    # (1) 取得專案根目錄。
    # =========================================================================
    current_file_path = Path(__file__).resolve()
    project_root = current_file_path.parent
    logging_process_step(f"專案根目錄為: {project_root}")

    # =========================================================================
    # (2) 獲取當前作用中的 Excel 檔案。
    # =========================================================================
    wb = None
    try:
        # 嘗試獲取當前作用中的 Excel 工作簿
        wb = xw.apps.active.books.active
    except Exception as e:
        logging_process_step(f"發生錯誤: {e}")
        logging.error(f"無法找到作用中的 Excel 工作簿: {e}", exc_info=True)
        return EXIT_CODE_NO_FILE

    if not wb:
        logging_process_step("無法作業，因未有任何 Excel 檔案已開啟。")
        return EXIT_CODE_NO_FILE

    try:
        # =========================================================================
        # (3) 執行【處理作業】
        # =========================================================================
        result_code = process(wb)
        if result_code != EXIT_CODE_SUCCESS:
            logging_process_step("處理作業失敗，過程中出錯！")
            return result_code

    except Exception as e:
        print(f"執行過程中發生未知錯誤: {e}")
        logging.error(f"執行過程中發生未知錯誤: {e}", exc_info=True)
        return EXIT_CODE_UNKNOWN_ERROR

    finally:
        if wb:
            logging.info("處理完成。")

    # 結束作業
    logging.info("作業成功完成！")
    return EXIT_CODE_SUCCESS


if __name__ == "__main__":
    exit_code = main()
    if exit_code == EXIT_CODE_SUCCESS:
        print("作業正常結束！")
    else:
        print(f"作業異常終止，錯誤代碼為: {exit_code}")
    sys.exit(exit_code)
