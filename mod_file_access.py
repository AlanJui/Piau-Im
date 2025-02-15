import argparse
import importlib
import logging
import os
import os.path
import subprocess
import time
from pathlib import Path

import xlwings as xw

# from openpyxl import load_workbook

# 指定虛擬環境的 Python 路徑
# venv_python = os.path.join(".venv", "Scripts", "python.exe") if sys.platform == "win32" else os.path.join(".venv", "bin", "python")

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

#----------------------------------------------------------------
# 動態載入模組和函數
#----------------------------------------------------------------
def load_module_function(module_name, function_name):
    module = importlib.import_module(module_name)
    return getattr(module, function_name)

#----------------------------------------------------------------
# 依 env 工作表的設定，另存新檔到指定目錄。
#----------------------------------------------------------------
def save_as_new_file(wb, input_file_name=None):
    # 取得檔案名稱
    if input_file_name == '_working':
        file_name = input_file_name
    else:
        # 自 env 工作表取得檔案名稱
        try:
            title = str(wb.names['TITLE'].refers_to_range.value).strip()
            file_name = f"《{title}》"
        except KeyError:
            setting_sheet = wb.sheets["env"]
            file_name = str(setting_sheet.range("C4").value).strip()

    # 設定檔案輸出路徑，存於專案根目錄下的 output2 資料夾
    output_path = wb.names['OUTPUT_PATH'].refers_to_range.value
    hue_im = wb.names['語音類型'].refers_to_range.value
    piau_im_huat = wb.names['標音方法'].refers_to_range.value
    im_piat = hue_im[:2]  # 取 hue_im 前兩個字元
    # 檢查檔案名稱是否已包含副檔名
    new_file_name = ensure_xlsx_extension(file_name)
    new_file_path = os.path.join(
        ".\\{0}".format(output_path),
        f"【河洛{im_piat}注音-{piau_im_huat}】{new_file_name}")

    # 儲存新建立的工作簿
    wb.save(new_file_path)
    return f"{new_file_path}"

#----------------------------------------------------------------
# 依 env 工作表的設定，將【漢字注音】輸出成網頁檔案。
#----------------------------------------------------------------
def save_as_html_file(wb, dir_path=None):
    # 自 env 工作表取得檔案名稱
    try:
        title = str(wb.names['TITLE'].refers_to_range.value).strip()
    except KeyError:
        setting_sheet = wb.sheets["env"]
        title = str(setting_sheet.range("C4").value).strip()

    # 設定檔案輸出路徑，存於【專案根目錄】下的【子目錄】
    # output_path = wb.names['OUTPUT_PATH'].refers_to_range.value
    output_path = "docs"
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
    file_path = os.path.join(
        ".\\{0}".format(output_path),
        f"《{title}》【{hue_im}】{im_piau}.html")

    # 儲存新建立的工作簿
    try:
        wb.save(file_path)
    except Exception as e:
        logging.error(f"儲存檔案失敗！錯誤訊息：{e}", exc_info=True)
        return EXIT_CODE_PROCESS_FAILURE

    logging_process_step(f"儲存檔案至路徑：{file_path}")
    return EXIT_CODE_SUCCESS    # 作業正常結束

#----------------------------------------------------------------
# 將 Excel 檔案，另存成網頁檔案。
#----------------------------------------------------------------
def save_excel_as_html_file(wb):
    # 自 env 工作表取得檔案名稱
    try:
        title = str(wb.names['TITLE'].refers_to_range.value).strip()
    except KeyError:
        setting_sheet = wb.sheets["env"]
        title = str(setting_sheet.range("C4").value).strip()

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
    file_path = os.path.join(
        ".\\{0}".format(output_path),
        f"《{title}》【{hue_im}-{im_piau}】.html")
    # 儲存新建立的工作簿
    try:
        wb.save(file_path)
    except Exception as e:
        logging.error(f"儲存檔案失敗！錯誤訊息：{e}", exc_info=True)
        return EXIT_CODE_PROCESS_FAILURE

    logging_process_step(f"儲存檔案至路徑：{file_path}")
    return EXIT_CODE_SUCCESS    # 作業正常結束


#----------------------------------------------------------------
# 查詢語音類型，若未設定則預設為文讀音
#----------------------------------------------------------------
def get_sound_type(wb):
    try:
        if '語音類型' in wb.names:
            reading_type = wb.names['語音類型'].refers_to_range.value
        else:
            raise KeyError
    except KeyError:
        reading_type = "文讀音"
    return reading_type

#----------------------------------------------------------------
# 查詢標音使用之【漢字庫】，預設為【河洛話】漢字庫（Tai_Loo_Han_Ji_Khoo.db）
#----------------------------------------------------------------
def get_han_ji_khoo(wb):
    try:
        if '漢字庫' in wb.names:
            han_ji_khoo = wb.names['漢字庫'].refers_to_range.value
        else:
            raise KeyError
    except KeyError:
        han_ji_khoo = "河洛話"
    return han_ji_khoo

#----------------------------------------------------------------
# 使用範例
# type = get_named_value(wb, '語音類型', default_value="文讀音")
# ca_han_ji_thak_im(wb, '漢字注音', 'V3', type)
#----------------------------------------------------------------
def get_named_value(wb, name, default_value=None):
    """
    取得 Excel 活頁簿中名稱的值，如果名稱不存在或範圍無效，則回傳預設值。

    :param wb: 打開的 Excel 活頁簿
    :param name: 名稱
    :param default_value: 預設值，如果名稱不存在或無效則回傳該值
    :return: 儲存格中的值或預設值
    """
    try:
        # 檢查名稱是否存在
        if name in wb.names:
            # 嘗試取得名稱所指的範圍
            named_range = wb.names[name].refers_to_range
            return named_range.value
        else:
            # 如果名稱不存在，回傳預設值
            return default_value
    except (AttributeError, com_error) as e:
        # 捕捉 refers_to_range 相關的錯誤，回傳預設值
        return default_value


# ==========================================================
# 自動補上 Excel 檔案的副檔名 .xlsx (單個檔案處理)
# ==========================================================
def ensure_xlsx_extension(file_name):
    return file_name if file_name.lower().endswith('.xlsx') else file_name + '.xlsx'


def ensure_extension_name(file_name, extension):
    return file_name if file_name.lower().endswith(f'.{extension}') else file_name + '.xlsx'


def get_cmd_input():
    parser = argparse.ArgumentParser(description='Process some files.')
    parser.add_argument('-d', '--dir', default='output', help='Directory path')
    parser.add_argument('-i', '--input', default='Piau_Tsu_Im', help='Input file name')
    parser.add_argument('-o', '--output', default='', help='Output file name')
    args = parser.parse_args()

    return {
        "dir_path": args.dir,  # "output
        "input": args.input,
        "output": args.output,
    }


def open_excel_file(dir_path, main_file_name):
    # 檢查檔案名稱是否已包含副檔名
    file_name, file_extension = os.path.splitext(main_file_name)
    if not file_extension:
        # 如果沒有副檔名，添加 .xlsx
        excel_file_name = file_name + '.xlsx'
    else:
        excel_file_name = main_file_name

    current_path = os.getcwd()
    file_path = os.path.join(current_path, dir_path, excel_file_name)
    try:
        wb = xw.Book(file_path)
    except Exception as e:
        print(f"檔案：`{file_path}` 無法開啟，原因為：{e}")
        return None

    return wb


#==================================================================
# 程式碼流程：
# 1. 檢查提供的文件名是否包含副檔名，如果沒有則自動添加 .xlsx。
# 2. 打開原始 Excel 文件。
# 3. 另存為一份副本到 output 子目錄下，命名為 Piau-Tsu-Im.xlsx。
# 4. 關閉原始工作簿。
# 5. 重新打開新保存的副本。
# 6. 最後返回這個新打開的工作簿對象。
#==================================================================
def save_to_a_working_copy(main_file_name):
    global wb1
    # 检查文件名称是否已包含扩展名
    file_name, file_extension = os.path.splitext(main_file_name)
    if not file_extension:
        # 如果没有扩展名，添加 .xlsx
        excel_file_name = file_name + '.xlsx'
    else:
        excel_file_name = main_file_name

    # 获取当前工作目录并构建原始文件的完整路径
    current_path = os.getcwd()
    file_path = os.path.join(current_path, "output", excel_file_name)

    # 确保 output 文件夹存在
    if not os.path.exists(os.path.dirname(file_path)):
        os.makedirs(os.path.dirname(file_path))

    # 尝试打开 Excel 文件
    try:
        wb = load_workbook(file_path)
    except FileNotFoundError as e:
        print(f"檔案：{file_path} 無法開啟，原因為：{e}")
        return None

    # 在刪除文件前確保 working.xlsx 檔案已存在。
    del_working_file()

    try:
        # 指定新保存路径和新文件名
        # new_file_path = os.path.join(current_path, "output", "Piau-Tsu-Im.xlsx")
        new_file_path = os.path.join(current_path, "working.xlsx")

        # 使用另存为将文件保存至指定路径
        wb.save(new_file_path)
    finally:
        # 无论是否成功，都关闭原始工作簿
        wb.close()


def del_working_file():
    # 在刪除文件前確保 working.xlsx 檔案已存在。
    current_path = os.getcwd()
    tmp_file_path = os.path.join(current_path, "working.xlsx")
    if os.path.exists(tmp_file_path):
        try:
            os.remove(tmp_file_path)
        except Exception as e:
            print(f"工作暫存檔刪除失敗，原因為：{e}")


def close_excel_file(excel_workbook):
    # 關閉工作簿
    excel_workbook.close()


# -----------------------------------------------------------------
# 將「字串」轉換成「串列（Characters List）」
# Python code to convert string to list character-wise
# -----------------------------------------------------------------
def convert_string_to_chars_list(string):
    list1 = []
    list1[:0] = string
    return list1


# -----------------------------------------------------------------
# 要生成超連結的目錄
# directory = 'output'
# extenstion = 'xlsx'
# exculude_list = ['Piau-Tsu-Im.xlsx', 'env.xlsx', 'env_osX.xlsx']
# -----------------------------------------------------------------
def create_file_list(directory, extension, exculude_list):
    # 建立檔案清單
    file_list = []

    # 遍歷目錄下的檔案
    for filename in os.listdir(directory):
        # 排除 index.html 和 _template.html 檔案
        if filename not in exculude_list:
            if filename.endswith(extension):
                file_list.append(filename)

    return file_list


# -----------------------------------------------------
# Backup the original file
# -----------------------------------------------------
# def open_excel_file(main_file_name):
#     # 檢查檔案名稱是否已包含副檔名
#     file_name, file_extension = os.path.splitext(main_file_name)
#     if not file_extension:
#         # 如果沒有副檔名，添加 .xlsx
#         excel_file_name = file_name + '.xlsx'
#     else:
#         excel_file_name = main_file_name

#     # 獲取當前工作目錄並構建原始檔案的完整路徑
#     current_path = os.getcwd()
#     file_path = os.path.join(current_path, "output", excel_file_name)

#     # 確保 output 資料夾存在
#     if not os.path.exists(os.path.dirname(file_path)):
#         os.makedirs(os.path.dirname(file_path))

#     # 嘗試打開 Excel 文件
#     try:
#         wb = xw.Book(file_path)
#     except FileNotFoundError:
#         print(f"File {file_path} not found.")
#         return None

#     return wb


def write_to_excel_file(excel_workbook):
    # 儲存新建立的工作簿
    try:
        excel_workbook.save()
    except Exception as e:
        print(f"存檔失敗，原因：{e}")
        return

    # 等待一段時間讓 save 完成
    time.sleep(3)

    # 取得檔案的完整路徑
    full_path = excel_workbook.fullname

    # 使用 os.path 模組來分解路徑和檔案名稱
    dir_path = os.path.dirname(full_path)
    file_name = os.path.basename(full_path)

    print(f"\n將已變更之 Excel 檔案存檔...")
    print(f"檔案路徑：{dir_path}")
    print(f"檔案名稱：{file_name}")


def save_as_excel_file(excel_workbook):
    # 自工作表「env」取得新檔案名稱
    setting_sheet = excel_workbook.sheets["env"]
    new_file_name = str(
        setting_sheet.range("C4").value
    ).strip()
    current_path = os.getcwd()
    # new_file_path = os.path.join(
    #     current_path,
    #     "output",
    #     f"【河洛話注音】{new_file_name}" + ".xlsx")
    new_file_path = os.path.join(
        current_path,
        "output",
        f"{new_file_name}.xlsx")
    print(f"儲存輸出，置於檔案：{new_file_path}")

    # 在儲存文件前確保 output 資料夾存在。如果不存在，則先創建它
    output_dir = os.path.join(os.getcwd(), "output")
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # 儲存新建立的工作簿
    try:
        excel_workbook.save(new_file_path)
    except Exception as e:
        print(f"存檔失敗，原因：{e}")
        if os.path.exists(new_file_path):
            os.remove(new_file_path)
        try:
            excel_workbook.save(new_file_path)
        except Exception as e:
            print(f"再次存檔失敗，原因：{e}")
            return

    # 等待一段時間讓 save 完成
    time.sleep(3)


def copy_excel_sheet(excel_workbook, source_name='漢字注音', sheet_name='working'):
    # 複製工作表
    try:
        source_sheet = excel_workbook.sheets[source_name]
        new_sheet = source_sheet.copy(after=source_sheet)
        new_sheet.name = sheet_name
        print(f"將【{source_name}】工作表複製成：，【{sheet_name}】工作表。")
    except Exception as e:
        print(f"複製工作表失敗，原因：{e}")
        return

    # 等待一段時間讓 copy 完成
    time.sleep(3)


#--------------------------------------------------------------------------
# 將【漢字標音】儲存格內的資料清除
#--------------------------------------------------------------------------
def reset_han_ji_piau_im_cells(wb, sheet_name='漢字注音'):
    sheet = wb.sheets[sheet_name]  # 選擇工作表
    sheet.activate()               # 將「漢字注音」工作表設為作用中工作表
    sheet.range('A1').select()     # 將 A1 儲存格設為作用儲存格

    # 取得每頁總列數 = 迴圈執行總次數
    total_rows = int(wb.names['每頁總列數'].refers_to_range.value)
    start_row_no = 6
    row_step = 4  # 每次跳過 4 行

    for i in range(total_rows):
        # 計算要清除的行號，從 start_row_no 開始，依次遞增 4 行
        current_row_no = start_row_no + i * row_step
        # 清除指定範圍的內容
        sheet.range(f'D{current_row_no}:R{current_row_no}').clear_contents()


def San_Sing_Han_Ji_Zu_Im_Piau(wb):
    # 指定來源工作表
    source_sheet = wb.sheets["工作表1"]
    source_sheet.select()

    # 取得工作表內總列數
    source_row_no = int(
        source_sheet.range("A" + str(wb.sheets[0].cells.last_cell.row)).end("up").row
    )
    print(f"source_row_no = {source_row_no}")

    # ==========================================================
    # 備妥程式需使用之工作表
    # ==========================================================
    sheet_name_list = [
        "缺字表",
        "字庫表",
        "漢字注音表",
    ]
    # -----------------------------------------------------
    # 檢查工作表是否已存在
    for sheet_name in sheet_name_list:
        sheets =  [sheet.name for sheet in wb.sheets]  # 獲取所有工作表的名稱
        if sheet_name in sheets:
            sheet = wb.sheets[sheet_name]
            try:
                sheet.select()
                sheet.clear()
                continue
            except Exception as e:
                print(e)
        else:
            # CommandError 的 Exception 發生時，表工作表不存在
            # 新增程式需使用之工作表
            print(f"工作表 {sheet_name} 不存在，正在新增...")
            wb.sheets.add(name=sheet_name)

    # 選用「漢字注音表」
    try:
        han_ji_tsu_im_paiu = wb.sheets["漢字注音表"]
        han_ji_tsu_im_paiu.select()
    except Exception as e:
        # 处理找不到 "漢字注音表" 工作表的异常
        print(e)
        print("找不到：〖漢字注音表〗工作表。")
        return False

def s(x):
    """轉成字串並去除頭尾空白，若空則回傳 None"""
    return None if (x is None or str(x).strip() == "") else str(x).strip()

# 可用 git 指令取得專案根目錄
def get_git_root():
    try:
        return subprocess.check_output(["git", "rev-parse", "--show-toplevel"], text=True).strip()
    except subprocess.CalledProcessError:
        return None  # 若不是 git repo，則返回 None

#🔍 使用 pyproject.toml 或 setup.py 來尋找根目錄
def find_project_root():
    current_dir = Path(__file__).resolve().parent
    for parent in current_dir.parents:
        if (parent / "pyproject.toml").exists() or (parent / "setup.py").exists():
            return parent
    return None  # 未找到專案根目錄