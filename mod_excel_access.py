# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import logging
import os
import re
import sys
from pathlib import Path
from typing import Optional

# 載入第三方套件
import win32com.client  # 用於獲取作用中的 Excel 檔案

# 載入第三方套件
import xlwings as xw
from dotenv import load_dotenv

# 載入自訂模組
from mod_file_access import save_as_new_file

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

# 預設應有之工作表
DEFAULT_SHEET_LIST = [
    "漢字注音",
    "缺字表",
    "字庫表",
]

# =========================================================================
# 程式用函式
# =========================================================================
def get_active_excel_file():
    """
    獲取當前作用中的 Excel 檔案路徑。
    如果沒有作用中的 Excel 檔案，返回 None。
    """
    try:
        # 獲取 Excel 應用程式
        excel_app = win32com.client.GetObject(Class="Excel.Application")
        if excel_app is None:
            print("❌ 沒有作用中的 Excel 檔案。")
            return None

        # 獲取作用中的工作簿
        active_workbook = excel_app.ActiveWorkbook
        if active_workbook is None:
            print("❌ 沒有作用中的 Excel 工作簿。")
            return None

        # 獲取檔案路徑
        excel_file = active_workbook.FullName
        print(f"✅ 作用中的 Excel 檔案：{excel_file}")
        return excel_file

    except Exception as e:
        print(f"❌ 獲取作用中的 Excel 檔案失敗: {e}")
        return None


def excel_address_to_row_col(cell_address):
    """
    將 Excel 儲存格地址 (如 'D9') 轉換為 (row, col) 格式。

    :param cell_address: Excel 儲存格地址 (如 'D9', 'AA15')
    :return: (row, col) 元組，例如 (9, 4)
    """
    match = re.match(r"([A-Z]+)(\d+)", cell_address)  # 用 regex 拆分字母(列) 和 數字(行)

    if not match:
        raise ValueError(f"無效的 Excel 儲存格地址: {cell_address}")

    col_letters, row_number = match.groups()

    # 將 Excel 字母列轉換成數字，例如 A -> 1, B -> 2, ..., Z -> 26, AA -> 27
    col_number = 0
    for letter in col_letters:
        col_number = col_number * 26 + (ord(letter) - ord("A") + 1)

    return int(row_number), col_number


def check_and_update_pronunciation(wb, han_ji, position, artificial_pronounce):
    """
    查詢【標音字庫】工作表，確認是否有該【漢字】與【座標】，
    且【校正音標】是否為 'N/A'，若符合則更新為【人工標音】。

    :param wb: Excel 活頁簿物件
    :param han_ji: 查詢的漢字
    :param position: (row, col) 該漢字的座標
    :param artificial_pronounce: 需要更新的【人工標音】
    :return: 是否更新成功 (True/False)
    """
    sheet_name = "標音字庫"

    try:
        sheet = wb.sheets[sheet_name]
    except Exception:
        print(f"⚠️ 無法找到工作表: {sheet_name}")
        return False

    # 讀取資料範圍
    data = sheet.range("A2").expand("table").value  # 讀取所有資料

    # 確保資料為 2D 列表
    if not isinstance(data[0], list):
        data = [data]

    for idx, row in enumerate(data):
        row_han_ji = row[0]  # A 欄: 漢字
        correction_pronounce_cell = sheet.range(f"D{idx+2}")  # D 欄: 校正音標
        coordinates = row[4]  # E 欄: 座標 (可能是 "(9, 4); (25, 9)" 這類格式)

        if row_han_ji == han_ji and coordinates:
            # 將座標解析成一個 set
            coord_list = coordinates.split("; ")
            parsed_coords = {convert_to_excel_address(coord) for coord in coord_list}

            # 確認該座標是否存在於【標音字庫】中
            # if convert_to_excel_address(str(position)) in parsed_coords:
            position_address = convert_to_excel_address(str(position))
            if position_address in parsed_coords:
                # 檢查標正音標是否為 'N/A'
                if correction_pronounce_cell.value == "N/A":
                    # 更新【校正音標】為【人工標音】
                    correction_pronounce_cell.value = artificial_pronounce
                    print(f"✅ 更新成功: {han_ji} ({position}) -> {artificial_pronounce}")
                    return True

    print(f"❌ 未找到匹配的資料或不符合更新條件: {han_ji} ({position})")
    return False


def convert_to_excel_address(coord_str):
    """
    轉換 `(row, col)` 格式為 Excel 座標 (如 `(9, 4)` 轉換為 "D9")

    :param coord_str: 例如 "(9, 4)"
    :return: Excel 座標字串，例如 "D9"
    """
    coord_str = coord_str.strip("()")  # 去除括號
    try:
        row, col = map(int, coord_str.split(", "))
        return f"{chr(64 + col)}{row}"  # 轉換成 Excel 座標
    except ValueError:
        return ""  # 避免解析錯誤


# def convert_to_excel_address(coord_str):
#     """
#     轉換 `(row, col)` 格式為 Excel 座標 (如 `(9, 4)` 轉換為 "D9")

#     :param coord_str: 例如 "(9, 4)"
#     :return: Excel 座標字串，例如 "D9"
#     """
#     coord_str = coord_str.strip("()")  # 去除括號
#     try:
#         row, col = map(int, coord_str.split(", "))
#         return f"{chr(64 + col)}{row}"  # 轉換成 Excel 座標
#     except ValueError:
#         return ""  # 避免解析錯誤


# def excel_address_to_row_col(cell_address):
#     """
#     將 Excel 儲存格地址 (如 'D9') 轉換為 (row, col) 格式。

#     :param cell_address: Excel 儲存格地址 (如 'D9', 'AA15')
#     :return: (row, col) 元組，例如 (9, 4)
#     """
#     match = re.match(r"([A-Z]+)(\d+)", cell_address)  # 用 regex 拆分字母(列) 和 數字(行)

#     if not match:
#         raise ValueError(f"無效的 Excel 儲存格地址: {cell_address}")

#     col_letters, row_number = match.groups()

#     # 將 Excel 字母列轉換成數字，例如 A -> 1, B -> 2, ..., Z -> 26, AA -> 27
#     col_number = 0
#     for letter in col_letters:
#         col_number = col_number * 26 + (ord(letter) - ord("A") + 1)

#     return int(row_number), col_number


def get_active_cell_info(wb):
    """
    取得目前 Excel 作用儲存格的資訊：
    - 作用儲存格的位置 (row, col)
    - 取得【漢字】的值
    - 計算【人工標音】儲存格位置，並取得【人工標音】值

    :param wb: Excel 活頁簿物件
    :return: (sheet_name, han_ji, (row, col), artificial_pronounce, (artificial_row, col))
    """
    active_cell = wb.app.selection  # 取得目前作用中的儲存格
    sheet_name = active_cell.sheet.name  # 取得所在的工作表名稱
    cell_address = active_cell.address.replace("$", "")  # 取得 Excel 格式地址 (去掉 "$")

    row, col = excel_address_to_row_col(cell_address)  # 轉換為 (row, col)

    # 取得【漢字】 (作用儲存格的值)
    han_ji = active_cell.value

    # 計算【人工標音】位置 (row-2, col) 並取得其值
    artificial_row = row - 2
    artificial_cell = wb.sheets[sheet_name].cells(artificial_row, col)
    artificial_pronounce = artificial_cell.value  # 取得人工標音的值

    return sheet_name, han_ji, (row, col), artificial_pronounce, (artificial_row, col)


def get_active_cell(wb):
    """
    獲取目前作用中的 Excel 儲存格 (Active Cell)

    :param wb: Excel 活頁簿物件 (xlwings.Book)
    :return: (工作表名稱, 儲存格地址)，如 ("漢字注音", "D9")
    """
    active_cell = wb.app.selection  # 獲取目前作用中的儲存格
    sheet_name = active_cell.sheet.name  # 獲取所在的工作表名稱
    cell_address = active_cell.address.replace("$", "")  # 取得 Excel 格式地址 (去掉 "$")

    return sheet_name, cell_address


def set_active_cell(wb, sheet_name, cell_address):
    """
    設定 Excel 作用儲存格位置。

    :param wb: Excel 活頁簿物件 (xlwings.Book)
    :param sheet_name: 目標工作表名稱 (str)
    :param cell_address: 目標儲存格位址 (如 "F33")
    """
    try:
        sheet = wb.sheets[sheet_name]  # 獲取指定工作表
        sheet.activate()  # 確保工作表為作用中的表單
        sheet.range(cell_address).select()  # 設定作用儲存格
        print(f"✅ 已將作用儲存格設為：{sheet_name} -> {cell_address}")
    except Exception as e:
        print(f"❌ 設定作用儲存格失敗: {e}")


def get_sheet_data(sheet, start_cell):
    """
    從指定工作表讀取資料，並確保返回 2D 列表。
    :param sheet: 工作表物件。
    :param start_cell: 起始儲存格（例如 "A2"）。
    :return: 2D 列表，若無資料則返回空列表。
    """
    data = sheet.range(start_cell).expand("table").value
    if data is None:
        return []
    return data if isinstance(data[0], list) else [data]


def ensure_sheet_exists(wb, sheet_name):
    """
    確保指定名稱的工作表存在，如果不存在則新增。

    :param wb: Excel 活頁簿物件。
    :param sheet_name: 工作表名稱。
    :return: 確保存在的工作表物件。
    """
    try:
        # 先確保 `wb` 不是 None，並且 `wb.sheets` 可以被存取
        if not wb or not wb.sheets:
            raise ValueError("Excel 活頁簿 `wb` 無效或未正確開啟！")

        # **使用 `name` 屬性來檢查是否存在該工作表**
        sheet_names = [sheet.name for sheet in wb.sheets]

        if sheet_name in sheet_names:
            sheet = wb.sheets[sheet_name]  # 取得現有工作表
        else:
            sheet = wb.sheets.add(sheet_name)  # 新增工作表

        return sheet

    except Exception as e:
        print(f"⚠️ 無法確保工作表存在: {e}")
        return None  # 若發生錯誤，返回 None

def delete_sheet_by_name(wb, sheet_name: str, show_msg: bool=False):
    """
    刪除指定名稱的工作表
    wb: Excel 活頁簿物件
    sheet_name: 要刪除的工作表名稱
    """
    try:
        # 檢查工作表是否存在
        if sheet_name in [sheet.name for sheet in wb.sheets]:
            sheet = wb.sheets[sheet_name]
            sheet.delete()  # 刪除工作表
            if show_msg: print(f"已成功刪除工作表：{sheet_name}")
        else:
            if show_msg: print(f"無法刪除，工作表 {sheet_name} 不存在")
    except Exception as e:
        if show_msg: print(f"刪除工作表時發生錯誤：{e}")


def get_value_by_name(wb, name):
    try:
        if name in wb.names:
            value = wb.names[name].refers_to_range.value
        else:
            raise KeyError
    except KeyError:
        value = None
    return value


def get_ji_khoo(wb, sheet_name="標音字庫"):
    """
    從 Excel 工作表中取得漢字庫
    wb: Excel 活頁簿物件
    sheet_name: 工作表名稱
    """
    # 取得或新增工作表
    if sheet_name not in [s.name for s in wb.sheets]:
        sheet = wb.sheets.add(sheet_name, after=wb.sheets['漢字注音'])
        print(f"已新增工作表：{sheet_name}")
        # 新增標題列
        sheet.range("A1").value = ["漢字", "台語音標", "總數", "校正音標"]
    else:
        sheet = wb.sheets[sheet_name]

    return sheet


def maintain_ji_khoo(sheet, han_ji, tai_gi, show_msg=False):
    """
    維護【漢字庫】工作表，新增或更新漢字及台語音標
    wb: Excel 活頁簿物件
    sheet_name: 工作表名稱
    han_ji: 要新增的漢字
    tai_gi: 對應的台語音標
    """
    # 如果台語音標為空字串，設置為"NA"（或其他標示值）
    tai_gi = tai_gi if tai_gi.strip() else "NA"

    # 取得 A、B、C 欄的所有值
    data = sheet.range("A2").expand("table").value

    # 如果只有一行資料，將其轉換為 2D 列表
    if data and not isinstance(data[0], list):
        data = [data]

    if data is None:  # 如果工作表中沒有資料
        data = []

    # 將資料轉換為標準的列表格式，並將空白欄位替換為空字串
    records = [[r if r is not None else "" for r in row] for row in data]

    # 檢查是否已存在相同的「漢字」和「台語音標」
    found = False
    for i, row in enumerate(records):
        if row[0] == han_ji and row[1] == tai_gi:
            row[2] = (row[2] if isinstance(row[2], (int, float)) else 0) + 1  # 確保存在總數是數字
            found = True
            if show_msg: print(f"漢字：【{han_ji}（{tai_gi}）】紀錄己有，總數為： {int(row[2])}")
            break

    # 若未找到則新增一筆資料
    if not found:
        records.append([han_ji, tai_gi, 1])
        if show_msg: print(f"新增漢字：【{han_ji}】（{tai_gi}）")


    # 更新工作表的內容
    sheet.range("A2").expand("table").clear_contents()  # 清空舊資料
    sheet.range("A2").value = records  # 寫入更新後的資料

    # if show_msg: print(f"已完成【漢字庫】工作表的更新！")


def get_tai_gi_by_han_ji(sheet, han_ji, show_msg=False):
    """
    根據漢字取得台語音標
    wb: Excel 活頁簿物件
    sheet_name: 工作表名稱
    han_ji: 欲查詢的漢字
    """
    # 取得 A、B 欄的所有值
    data = sheet.range("A2").expand("table").value

    if data is None:  # 如果工作表中沒有資料
        if show_msg: print("【漢字庫】工作表中沒有任何資料")
        return None

    # 確保資料為 2D 列表
    if not isinstance(data[0], list):
        data = [data]

    # 將資料轉換為標準格式，並查找對應的台語音標
    for row in data:
        han_ji_cell = row[0] if row[0] is not None else ""
        tai_gi_cell = row[1] if row[1] is not None else ""
        if han_ji_cell == han_ji:
            if show_msg: print(f"找到台語音標：【{tai_gi_cell}】")
            return tai_gi_cell

    if show_msg: print(f"漢字：【{han_ji}】不存在於【漢字庫】")
    return None


def create_dict_by_sheet(wb, sheet_name: str, allow_empty_correction: bool = False) -> Optional[dict]:
    """
    更新【標音字庫】表中的【台語音標】欄位內容，依據【漢字注音】表中的【人工標音】欄位進行更新，並將【人工標音】覆蓋至原【台語音標】。
    """
    # 取得工作表
    ji_khoo_sheet = wb.sheets[sheet_name]
    ji_khoo_sheet.activate()

    # 取得【標音字庫】表格範圍的所有資料
    data = ji_khoo_sheet.range("A2").expand("table").value

    if data is None:
        print(f"【{sheet_name}】工作表無資料")
        return None

    # 確保資料為 2D 列表
    if not isinstance(data[0], list):
        data = [data]

    # 將資料轉為字典格式，key: 漢字, value: (台語音標, 校正音標, 次數)
    han_ji_dict = {}
    for i, row in enumerate(data, start=2):
        han_ji = row[0] or ""
        tai_gi_im_piau = row[1] or ""
        total_count = int(row[2]) if len(row) > 2 and isinstance(row[2], (int, float)) else 0
        corrected_tai_gi = row[3] if len(row) > 3 else ""  # 若無 D 欄資料則設為空字串

        # 在 dict 新增一筆紀錄：（1）已填入校正音標，且校正音標不同於現有之台語音標；（2）允許校正音標為空時也加入字典
        if allow_empty_correction or (corrected_tai_gi and corrected_tai_gi != tai_gi_im_piau):
            han_ji_dict[han_ji] = (tai_gi_im_piau, corrected_tai_gi, total_count, i)  # i 為資料列索引

    # 若 han_ji_dict 為空，表查找不到【漢字】對應的【台語音標】
    if not han_ji_dict:
        print(f"無法依據【{sheet_name}】工作表，建置【字庫】字典")
        return None

    return han_ji_dict


def get_sheet_by_name(wb, sheet_name="工作表1"):
    try:
        # 嘗試取得工作表
        sheet = wb.sheets[sheet_name]
        print(f"取得工作表：{sheet_name}")
    except Exception:
        # 若不存在，則新增工作表
        print(f"無法取得，故新建工作表：{sheet_name}...")
        sheet = wb.sheets.add(sheet_name, after=wb.sheets[-1])
        print(f"新建工作表：{sheet_name}")

    # 傳回 sheet 物件
    return sheet


def prepare_working_sheets(wb, sheet_list=DEFAULT_SHEET_LIST):
    # 確認作業用工作表已存在；若無，則建置
    for sheet_name in sheet_list:
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


def get_total_rows_in_sheet(wb, sheet_name):
    try:
        # 取得指定的工作表
        sheet = wb.sheets[sheet_name]

        # 從最後一列往上尋找第一個有資料的儲存格所在的列
        last_row = sheet.range("A1048576").end("up").row

        # 若 A1 也為空，代表整個 A 欄都沒有資料
        if sheet.range(f"A{last_row}").value is None:
            total_rows = 0
        else:
            total_rows = last_row

    except Exception as e:
        print(f"無法取得工作表：{sheet_name} （錯誔訊息：{e}）")
        total_rows = 0

    return total_rows


#--------------------------------------------------------------------------
# 將待注音的【漢字儲存格】，文字顏色重設為黑色（自動 RGB: 0, 0, 0）；填漢顏色重設為無填滿
#--------------------------------------------------------------------------
def reset_han_ji_cells(wb, sheet_name='漢字注音'):
    # 選擇指定的工作表
    sheet = wb.sheets[sheet_name]
    sheet.activate()  # 將「漢字注音」工作表設為作用中工作表
    sheet.range('A1').select()     # 將 A1 儲存格設為作用儲存格

    # 每頁最多處理的列數
    TOTAL_ROWS = int(wb.names['每頁總列數'].refers_to_range.value)  # 從名稱【每頁總列數】取得值
    # 每列最多處理的字數
    CHARS_PER_ROW = int(wb.names['每列總字數'].refers_to_range.value)  # 從名稱【每列總字數】取得值

    # 設定起始及結束的欄位（【D欄=4】到【R欄=18】）
    start_col = 4
    end_col = start_col + CHARS_PER_ROW

    # 從第 5 列開始，每隔 4 列進行重置（5, 9, 13, ...）
    for row in range(5, 5 + 4 * TOTAL_ROWS, 4):
        for col in range(start_col, end_col):
            cell = sheet.range((row, col))
            # 將文字顏色設為【自動】（黑色）
            cell.font.color = (0, 0, 0)  # 設定為黑色
            # 將儲存格的填滿色彩設為【無填滿】
            cell.color = None

    print("漢字儲存格已成功重置，文字顏色設為自動，填滿色彩設為無填滿。")

    return 0


#--------------------------------------------------------------------------
# 清除儲存格內容
#--------------------------------------------------------------------------
def clear_han_ji_kap_piau_im(wb, sheet_name='漢字注音'):
    sheet = wb.sheets[sheet_name]   # 選擇工作表
    sheet.activate()               # 將「漢字注音」工作表設為作用中工作表

    # 每頁最多處理的列數
    TOTAL_ROWS = int(wb.names['每頁總列數'].refers_to_range.value)  # 從名稱【每頁總列數】取得值

    cells_per_row = 4
    end_of_rows = int((TOTAL_ROWS * cells_per_row ) + 2)
    cells_range = f'D3:R{end_of_rows}'

    sheet.range(cells_range).clear_contents()     # 清除 C3:R{end_of_row} 範圍的內容


# =========================================================================
# 單元測試
# =========================================================================
def ut_get_sheet_data(wb=None):
    if not wb:
        wb = xw.Book('Test_Case_Sample.xlsx')
    sheet = wb.sheets['漢字注音']
    data = get_sheet_data(sheet, 'D5')
    for row in data:
        print(row)
    return EXIT_CODE_SUCCESS

def ut_khuat_ji_piau(wb=None):
    """缺字表登錄單元測試"""
    wb = xw.Book('Test_Case_Sample.xlsx')
    wb.activate()
    delete_sheet_by_name(wb, "缺字表", show_msg=True)
    sheet = get_ji_khoo(wb, "缺字表")
    sheet.activate()

    try:
        # 當【缺字表】工作表，尚不存在任何查找不到【標音】的【漢字】，新增一筆紀錄
        maintain_ji_khoo(sheet, "銜", "", show_msg=True)
        # 當【缺字表】已有一筆紀錄，新增第二筆紀錄
        maintain_ji_khoo(sheet, "暉", "", show_msg=True)
        # 在【缺字表】新增第三紀錄
        maintain_ji_khoo(sheet, "霪", "", show_msg=True)
    except Exception as e:
        print(e)
        return EXIT_CODE_UNKNOWN_ERROR

    # 檢查【缺字表】工作表的內容
    for row in sheet.range("A2").expand("table").value:
        print(row)
    return EXIT_CODE_SUCCESS


def ut_maintain_han_ji_koo(wb=None):
    wb = xw.Book('Test_Case_Sample.xlsx')
    sheet = get_ji_khoo(wb, "漢字庫")

    # 漢字庫工作表不存在：工作表將新增，且新增一筆紀錄，加入【說】字，【總數】為 1
    maintain_ji_khoo(sheet, "說", "sue3", show_msg=True)
    # 再次要求在漢字庫加入【說】：工作表會被選取，不會為【說】添增新紀錄，但【總數】更新為 2
    maintain_ji_khoo(sheet, "說", "sue3", show_msg=True)
    maintain_ji_khoo(sheet, "說", "sue3", show_msg=True)
    maintain_ji_khoo(sheet, "說", "uat4", show_msg=True)
    maintain_ji_khoo(sheet, "花", "hua1", show_msg=True)
    maintain_ji_khoo(sheet, "說", "uat4", show_msg=True)

    # 查詢【漢字】的台語音標
    print("\n===================================================")
    han_ji = "說"
    tai_gi = get_tai_gi_by_han_ji(sheet, han_ji)
    if tai_gi:
        print(f"查到【{han_ji}】的台語音標為：{tai_gi}")
    else:
        print(f"查不到【{han_ji}】的台語音標！")

    print("\n===================================================")
    han_ji = "龓"
    tai_gi = get_tai_gi_by_han_ji(sheet, han_ji)
    if tai_gi:
        print(f"查到【{han_ji}】的台語音標為：{tai_gi}")
    else:
        print(f"查不到【{han_ji}】的台語音標！")

    return EXIT_CODE_SUCCESS

def ut_prepare_working_sheets(wb=None):
    if not wb:
        wb = xw.Book()

    #  工作表已存在
    try:
        prepare_working_sheets(wb)
        print("工作表已存在")
    except Exception as e:
        print(e)
        return EXIT_CODE_UNKNOWN_ERROR

    #  工作表不存在
    try:
        prepare_working_sheets(wb, sheet_list=["工作表1", "工作表2"])
        print("工作表不存在")
    except Exception as e:
        print(e)
        return EXIT_CODE_UNKNOWN_ERROR

    return EXIT_CODE_SUCCESS

def ut_get_sheet_by_name(wb=None):
    if not wb:
        wb = xw.Book()

    #  工作表已存在
    try:
        sheet = get_sheet_by_name(wb, "漢字注音")
        print(sheet.name)
    except Exception as e:
        print(e)
        return EXIT_CODE_UNKNOWN_ERROR

    #  工作表不存在
    try:
        sheet = get_sheet_by_name(wb, "字庫表")
        print(sheet.name)
    except Exception as e:
        print(e)
        return EXIT_CODE_UNKNOWN_ERROR

    return EXIT_CODE_SUCCESS

def ut_get_total_rows_in_sheet(wb=None, sheet_name="字庫表"):
    #  工作表已存在
    try:
        total_rows = get_total_rows_in_sheet(wb, sheet_name)
        print(f"工作表 {sheet_name} 共有 {total_rows} 列")
    except Exception as e:
        print(e)
        return EXIT_CODE_UNKNOWN_ERROR

    #  工作表無資料
    sheet_name = "工作表1"
    try:
        total_rows = get_total_rows_in_sheet(wb, sheet_name)
        print(f"工作表 {sheet_name} 共有 {total_rows} 列")
    except Exception as e:
        print(e)
        return EXIT_CODE_UNKNOWN_ERROR

    #  工作表不存在
    sheet_name = "X"
    try:
        total_rows = get_total_rows_in_sheet(wb, sheet_name)
        print(f"工作表 {sheet_name} 共有 {total_rows} 列")
    except Exception as e:
        print(e)
        return EXIT_CODE_UNKNOWN_ERROR

    return EXIT_CODE_SUCCESS

def ut01_取得當前作用儲存格(wb):
    # 作業流程：獲取當前作用中的 Excel 儲存格
    sheet_name, cell_address = get_active_cell(wb)
    print(f"✅ 目前作用中的儲存格：{sheet_name} 工作表 -> {cell_address}")

    # 將 Excel 儲存格地址轉換為 (row, col) 格式
    row, col = excel_address_to_row_col(cell_address)
    print(f"📌 Excel 位址 {cell_address} 轉換為 (row, col): ({row}, {col})")

    # 取得作用中儲存格的值
    active_cell = wb.sheets[sheet_name].range(cell_address)
    cell_value = active_cell.value
    print(f"📌 作用儲存格{cell_address}的值為：{cell_value}")

    # 將 (row, col) 格式轉換為 Excel 儲存格地址
    # new_cell_address = convert_to_excel_address(f"({row}, {col})")
    new_cell_address = convert_to_excel_address(cell_value)
    print(f"📌 {cell_value} 座標，其 Excel 位址為：{new_cell_address}")

    # 利用 Excel 儲存格地址，將【標音字庫】工作表的 Excel 儲存格位置設為作用儲存格
    target_sheet = "漢字注音"
    target_cell_address = new_cell_address
    set_active_cell(wb, target_sheet, target_cell_address)


    return EXIT_CODE_SUCCESS


def ut02_利用列欄座標值定位漢字注音儲存格(wb):
    sheet_name = "人工標音字庫"
    cell_address = "E2"
    set_active_cell(wb, sheet_name, cell_address)

    # 取得作用中儲存格的值
    active_cell = wb.sheets[sheet_name].range(cell_address)
    cell_value = active_cell.value
    print(f"📌 作用儲存格{cell_address}的值為：{cell_value}")

    # 將 (row, col) 格式轉換為 Excel 儲存格地址
    new_cell_address = convert_to_excel_address(cell_value)
    print(f"📌 {cell_value} 座標，其 Excel 位址為：{new_cell_address}")

    # 利用 Excel 儲存格地址，將【標音字庫】工作表的 Excel 儲存格位置設為作用儲存格
    target_sheet = "漢字注音"
    target_cell_address = new_cell_address
    set_active_cell(wb, target_sheet, target_cell_address)


    return EXIT_CODE_SUCCESS


# =========================================================================
# 作業程序
# =========================================================================
def process(wb):
    return_code = ut02_利用列欄座標值定位漢字注音儲存格(wb=wb)
    if return_code != EXIT_CODE_SUCCESS:
        return return_code
    # ---------------------------------------------------------------------
    # return_code = ut01_取得當前作用儲存格(wb=wb)
    # if return_code != EXIT_CODE_SUCCESS:
    #     return return_code
    # ---------------------------------------------------------------------
    # return_code = ut_get_sheet_data(wb=wb)
    # if return_code != EXIT_CODE_SUCCESS:
    #     return return_code
    # ---------------------------------------------------------------------
    # return_code = ut_khuat_ji_piau(wb=wb)
    # if return_code != EXIT_CODE_SUCCESS:
    #     return return_code
    # han_ji_dict = create_dict_by_sheet(wb=wb, sheet_name='缺字表', allow_empty_correction=True)
    # han_ji = '霪'
    # if han_ji_dict and han_ji in han_ji_dict:
    #     original_tai_gi, corrected_tai_gi, total_count, row_index_in_ji_khoo = han_ji_dict[han_ji]
    #     if not corrected_tai_gi:
    #         corrected_tai_gi = "NA"
    #     print(f"【{han_ji}】的台語音標為：{original_tai_gi}，校正音標為：{corrected_tai_gi}，總數：{total_count}，列索引：{row_index_in_ji_khoo}")
    # else:
    #     return EXIT_CODE_PROCESS_FAILURE
    # ---------------------------------------------------------------------
    # return_code = ut_maintain_han_ji_koo(wb=wb)
    # if return_code != EXIT_CODE_SUCCESS:
    #     return return_code
    # ---------------------------------------------------------------------
    # return_code = ut_prepare_working_sheets(wb=wb)
    # if return_code != EXIT_CODE_SUCCESS:
    #     return return_code
    # ---------------------------------------------------------------------
    # return_code = ut_get_sheet_by_name(wb=wb)
    # if return_code != EXIT_CODE_SUCCESS:
    #     return return_code
    # ---------------------------------------------------------------------
    # return_code = ut_get_total_rows_in_sheet(wb=wb)
    # if return_code != EXIT_CODE_SUCCESS:
    #     return return_code
    # ---------------------------------------------------------------------

    return EXIT_CODE_SUCCESS

# =============================================================================
# 程式主流程
# =============================================================================
def main():
    logging.info("作業開始")

    # =========================================================================
    # (1) 取得專案根目錄
    # =========================================================================
    current_file_path = Path(__file__).resolve()
    project_root = current_file_path.parent
    print(f"專案根目錄為: {project_root}")
    logging.info(f"專案根目錄為: {project_root}")

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
            logging_process_step("作業異常終止！")
            return result_code

    except Exception as e:
        print(f"作業過程發生未知的異常錯誤: {e}")
        logging.error(f"作業過程發生未知的異常錯誤: {e}", exc_info=True)
        return EXIT_CODE_UNKNOWN_ERROR

    finally:
        if wb:
            # xw.apps.active.quit()  # 確保 Excel 被釋放資源，避免開啟殘留
            print("程式已執行完畢！")

    return EXIT_CODE_SUCCESS


if __name__ == "__main__":
    exit_code = main()
    if exit_code == EXIT_CODE_SUCCESS:
        print("程式正常完成！")
    else:
        print(f"程式異常終止，錯誤代碼為: {exit_code}")
    sys.exit(exit_code)