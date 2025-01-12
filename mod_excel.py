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
from p709_reset_han_ji_cells import reset_han_ji_cells

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
def maintain_han_ji_koo(wb, sheet_name, han_ji, tai_gi):
    """
    維護【漢字庫】工作表，新增或更新漢字及台語音標
    wb: Excel 活頁簿物件
    sheet_name: 工作表名稱
    han_ji: 要新增的漢字
    tai_gi: 對應的台語音標
    """
    # 取得或新增工作表
    if sheet_name not in [s.name for s in wb.sheets]:
        sheet = wb.sheets.add(sheet_name)
        print(f"已新增工作表：{sheet_name}")
        # 新增標題列
        sheet.range("A1").value = ["漢字", "台語音標", "存在總數"]
    else:
        sheet = wb.sheets[sheet_name]

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
            print(f"漢字：【{han_ji}（{tai_gi}）】已存在，總數為： {int(row[2])}")
            break

    # 若未找到則新增一筆資料
    if not found:
        records.append([han_ji, tai_gi, 1])
        print(f"新增漢字：【{han_ji}（{tai_gi}）】")


    # 更新工作表的內容
    sheet.range("A2").expand("table").clear_contents()  # 清空舊資料
    sheet.range("A2").value = records  # 寫入更新後的資料

    # print(f"已完成【漢字庫】工作表的更新！")


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


# =========================================================================
# 單元測試
# =========================================================================
def ut_maintain_han_ji_koo(wb=None):
    # 漢字庫工作表不存在：工作表將新增，且新增一筆紀錄，加入【說】字，【總數】為 1
    maintain_han_ji_koo(wb, "漢字庫", "說", "sue3")
    # 再次要求在漢字庫加入【說】：工作表會被選取，不會為【說】添增新紀錄，但【總數】更新為 2
    maintain_han_ji_koo(wb, "漢字庫", "說", "sue3")
    maintain_han_ji_koo(wb, "漢字庫", "說", "sue3")
    maintain_han_ji_koo(wb, "漢字庫", "說", "uat4")
    maintain_han_ji_koo(wb, "漢字庫", "花", "hua1")
    maintain_han_ji_koo(wb, "漢字庫", "說", "uat4")

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

# =========================================================================
# 作業程序
# =========================================================================
def process(wb):
    return_code = ut_maintain_han_ji_koo(wb=wb)
    if return_code != EXIT_CODE_SUCCESS:
        return return_code
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
            logging.info("a701_作業中活頁簿填入漢字.py 程式已執行完畢！")

    return EXIT_CODE_SUCCESS


if __name__ == "__main__":
    exit_code = main()
    if exit_code == EXIT_CODE_SUCCESS:
        print("程式正常完成！")
    else:
        print(f"程式異常終止，錯誤代碼為: {exit_code}")
    sys.exit(exit_code)