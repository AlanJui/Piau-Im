# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import logging
import os
import re
import sys
from pathlib import Path

# 載入第三方套件
import xlwings as xw
from dotenv import load_dotenv

# 載入自訂模組/函式
from mod_ca_ji_tian import HanJiTian
from mod_database import DatabaseManager
from mod_excel_access import delete_sheet_by_name, save_as_new_file
from mod_字庫 import JiKhooDict
from mod_帶調符音標 import tng_im_piau, tng_tiau_ho
from mod_標音 import (
    PiauIm,  # 漢字標音物件
    convert_tlpa_to_tl,
    tlpa_tng_han_ji_piau_im,  # 台語音標轉台語音標
)

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
from mod_logging import (
    init_logging,
    logging_exc_error,
    logging_exception,
    logging_process_step,
    logging_warning,
)

init_logging()

# =========================================================================
# 資料層類別：存放配置參數(configurations)
# =========================================================================
class ProcessConfig:
    """處理配置資料類別"""

    def __init__(self, wb, args, hanji_piau_im_sheet: str = '漢字注音'):
        self.wb = wb
        self.args = args
        # 【漢字注音】工作表描述
        self.hanji_piau_im_sheet = hanji_piau_im_sheet
        self.TOTAL_LINES = int(wb.names['每頁總列數'].refers_to_range.value)
        self.ROWS_PER_LINE = 4
        self.line_start_row = 3  # 第一行【標音儲存格】所在 Excel 列號: 3
        self.line_end_row = self.line_start_row + (self.TOTAL_LINES * self.ROWS_PER_LINE)
        self.CHARS_PER_ROW = int(wb.names['每列總字數'].refers_to_range.value)
        self.start_col = 4
        self.end_col = self.start_col + self.CHARS_PER_ROW
        self.han_ji_orgin_cell = 'V3'  # 原始漢字儲存格位置
        # 每一行【漢字標音行】組成結構
        self.jin_kang_piau_im_row_offset = 0    # 人工標音儲存格
        self.tai_gi_im_piau_row_offset = 1      # 台語音標儲存格
        self.han_ji_row_offset = 2              # 漢字儲存格
        self.han_ji_piau_im_row_offset = 3      # 漢字標音儲存格
        # 漢字起始列號
        self.han_ji_start_row = self.line_start_row + self.han_ji_row_offset
        # 初始化字典物件
        self.han_ji_khoo_name = wb.names['漢字庫'].refers_to_range.value
        self.db_name = DB_HO_LOK_UE if self.han_ji_khoo_name == '河洛話' else DB_KONG_UN
        self.ji_tian = HanJiTian(self.db_name)
        self.piau_im = PiauIm(han_ji_khoo=self.han_ji_khoo_name)
        # 標音相關
        self.piau_im_huat = wb.names['標音方法'].refers_to_range.value
        self.ue_im_lui_piat = wb.names['語音類型'].refers_to_range.value    # 文讀音或白話音


# =========================================================================
# 作業層類別：處理儲存格存放內容
# =========================================================================
class CellProcessor:
    """儲存格處理器"""

    def __init__(
        self,
        config: ProcessConfig,
        jin_kang_piau_im_ji_khoo: JiKhooDict,
        piau_im_ji_khoo: JiKhooDict,
        khuat_ji_piau_ji_khoo: JiKhooDict,
    ):
        self.config = config
        self.ji_tian = config.ji_tian
        self.piau_im = config.piau_im
        self.piau_im_huat = config.piau_im_huat
        self.ue_im_lui_piat = config.ue_im_lui_piat
        self.han_ji_khoo = config.han_ji_khoo_name
        self.jin_kang_piau_im_ji_khoo = jin_kang_piau_im_ji_khoo
        self.piau_im_ji_khoo = piau_im_ji_khoo
        self.khuat_ji_piau_ji_khoo = khuat_ji_piau_ji_khoo
        # 初始化資料庫管理器
        self.db_manager = DatabaseManager()
        self.db_manager.connect(config.db_name)


# =========================================================================
# 作業處理函數
# =========================================================================

def _initialize_ji_khoo(
    wb,
    new_jin_kang_piau_im_ji_khoo_sheet: bool,
    new_piau_im_ji_khoo_sheet: bool,
    new_khuat_ji_piau_sheet: bool,
) -> tuple[JiKhooDict, JiKhooDict, JiKhooDict]:
    """初始化字庫工作表"""

    # 人工標音字庫
    jin_kang_piau_im_sheet_name = '人工標音字庫'
    if new_jin_kang_piau_im_ji_khoo_sheet:
        delete_sheet_by_name(wb=wb, sheet_name=jin_kang_piau_im_sheet_name)
    jin_kang_piau_im_ji_khoo = JiKhooDict.create_ji_khoo_dict_from_sheet(
        wb=wb,
        sheet_name=jin_kang_piau_im_sheet_name
    )

    # 標音字庫
    piau_im_sheet_name = '標音字庫'
    if new_piau_im_ji_khoo_sheet:
        delete_sheet_by_name(wb=wb, sheet_name=piau_im_sheet_name)
    piau_im_ji_khoo = JiKhooDict.create_ji_khoo_dict_from_sheet(
        wb=wb,
        sheet_name=piau_im_sheet_name
    )

    # 缺字表
    khuat_ji_piau_name = '缺字表'
    if new_khuat_ji_piau_sheet:
        delete_sheet_by_name(wb=wb, sheet_name=khuat_ji_piau_name)
    khuat_ji_piau_ji_khoo = JiKhooDict.create_ji_khoo_dict_from_sheet(
        wb=wb,
        sheet_name=khuat_ji_piau_name
    )

    return jin_kang_piau_im_ji_khoo, piau_im_ji_khoo, khuat_ji_piau_ji_khoo


def _save_ji_khoo_to_excel(
    wb,
    jin_kang_piau_im_ji_khoo: JiKhooDict,
    piau_im_ji_khoo: JiKhooDict,
    khuat_ji_piau_ji_khoo: JiKhooDict,
):
    """儲存字庫到 Excel"""
    jin_kang_piau_im_ji_khoo.write_to_excel_sheet(wb=wb, sheet_name='人工標音字庫')
    piau_im_ji_khoo.write_to_excel_sheet(wb=wb, sheet_name='標音字庫')
    khuat_ji_piau_ji_khoo.write_to_excel_sheet(wb=wb, sheet_name='缺字表')


def _process_sheet(sheet, config: ProcessConfig, processor: CellProcessor):
    """處理整個工作表"""

    # 處理所有的儲存格
    active_cell = sheet.range(f'{xw.utils.col_name(config.start_col)}{config.line_start_row}')
    active_cell.select()

    # 調整 row 值至【漢字】列（每 4 列為一組【列群】，漢字在第 3 列：5, 9, 13, ... ）
    is_eof = False
    for r in range(1, config.TOTAL_LINES + 1):
        if is_eof: break
        line_no = r
        print('=' * 80)
        print(f"處理第 {line_no} 行...")
        row = config.line_start_row + (r - 1) * config.ROWS_PER_LINE + config.han_ji_row_offset
        new_line = False
        for c in range(config.start_col, config.end_col + 1):
            if is_eof: break
            row = row
            col = c
            active_cell = sheet.range((row, col))
            active_cell.select()
            # 處理儲存格
            print('-' * 60)
            print(f"儲存格：{xw.utils.col_name(col)}{row}（{row}, {col}）")
            is_eof, new_line = processor.process_cell(active_cell, row, col)
            if new_line: break
            if is_eof: break


# =========================================================================
# 程式區域函式
# =========================================================================
#-------------------------------------------------------------------------
# 將【缺字表】工作表，已填入【台語音標】之資料，登錄至【標音字庫】工作表
# 使用【缺字表】工作表中的【校正音標】，更正【漢字注音】工作表中之【台語音標】、【漢字標音】；
# 並依【缺字表】工作表中的【台語音標】儲存格內容，更新【標音字庫】工作表中之【台語音標】及【校正音標】欄位
#-------------------------------------------------------------------------
def update_khuat_ji_piau(wb, config: ProcessConfig, processor: CellProcessor) -> int:
    """
    讀取 Excel 檔案，依據【缺字表】工作表中的資料執行下列作業：
      1. 由 A 欄讀取漢字，從 C 欄取得原始輸入之【校正音標】，並轉換為 TLPA+ 格式，然後更新 B 欄（台語音標）。
      2. 從 D 欄讀取座標字串（可能含有多組座標），每組座標指向【漢字注音】工作表中該漢字儲存格，
         將【缺字表】取得之【台語音標】，填入【漢字注音】工作表之【台語音標】欄位（於【漢字】儲存格上方一列（row - 1））;
         並在【漢字】儲存格下方一列（row + 1）填入【漢字標音】。
    """
    # 取得【標音方法】
    piau_im_huat = config.piau_im_huat

    # 取得【漢字標音】物件
    piau_im = processor.piau_im

    # 取得【缺字表】工作表
    try:
        khuat_ji_piau_sheet_name = '缺字表'
        khuat_ji_piau_sheet = wb.sheets[khuat_ji_piau_sheet_name]
    except Exception as e:
        logging_exc_error("找不到名為『缺字表』的工作表", e)
        return EXIT_CODE_INVALID_INPUT

    # 取得【漢字注音】工作表
    try:
        han_ji_piau_im_sheet = wb.sheets["漢字注音"]
    except Exception as e:
        logging_exc_error("找不到名為『漢字注音』的工作表", e)
        return EXIT_CODE_INVALID_INPUT

    # 取得【標音字庫】查詢表（dict）
    piau_im_ji_khoo_dict = processor.piau_im_ji_khoo

    #-------------------------------------------------------------------------
    # 在【缺字表】工作表中，逐列讀取資料進行處理：【校正音標】欄（C）有填音標者，
    # 將【校正音標】正規化為 TLPA+ 格式，並更新【台語音標】欄（B）；
    # 並依據【座標】欄（D）內容，將【校正音標】填入【漢字注音】工作表中相對應之【台語音標】儲存格，
    # 以及使用【校正音標】轉換後之【漢字標音】填入【漢字注音】工作表中相對應之【漢字標音】儲存格。
    #-------------------------------------------------------------------------
    row = 2  # 從第 2 列開始（跳過標題列）
    while True:
        han_ji = khuat_ji_piau_sheet.range(f"A{row}").value  # 讀取 A 欄（漢字）
        if not han_ji:  # 若 A 欄為空，則結束迴圈
            break

        # 查檢【缺字表】中【台語音標】欄（B 欄）
        im_piau_str = khuat_ji_piau_sheet.range(f"B{row}").value
        if im_piau_str == "N/A" or not im_piau_str:  # 若 B 欄為空，則結束迴圈
            row += 1
            continue

        # 取得使用者填入的【台羅拚音】/【台語音標】並轉換為 TLPA+ 格式
        tai_gi_im_piau = tng_im_piau(im_piau_str)   # 將【音標】使用之【拼音字母】轉換成【TLPA拼音字母】；【音標調符】仍保持
        tai_gi_im_piau = tng_tiau_ho(tai_gi_im_piau).lower()  # 將【音標調符】轉換成【數值調號】

        # 更新 C 欄（校正音標）
        khuat_ji_piau_sheet.range(f"C{row}").value = tai_gi_im_piau

        # 讀取【缺字表】中【座標】欄（D 欄）的內容
        # 欄中內容可能含有多組座標，如 "(5, 17); (33, 8); (77, 5)"，表【漢字注音】工作表中有多處需要更新
        coordinates_str = khuat_ji_piau_sheet.range(f"D{row}").value
        print('-' * 80)
        print(f"{row-1}. (A{row}) ==> {coordinates_str} 【{han_ji}】： 台語音標：{im_piau_str}, 校正音標：{tai_gi_im_piau}\n")

        # 將【座標】欄位內容解析成 (row, col) 座標：此座標指向【漢字注音】工作表中之【漢字】儲存格位置
        if coordinates_str:
            # 利用正規表達式解析所有形如 (row, col) 的座標
            coordinate_tuples = re.findall(r"\((\d+)\s*,\s*(\d+)\)", coordinates_str)
            for tup in coordinate_tuples:
                try:
                    r_coord = int(tup[0])
                    c_coord = int(tup[1])
                except ValueError:
                    continue  # 若轉換失敗，跳過該組座標

                han_ji_cell = (r_coord, c_coord)  # 漢字儲存格座標

                # 根據說明，【台語音標】應填入漢字儲存格上方一列 (row - 1)，相同欄位
                target_row = r_coord - 1
                tai_gi_im_piau_cell = (target_row, c_coord)

                # 將【校正音標】填入【漢字注音】工作表漢字之【台語音標】儲存格
                han_ji_piau_im_sheet.range(tai_gi_im_piau_cell).value = tai_gi_im_piau
                excel_address = han_ji_piau_im_sheet.range(tai_gi_im_piau_cell).address
                excel_address = excel_address.replace("$", "")  # 去除 "$" 符號
                print(f"   台語音標：【{tai_gi_im_piau}】，填入【漢字注音】工作表之 {excel_address} 儲存格 = {tai_gi_im_piau_cell}")

                # 轉換【台語音標】，取得【漢字標音】
                han_ji_piau_im = tlpa_tng_han_ji_piau_im(
                    piau_im=piau_im, piau_im_huat=piau_im_huat, tai_gi_im_piau=tai_gi_im_piau
                )

                # 將【漢字標音】填入【漢字注音】工作表，【漢字】儲存格下之【漢字標音】儲存格（即：row + 1)
                target_row = r_coord + 1
                han_ji_piau_im_cell = (target_row, c_coord)

                # 將【校正音標】填入【漢字注音】工作表漢字之【台語音標】儲存格
                han_ji_piau_im_sheet.range(han_ji_piau_im_cell).value = han_ji_piau_im
                excel_address = han_ji_piau_im_sheet.range(han_ji_piau_im_cell).address
                excel_address = excel_address.replace("$", "")  # 去除 "$" 符號
                print(f"   漢字標音：【{han_ji_piau_im}】，填入【漢字注音】工作表之 {excel_address} 儲存格 = {han_ji_piau_im_cell}\n")

                # 將【漢字注音】工作表之【漢字】儲存格之底色，重置為【無底色】
                han_ji_piau_im_sheet.range(han_ji_cell).color = None

                # 更新【標音字庫】工作表之資料紀錄
                tiau_zing_piau_im_ji_khoo_dict(
                    piau_im_ji_khoo_dict=piau_im_ji_khoo_dict,
                    han_ji=han_ji,
                    tai_gi_im_piau=tai_gi_im_piau,
                    row=r_coord,
                    col=c_coord,
                )

        row += 1  # 讀取下一列

    # 依據 Dict 內容，更新【標音字庫】、【缺字表】工作表之資料紀錄
    piau_im_ji_khoo_dict.write_to_excel_sheet(wb=wb, sheet_name=piau_im_ji_khoo_dict.name)

    return EXIT_CODE_SUCCESS


def insert_or_update_to_db(db_manager: DatabaseManager, table_name: str, han_ji: str, tai_gi_im_piau: str, piau_im_huat: str):
    """
    將【漢字】與【台語音標】插入或更新至資料庫。
    使用 DatabaseManager 來管理資料庫連線和交易。

    :param db_manager: DatabaseManager 實例
    :param table_name: 資料表名稱。
    :param han_ji: 漢字。
    :param tai_gi_im_piau: 台語音標。
    :param piau_im_huat: 標音方法（用於設定常用度）。
    """
    # 確保資料表存在
    db_manager.execute(f"""
    CREATE TABLE IF NOT EXISTS {table_name} (
        識別號 INTEGER NOT NULL UNIQUE PRIMARY KEY AUTOINCREMENT,
        漢字 TEXT,
        台羅音標 TEXT,
        常用度 REAL,
        摘要說明 TEXT,
        建立時間 TEXT NOT NULL DEFAULT (DATETIME('now', 'localtime')),
        更新時間 TEXT NOT NULL DEFAULT (DATETIME('now', 'localtime'))
    );
    """)

    # 檢查是否已存在該漢字和音標的組合
    row = db_manager.fetchone(
        f"SELECT 識別號 FROM {table_name} WHERE 漢字 = ? AND 台羅音標 = ?",
        (han_ji, tai_gi_im_piau)
    )

    siong_iong_too = 0.8 if piau_im_huat == "文讀音" else 0.6

    try:
        with db_manager.transaction():
            if row:
                # 更新資料
                from datetime import datetime
                db_manager.execute(f"""
                UPDATE {table_name}
                SET 常用度 = ?, 更新時間 = ?
                WHERE 識別號 = ?;
                """, (siong_iong_too, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), row[0]))
                print(f"  ✅ 已更新：{han_ji} - {tai_gi_im_piau}")
            else:
                # 新增資料
                db_manager.execute(f"""
                INSERT INTO {table_name} (漢字, 台羅音標, 常用度, 摘要說明)
                VALUES (?, ?, ?, NULL);
                """, (han_ji, tai_gi_im_piau, siong_iong_too))
                print(f"  ✅ 已新增：{han_ji} - {tai_gi_im_piau}")
    except Exception as e:
        print(f"  ❌ 資料庫操作失敗：{han_ji} - {tai_gi_im_piau}，錯誤：{e}")
        raise


def khuat_ji_piau_poo_im_piau(wb, config: ProcessConfig, processor: CellProcessor) -> int:
    """
    讀取 Excel 的【缺字表】工作表，並將資料回填至 SQLite 資料庫。

    :param wb: Excel 活頁簿物件
    :param config: ProcessConfig 配置物件
    :param processor: CellProcessor 處理器物件
    """
    sheet_name = "缺字表"
    sheet = wb.sheets[sheet_name]
    piau_im_huat = config.piau_im_huat
    hue_im = config.ue_im_lui_piat
    table_name = "漢字庫"
    siong_iong_too = 0.8 if hue_im == "文讀音" else 0.6  # 根據語音類型設定常用度

    # 讀取資料表範圍
    data = sheet.range("A2").expand("table").value

    # 若完全無資料或只有空列，視為異常處理
    if not data or data == [[]]:
        raise ValueError("【缺字表】工作表內，無任何資料，略過後續處理作業。")

    # 若只有一列資料（如一筆記錄），資料可能不是 2D list，要包成 list
    if not isinstance(data[0], list):
        data = [data]

    idx = 0
    for row in data:
        han_ji = row[0] # 漢字
        org_tai_gi_im_piau = row[1] # 台語音標
        hau_ziann_im_piau = row[2] # 校正音標
        zo_piau = row[3] # (儲存格位置)座標

        if han_ji and (org_tai_gi_im_piau != 'N/A' or hau_ziann_im_piau != 'N/A'):
            # 將 Excel 工作表存放的【台語音標（TLPA）】，改成資料庫保存的【台羅拼音（TL）】
            tlpa_im_piau = tng_im_piau(org_tai_gi_im_piau)   # 將【音標】使用之【拼音字母】轉換成【TLPA拼音字母】；【音標調符】仍保持
            tlpa_im_piau_cleanned = tng_tiau_ho(tlpa_im_piau).lower()  # 將【音標調符】轉換成【數值調號】
            tai_gi_im_piau = convert_tlpa_to_tl(tlpa_im_piau_cleanned)

            # 使用 processor 中的 db_manager 來操作資料庫
            print('\n')
            print('-' * 80)
            print(f"📌 {idx+1}. 【{han_ji}】==> {zo_piau}：台語音標：【{tai_gi_im_piau}】（填入音標：【{org_tai_gi_im_piau}】）、校正音標：【{hau_ziann_im_piau}】、座標：{zo_piau}")
            insert_or_update_to_db(
                processor.db_manager,
                table_name,
                han_ji,
                tai_gi_im_piau,
                piau_im_huat
            )
            idx += 1

    logging_process_step(f"\n【缺字表】中的資料已成功回填至資料庫： {config.db_name} 的【{table_name}】資料表中。")
    return EXIT_CODE_SUCCESS

#--------------------------------------------------------------------------
# 重整【標音字庫】查詢表：重整【標音字庫】工作表使用之 Dict
# 依據【缺字表】工作表之【漢字】+【台語音標】資料，在【標音字庫】工作表【添增】此筆資料紀錄
#--------------------------------------------------------------------------
def tiau_zing_piau_im_ji_khoo_dict(piau_im_ji_khoo_dict,
                                    han_ji:str, tai_gi_im_piau:str,
                                    row:int, col:int):

    # Step 1: 在【標音字庫】搜尋該筆【漢字】+【台語音標】
    existing_entries = piau_im_ji_khoo_dict.ji_khoo_dict.get(han_ji, [])

    # 標記是否找到
    entry_found = False

    for existing_entry in existing_entries:
        # Step 2: 若找到，移除該筆資料內的座標
        if (row, col) in existing_entry["coordinates"]:
            existing_entry["coordinates"].remove((row, col))
        entry_found = True
        break  # 找到即可離開迴圈

    # Step 3: 將此筆資料（校正音標為 'N/A'）於【標音字庫】底端新增
    piau_im_ji_khoo_dict.add_entry(
        han_ji=han_ji,
        tai_gi_im_piau=tai_gi_im_piau,
        kenn_ziann_im_piau="N/A",  # 預設值
        coordinates=(row, col)
    )

# =========================================================================
# 本程式主要處理作業程序
# =========================================================================
def process(wb, args) -> int:
    """
    更新【漢字注音】表中【台語音標】儲存格的內容，依據【標音字庫】中的【校正音標】欄位進行更新，並將【校正音標】覆蓋至原【台語音標】。
    Args:
        wb: Excel Workbook 物件
        args: 命令列參數

    Returns:
        處理結果代碼
    """
    #--------------------------------------------------------------------------
    # 作業初始化
    #--------------------------------------------------------------------------
    logging_process_step("<=========== 作業開始！==========>")

    try:
        #--------------------------------------------------------------------------
        # 初始化 process config
        #--------------------------------------------------------------------------
        config = ProcessConfig(wb, args, hanji_piau_im_sheet='漢字注音')

        # 建立字庫工作表
        if args.new:
            jin_kang_piau_im_ji_khoo_dict, piau_im_ji_khoo_dict, khuat_ji_piau_ji_khoo_dict = _initialize_ji_khoo(
                wb=wb,
                new_jin_kang_piau_im_ji_khoo_sheet=True,
                new_piau_im_ji_khoo_sheet=True,
                new_khuat_ji_piau_sheet=True,
            )
        else:
            jin_kang_piau_im_ji_khoo_dict, piau_im_ji_khoo_dict, khuat_ji_piau_ji_khoo_dict = _initialize_ji_khoo(
                wb=wb,
                new_jin_kang_piau_im_ji_khoo_sheet=False,
                new_piau_im_ji_khoo_sheet=False,
                new_khuat_ji_piau_sheet=False,
            )

        # 建立儲存格處理器
        processor = CellProcessor(
            config=config,
            jin_kang_piau_im_ji_khoo=jin_kang_piau_im_ji_khoo_dict,
            piau_im_ji_khoo=piau_im_ji_khoo_dict,
            khuat_ji_piau_ji_khoo=khuat_ji_piau_ji_khoo_dict,
        )
    except Exception as e:
        logging.exception("處理作業，發生例外！")
        raise

    #-------------------------------------------------------------------------
    # 檢驗【漢字注音】工作表是否存在
    #-------------------------------------------------------------------------
    try:
        # 取得工作表
        han_ji_piau_im_sheet = wb.sheets['漢字注音']
        han_ji_piau_im_sheet.activate()
    except Exception as e:
        logging_exc_error(msg=f"找不到【漢字注音】工作表 ！", error=e)
        return EXIT_CODE_PROCESS_FAILURE
    logging_process_step(f"已完成作業所需之初始化設定！")

    #-------------------------------------------------------------------------
    # 【缺字表】工作表，原先找不到【音標】之漢字，已補填【台語音標】之後續處理作業
    #-------------------------------------------------------------------------
    print('\n')
    print('=' * 100)
    logging_process_step(f"開始：處理【缺字表】作業")
    try:
        sheet_name = '缺字表'
        wb.sheets[sheet_name].activate()
        update_khuat_ji_piau(wb, config, processor)
    except Exception as e:
        logging_exc_error(msg=f"處理【缺字表】作業異常！", error=e)
        return EXIT_CODE_PROCESS_FAILURE
    logging_process_step(f"完成：處理【缺字表】作業")

    #-------------------------------------------------------------------------
    # 將【缺字表】之【漢字】與【台語音標】存入【漢字庫】作業
    #-------------------------------------------------------------------------
    print('\n')
    print('=' * 100)
    logging_process_step(f"開始：將【缺字表】之【漢字】與【台語音標】存入【漢字庫】作業")
    try:
        wb.sheets['缺字表'].activate()
        khuat_ji_piau_poo_im_piau(wb, config, processor)
    except Exception as e:
        logging_exc_error(
            msg=f"將【缺字表】之【漢字】與【台語音標】存入【漢字庫】作業，發生執行異常！",
            error=e)
        return EXIT_CODE_PROCESS_FAILURE
    finally:
        # 關閉資料庫連線
        if processor.db_manager:
            processor.db_manager.disconnect()
            logging_process_step(f"已關閉資料庫連線")
    print('\n')
    print('-' * 100)
    logging_process_step(f"完成：將【缺字表】之【漢字】與【台語音標】存入【漢字庫】作業")
    print('=' * 100)

    #--------------------------------------------------------------------------
    # 結束作業
    #--------------------------------------------------------------------------
    # 寫回字庫到 Excel
    _save_ji_khoo_to_excel(
        wb=wb,
        jin_kang_piau_im_ji_khoo=jin_kang_piau_im_ji_khoo_dict,
        piau_im_ji_khoo=piau_im_ji_khoo_dict,
        khuat_ji_piau_ji_khoo=khuat_ji_piau_ji_khoo_dict,
    )
    print('\n')
    logging_process_step("<=========== 作業結束！==========>")

    return EXIT_CODE_SUCCESS

# =========================================================================
# 程式主要作業流程
# =========================================================================
def main(args) -> int:
    # =========================================================================
    # (0) 程式初始化
    # =========================================================================
    # 取得專案根目錄。
    current_file_path = Path(__file__).resolve()
    project_root = current_file_path.parent
    # 取得程式名稱
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
        exit_code = process(wb, args)
    except Exception as e:
        msg = f"程式異常終止：{program_name}"
        logging_exc_error(msg=msg, error=e)
        return EXIT_CODE_UNKNOWN_ERROR

    if exit_code != EXIT_CODE_SUCCESS:
        msg = f"程式異常終止：{program_name}（非例外，而是返回失敗碼）"
        logging.error(msg)
        return EXIT_CODE_PROCESS_FAILURE

    # =========================================================================
    # (4) 儲存檔案
    # =========================================================================
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

    # =========================================================================
    # 結束程式
    # =========================================================================
    logging_process_step(f"《========== 程式終止執行：{program_name} ==========》")
    return EXIT_CODE_SUCCESS    # 作業正常結束


# =============================================================================
# 測試程式
# =============================================================================
def test_01():
    """
    測試程式主要作業流程
    """
    print("\n\n")
    print("=" * 100)
    print("執行測試：test_01()")
    print("=" * 100)
    # 執行主要作業流程
    return EXIT_CODE_SUCCESS


# =============================================================================
# 程式作業模式切換
# =============================================================================
if __name__ == "__main__":
    import argparse
    import sys

    # 解析命令行參數
    parser = argparse.ArgumentParser(
        description='缺字表修正後續作業程式',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
使用範例：
  python a300.py          # 執行一般模式
  python a300.py -new     # 建立新的字庫工作表
  python a300.py -test    # 執行測試模式
'''
        )
    parser.add_argument(
        '--test',
        action='store_true',
        help='執行測試模式',
    )
    parser.add_argument(
        '--new',
        action='store_true',
        help='建立新的標音字庫工作表',
    )
    args = parser.parse_args()

    if args.test:
        # 執行測試
        sys.exit(test_01())
    else:
        # 從 Excel 呼叫
        sys.exit(main(args))
