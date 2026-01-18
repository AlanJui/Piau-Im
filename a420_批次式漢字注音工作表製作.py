import logging
import os
import sys
from typing import Tuple

import xlwings as xw
from dotenv import load_dotenv

from mod_ca_ji_tian import HanJiTian
from mod_excel_access import delete_sheet_by_name, ensure_sheet_exists
from mod_logging import logging_exception
from mod_字庫 import JiKhooDict
from mod_帶調符音標 import is_han_ji
from mod_標音 import (
    PiauIm,
    ca_ji_tng_piau_im,
    convert_tl_with_tiau_hu_to_tlpa,
    is_han_ji,
    is_punctuation,
    kam_si_u_tiau_hu,
    split_hong_im_hu_ho,
    tlpa_tng_han_ji_piau_im,
)

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
# 作業用類別
# =========================================================================
class ProcessConfig:
    """處理配置資料類別"""

    def __init__(self, wb, hanji_piau_im_sheet: str):
        # 【漢字注音】工作表描述
        self.hanji_piau_im_sheet = hanji_piau_im_sheet
        self.TOTAL_LINES = int(wb.names['每頁總列數'].refers_to_range.value)
        self.ROWS_PER_LINE = 4
        self.line_start_row = 3  # 第一行【標音儲存格】所在 Excel 列號: 3
        self.line_end_row = self.line_start_row + (self.TOTAL_LINES * self.ROWS_PER_LINE)
        self.CHARS_PER_ROW = int(wb.names['每列總字數'].refers_to_range.value)
        self.start_col = 4
        self.end_col = self.start_col + self.CHARS_PER_ROW
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


class CellProcessor:
    """儲存格處理器"""

    def __init__(
        self,
        config: ProcessConfig,
        jin_kang_piau_im_ji_khoo: JiKhooDict,
        piau_im_ji_khoo: JiKhooDict,
        khuat_ji_piau_ji_khoo: JiKhooDict,
    ):
        self.ji_tian = config.ji_tian
        self.piau_im = config.piau_im
        self.piau_im_huat = config.piau_im_huat
        self.ue_im_lui_piat = config.ue_im_lui_piat
        self.han_ji_khoo = config.han_ji_khoo_name
        self.jin_kang_piau_im_ji_khoo = jin_kang_piau_im_ji_khoo
        self.piau_im_ji_khoo = piau_im_ji_khoo
        self.khuat_ji_piau_ji_khoo = khuat_ji_piau_ji_khoo

    def _reset_cell_style(self, cell):
        """重置儲存格樣式"""
        cell.font.color = (0, 0, 0)  # 黑色
        cell.color = None  # 無填滿

    def _cu_jin_kang_piau_im(self, han_ji: str, jin_kang_piau_im: str, piau_im: PiauIm, piau_im_huat: str):
        """
        取人工標音【台語音標】
        """

        tai_gi_im_piau = ''
        han_ji_piau_im = ''

        # 清除使用者輸入前後的空白，避免後續拆解時產生「空白聲母」導致注音前多一格空白
        jin_kang_piau_im = (jin_kang_piau_im or "").strip()

        if '〔' in jin_kang_piau_im and '〕' in jin_kang_piau_im:
            # 將人工輸入的〔台語音標〕轉換成【方音符號】
            im_piau = jin_kang_piau_im.split('〔')[1].split('〕')[0].strip()
            tai_gi_im_piau = convert_tl_with_tiau_hu_to_tlpa(im_piau)
            # 依使用者指定之【標音方法】，將【台語音標】轉換成其所需之【漢字標音】
            han_ji_piau_im = tlpa_tng_han_ji_piau_im(
                piau_im=piau_im,
                piau_im_huat=piau_im_huat,
                tai_gi_im_piau=tai_gi_im_piau
            )
        elif '【' in jin_kang_piau_im and '】' in jin_kang_piau_im:
            # 將人工輸入的【方音符號】轉換成【台語音標】
            han_ji_piau_im = jin_kang_piau_im.split('【')[1].split('】')[0].strip()
            siann, un, tiau = split_hong_im_hu_ho(han_ji_piau_im)
            # 依使用者指定之【標音方法】，將【台語音標】轉換成其所需之【漢字標音】
            tai_gi_im_piau = piau_im.hong_im_tng_tai_gi_im_piau(
                siann=siann,
                un=un,
                tiau=tiau)['台語音標']
        elif jin_kang_piau_im.startswith('=') and jin_kang_piau_im.endswith('='):
            # 若【人工標音】欄輸入為【=】，表【台語音標】欄自【人工標音字庫】工作表之【台語音標】欄取標音
            tai_gi_im_piau = self.jin_kang_piau_im_ji_khoo.get_tai_gi_im_piau_by_han_ji(han_ji=han_ji)
            # 若查無結果，則設為空字串
            if not tai_gi_im_piau:
                tai_gi_im_piau = ''
                han_ji_piau_im = ''
            else:
                # 依指定之【標音方法】，將【台語音標】轉換成其所需之【漢字標音】
                han_ji_piau_im = tlpa_tng_han_ji_piau_im(
                    piau_im=piau_im,
                    piau_im_huat=piau_im_huat,
                    tai_gi_im_piau=tai_gi_im_piau
                )
        else:
            # 查檢輸入的【人工標音】是否為帶【調號】的【台語音標】或【台羅拼音】
            if kam_si_u_tiau_hu(jin_kang_piau_im):
                # 將帶【聲調符號】的【人工標音】，轉換為帶【調號】的【台語音標】
                tai_gi_im_piau = convert_tl_with_tiau_hu_to_tlpa(jin_kang_piau_im)
            else:
                tai_gi_im_piau = jin_kang_piau_im
            # 依指定之【標音方法】，將【台語音標】轉換成其所需之【漢字標音】
            han_ji_piau_im = tlpa_tng_han_ji_piau_im(
                piau_im=piau_im,
                piau_im_huat=piau_im_huat,
                tai_gi_im_piau=tai_gi_im_piau
            )

        return tai_gi_im_piau, han_ji_piau_im

    def _process_jin_kang_piau_im(self, jin_kang_piau_im, cell, row, col):
        """處理人工標音內容"""
        # 預設未能依【人工標音】欄，找到對應的【台語音標】和【漢字標音】
        sing_kong = False
        han_ji=cell.value
        jin_kang_piau_im = str(jin_kang_piau_im).strip()

        # 自【人工標音】儲存格，取出【人工標音】
        tai_gi_im_piau, han_ji_piau_im = self._cu_jin_kang_piau_im(
            han_ji=han_ji,
            jin_kang_piau_im=str(jin_kang_piau_im),
            piau_im=self.piau_im,
            piau_im_huat=self.piau_im_huat,
        )
        if tai_gi_im_piau != '':
            # 依指定之【標音方法】，將【台語音標】轉換成其所需之【漢字標音】
            # han_ji_piau_im = tlpa_tng_han_ji_piau_im(
            #     piau_im=self.piau_im,
            #     piau_im_huat=self.piau_im_huat,
            #     tai_gi_im_piau=tai_gi_im_piau
            # )
            # 記錄到人工標音字庫
            self.jin_kang_piau_im_ji_khoo.add_entry(
                han_ji=han_ji,
                tai_gi_im_piau=tai_gi_im_piau,
                kenn_ziann_im_piau='N/A',
                coordinates=(row, col)
            )
            # 記錄到標音字庫
            self.piau_im_ji_khoo.add_entry(
                han_ji=han_ji,
                tai_gi_im_piau=tai_gi_im_piau,
                kenn_ziann_im_piau='N/A',
                coordinates=(row, col)
            )
            print(f"已將【{cell.value}】之【人工標音】記錄到【人工標音字庫】工作表的【校正音標】儲存格。")
            sing_kong = True
        elif  tai_gi_im_piau == '' and cell.offset(-2, 0).value == '=':
            # 【人工標音】欄輸入為【=】，但【人工標音字庫】工作表查無結果，再自【標音字庫】工作表，嚐試查找【台語音標】
            tai_gi_im_piau = self.jin_kang_piau_im_ji_khoo.get_tai_gi_im_piau_by_han_ji(han_ji=han_ji)
            if tai_gi_im_piau != '':
                # 依指定之【標音方法】，將【台語音標】轉換成【漢字標音】
                han_ji_piau_im = tlpa_tng_han_ji_piau_im(
                    piau_im=self.piau_im,
                    piau_im_huat=self.piau_im_huat,
                    tai_gi_im_piau=tai_gi_im_piau
                )
                # cell.offset(-2, 0).value = ''  # 清空【人工標音】欄【=】
                print(f"自【人工標音字庫】引用【{cell.value}】既有的【人工標音】。")
                sing_kong = True
            else:
                # 若在【人工標音字庫】找不到【人工標音】對應的【台語音標】，則自【標音字庫】工作表查找
                tai_gi_im_piau = self.piau_im_ji_khoo.get_tai_gi_im_piau_by_han_ji(han_ji=han_ji)
                if tai_gi_im_piau != '':
                    # 依指定之【標音方法】，將【台語音標】轉換成【漢字標音】
                    han_ji_piau_im = tlpa_tng_han_ji_piau_im(
                        piau_im=self.piau_im,
                        piau_im_huat=self.piau_im_huat,
                        tai_gi_im_piau=tai_gi_im_piau
                    )
                    cell.offset(-2, 0).value = ''  # 清空【人工標音】欄【=】
                    print(f"【{cell.value}】的【人工標音】，在【人工標音字庫】找不到，改用【標音字庫】中的【台語音標】替代。")
                    sing_kong = True
                else:
                    # 若找不到【人工標音】對應的【台語音標】，則記錄到【缺字表】
                    self.khuat_ji_piau_ji_khoo.add_entry(
                        han_ji=cell.value,
                        tai_gi_im_piau='N/A',
                        kenn_ziann_im_piau='N/A',
                        coordinates=(row, col)
                    )

        if sing_kong:
            # 寫入儲存格
            cell.offset(-1, 0).value = tai_gi_im_piau  # 上方儲存格：台語音標
            cell.offset(1, 0).value = han_ji_piau_im    # 下方儲存格：漢字標音
            msg = f"{han_ji}： [{tai_gi_im_piau}] /【{han_ji_piau_im}】（人工標音：【{jin_kang_piau_im}】）"
        else:
            msg = f"找不到【{han_ji}】此字的【台語音標】！"

        # 依據【人工標音】欄，在【人工標音字庫】工作表找到相對應的【台語音標】和【漢字標音】
        print(f"漢字儲存格：{xw.utils.col_name(col)}{row}（{row}, {col}）==> {msg}")

    def _process_non_han_ji(self, cell_value) -> Tuple[str, bool]:
        """處理非漢字內容"""
        if cell_value is None or str(cell_value).strip() == "":
            return "【空白】", False

        str_value = str(cell_value).strip()

        if is_punctuation(str_value):
            return f"{cell_value}【標點符號】", False
        elif isinstance(cell_value, float) and cell_value.is_integer():
            return f"{int(cell_value)}【英/數半形字元】", False
        else:
            return f"{cell_value}【其他字元】", False

    def _convert_piau_im(self, entry) -> Tuple[str, str]:
        """
        將查詢結果轉換為音標

        Args:
            result: 查詢結果列表

        Returns:
            (tai_gi_im_piau, han_ji_piau_im)
        """
        # 使用原有的轉換邏輯
        # 這裡需要適配 result 的格式
        # 假設 result 是從 HanJiSuTian 回傳的格式
        tai_gi_im_piau, han_ji_piau_im = ca_ji_tng_piau_im(
            entry=entry,
            han_ji_khoo=self.han_ji_khoo,
            piau_im=self.piau_im,
            piau_im_huat=self.piau_im_huat
        )
        return tai_gi_im_piau, han_ji_piau_im


    def _process_han_ji(
        self,
        han_ji: str,
        cell,
        row: int,
        col: int,
    ) -> Tuple[str, bool]:
        """處理漢字"""
        if han_ji == '':
            return "【空白】", False

        # 使用 HanJiTian 查詢漢字讀音
        result = self.ji_tian.han_ji_ca_piau_im(
            han_ji=han_ji,
            ue_im_lui_piat=self.ue_im_lui_piat
        )

        # 查無此字
        if not result:
            self.khuat_ji_piau_ji_khoo.add_entry(
                han_ji=han_ji,
                tai_gi_im_piau='',
                kenn_ziann_im_piau='N/A',
                coordinates=(row, col)
            )
            return f"【{han_ji}】查無此字！", False

        # 有多個讀音 len(result) > 1
        print(f"漢字儲存格：{xw.utils.col_name(col)}{row}（{row}, {col}）：【{han_ji}】有 {len(result)} 個讀音...")
        for idx, entry in enumerate(result):
            # 轉換音標
            tai_gi_im_piau, han_ji_piau_im = self._convert_piau_im(entry)

            # 寫入儲存格
            cell.offset(-1, 0).value = tai_gi_im_piau  # 上方儲存格：台語音標
            cell.offset(1, 0).value = han_ji_piau_im    # 下方儲存格：漢字標音

            msg = f"{han_ji}： [{tai_gi_im_piau}] /【{han_ji_piau_im}】"

            # 顯示處理進度
            col_name = xw.utils.col_name(col)
            print(f"{idx + 1}. {msg}")

    def process_cell(
        self,
        cell,
        row: int,
        col: int,
    ) -> bool:
        """
        處理單一儲存格

        Returns:
            is_eof: 是否已達文件結尾
            new_line: 是否需換行
        """
        # 初始化樣式
        self._reset_cell_style(cell)

        cell_value = cell.value

        jin_kang_piau_im = cell.offset(-2, 0).value  # 人工標音
        if jin_kang_piau_im and str(jin_kang_piau_im).strip() != "":
            # 處理人工標音內容
            self._process_jin_kang_piau_im(jin_kang_piau_im, cell, row, col)
            return False, False

        # 檢查特殊字元
        if cell_value == 'φ':
            # 【文字終結】
            print(f"【文章結束】，結束行處理作業。")
            return True, True
        elif cell_value == '\n':
            #【換行】
            print(f"【換行】，結束行中各欄處理作業。")
            return False, True
        elif not is_han_ji(cell_value):
            # 處理【標點符號】、【英數字元】、【其他字元】
            self._process_non_han_ji(cell_value)
            print(f"{cell_value} 非漢字，跳過處理。")
            return False, False
        else:
            self._process_han_ji(cell_value, cell, row, col)
            return False, False


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


def _process_sheet(sheet, config: ProcessConfig, processor: CellProcessor):
    """處理整個工作表"""

    # 處理所有的儲存格
    # active_cell = sheet.range((config.line_start_row, config.start_col))
    active_cell = sheet.range(f'{xw.utils.col_name(config.start_col)}{config.line_start_row}')
    active_cell.select()
    # 顯示【作用儲存格】位置
    # active_row = active_cell.row
    # active_col = active_cell.column
    # active_col_name = xw.utils.col_name(active_col)
    # print(f"作用儲存格：{active_col_name}{active_row}（{active_cell.row}, {active_cell.column}）")

    # 調整 row 值至【漢字】列（每 4 列為一組，漢字在第 3 列：5, 9, 13, ... ）
    is_eof = False
    for r in range(1, config.TOTAL_LINES + 1):
        if is_eof: break
        line_no = r
        print('-' * 60)
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
            print(f"儲存格：{xw.utils.col_name(col)}{row}（{row}, {col}）")
            is_eof, new_line = processor.process_cell(active_cell, row, col)
            if new_line: break
            if is_eof: break


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


def process(wb):
    """
    將 Excel 工作表中的漢字和標音整合輸出。
    """
    #--------------------------------------------------------------------------
    # 作業初始化
    #--------------------------------------------------------------------------
    logging_process_step("<----------- 作業開始！---------->")

    try:
        #--------------------------------------------------------------------------
        # 初始化 process config
        #--------------------------------------------------------------------------
        config = ProcessConfig(wb, hanji_piau_im_sheet='漢字注音')

        # 建立字庫工作表
        jin_kang_piau_im_ji_khoo, piau_im_ji_khoo, khuat_ji_piau_ji_khoo = _initialize_ji_khoo(
            wb=wb,
            new_jin_kang_piau_im_ji_khoo_sheet=False,
            new_piau_im_ji_khoo_sheet=False,
            new_khuat_ji_piau_sheet=False,
        )

        # 建立儲存格處理器
        processor = CellProcessor(
            config=config,
            jin_kang_piau_im_ji_khoo=jin_kang_piau_im_ji_khoo,
            piau_im_ji_khoo=piau_im_ji_khoo,
            khuat_ji_piau_ji_khoo=khuat_ji_piau_ji_khoo,
        )

        # --------------------------------------------------------------------------
        # 處理工作表
        # --------------------------------------------------------------------------
        piau_im_name_list  = [
            # '台語音標',
            # '方音符號',
            # '十五音',
            '雅俗通',
            '台羅拼音',
            '白話字',
            '閩拼調號',
            '閩拼調符',
        ]

        # 處理工作表
        for piau_im_name in piau_im_name_list:
            print('=' * 80)
            print(f"處理【標音方法】：{piau_im_name} ...")
            # 設定目前使用的標音方法
            processor.piau_im_huat = piau_im_name

            # 切換工作表
            sheet_name = f'漢字注音【{piau_im_name}】'
            # 使用【漢字注音】工作表複製新工作表
            try:
                if sheet_name not in [sheet.name for sheet in wb.sheets]:
                    # 複製工作表
                    source_sheet = wb.sheets['漢字注音']
                    new_sheet = source_sheet.copy(name=sheet_name, after=source_sheet)
                    print(f"✅ 已複製【漢字注音】工作表為 '{sheet_name}'")
                else:
                    print(f"⚠️ 工作表 '{sheet_name}' 已存在")

            except Exception as e:
                raise ValueError(f"無法找到或建立工作表 '{sheet_name}'：{e}")

            sheet = wb.sheets[sheet_name]
            sheet.activate()

            # 逐列處理
            _process_sheet(
                sheet=sheet,
                config=config,
                processor=processor,
            )

        # 寫回字庫到 Excel
        _save_ji_khoo_to_excel(
            wb=wb,
            jin_kang_piau_im_ji_khoo=jin_kang_piau_im_ji_khoo,
            piau_im_ji_khoo=piau_im_ji_khoo,
            khuat_ji_piau_ji_khoo=khuat_ji_piau_ji_khoo,
        )

        return EXIT_CODE_SUCCESS

    except Exception as e:
        logging.exception("自動為【漢字】查找【台語音標】作業，發生例外！")
        raise


def main():
    """主程式"""
    try:
        # 取得 Excel 活頁簿
        wb = None
        try:
            # 嘗試從 Excel 呼叫取得（RunPython）
            wb = xw.Book.caller()
        except:
            # 若失敗，則取得作用中的活頁簿
            try:
                wb = xw.apps.active.books.active
            except Exception as e:
                logging.error(f"無法找到作用中的 Excel 工作簿: {e}")
                return EXIT_CODE_NO_FILE

        if not wb:
            logging.error("無法取得 Excel 活頁簿")
            return EXIT_CODE_NO_FILE

        # 執行處理
        exit_code = process(wb)

        return exit_code

    except FileNotFoundError as fnf_error:
        logging_exception(msg="找不到指定的檔案！", error=fnf_error)
        return EXIT_CODE_NO_FILE
    except ValueError as val_error:
        logging_exception(msg="輸入資料有誤！", error=val_error)
        return EXIT_CODE_INVALID_INPUT
    except Exception as e:
        logging_exception(msg="處理過程中發生未知錯誤！", error=e)
        return EXIT_CODE_UNKNOWN_ERROR

if __name__ == "__main__":
    import sys
    sys.exit(main())