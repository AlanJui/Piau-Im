# =========================================================================
# 程式功能摘要
# =========================================================================
# 用途：將漢字填入對應的儲存格
# 詳述：待加註讀音的漢字文置於 V3 儲存格。本程式將漢字逐字填入對應的儲存格：
# 【第一列】D5, E5, F5,... ,R5；
# 【第二列】D9, E9, F9,... ,R9；
# 【第三列】D13, E13, F13,... ,R13；
# 每個漢字佔一格，每格最多容納一個漢字。
# 漢字上方的儲存格（如：D4）為：【台語音標】欄，由【羅馬拼音字母】組成拼音。
# 漢字下方的儲存格（如：D6）為：【台語注音符號】欄，由【台語方音符號】組成注音。
# 漢字上上方的儲存格（如：D3）為：【人工標音】欄，可以只輸入【台語音標】；或
# 【台語音標】和【台語注音符號】皆輸入。

# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import argparse
import logging
import os
import re
from pathlib import Path
from typing import Optional, Tuple

# 載入第三方套件
import xlwings as xw
from dotenv import load_dotenv

# 載入自訂模組
from mod_ca_ji_tian import HanJiTian  # 新的查字典模組
from mod_excel_access import (
    clear_han_ji_kap_piau_im,
    delete_sheet_by_name,
    get_value_by_name,
    reset_cells_format_in_sheet,
    strip_cell,
)
from mod_file_access import save_as_new_file
from mod_字庫 import JiKhooDict
from mod_帶調符音標 import (
    fix_im_piau_spacing,
    is_han_ji,
    is_im_piau,
    read_text_with_han_ji,
    read_text_with_im_piau,
    tng_im_piau,
    tng_tiau_ho,
    tng_un_bu,
    zing_li_zuan_ku,
)
from mod_標音 import (
    PiauIm,
    ca_ji_kiat_ko_tng_piau_im,
    convert_tl_with_tiau_hu_to_tlpa,
    is_punctuation,
    split_hong_im_hu_ho,
    split_tai_gi_im_piau,
    tlpa_tng_han_ji_piau_im,
)

# =========================================================================
# 常數定義
# =========================================================================
EXIT_CODE_SUCCESS = 0
EXIT_CODE_NO_FILE = 1
EXIT_CODE_INVALID_INPUT = 2
EXIT_CODE_SAVE_FAILURE = 3
EXIT_CODE_PROCESS_FAILURE = 10
EXIT_CODE_UNKNOWN_ERROR = 99

# =========================================================================
# 載入環境變數
# =========================================================================
load_dotenv()

DB_HO_LOK_UE = os.getenv('DB_HO_LOK_UE', 'Ho_Lok_Ue.db')
DB_KONG_UN = os.getenv('DB_KONG_UN', 'Kong_Un.db')

# =========================================================================
# 設定日誌
# =========================================================================
from mod_logging import init_logging, logging_exc_error, logging_process_step

init_logging()


# =========================================================================
# 資料類別：儲存處理配置
# =========================================================================
class ProcessConfig:
    """處理配置資料類別"""

    def __init__(self, wb):
        # Excel 相關
        self.TOTAL_LINES = int(wb.names['每頁總列數'].refers_to_range.value)
        self.CHARS_PER_ROW = int(wb.names['每列總字數'].refers_to_range.value)
        self.ROWS_PER_LINE = 4
        self.start_row = 5
        self.start_col = 4
        self.end_row = self.start_row + (self.TOTAL_LINES * self.ROWS_PER_LINE)
        self.end_col = self.start_col + self.CHARS_PER_ROW

        # 標音相關
        self.han_ji_khoo_name = get_value_by_name(wb=wb, name='漢字庫')
        self.piau_im_huat = get_value_by_name(wb=wb, name='標音方法')


class CellProcessor:
    """儲存格處理器"""

    def __init__(
        self,
        ji_tian: HanJiTian,
        piau_im: PiauIm,
        piau_im_huat: str,
        ue_im_lui_piat: str,
        han_ji_khoo: str,
        jin_kang_piau_im_ji_khoo: JiKhooDict,
        piau_im_ji_khoo: JiKhooDict,
        khuat_ji_piau_ji_khoo: JiKhooDict,
    ):
        self.ji_tian = ji_tian
        self.piau_im = piau_im
        self.piau_im_huat = piau_im_huat
        self.ue_im_lui_piat = ue_im_lui_piat
        self.han_ji_khoo = han_ji_khoo
        self.jin_kang_piau_im_ji_khoo = jin_kang_piau_im_ji_khoo
        self.piau_im_ji_khoo = piau_im_ji_khoo
        self.khuat_ji_piau_ji_khoo = khuat_ji_piau_ji_khoo

    def process_cell(
        self,
        cell,
        row: int,
        col: int,
    ) -> Tuple[str, bool]:
        """
        處理單一儲存格

        Returns:
            (msg, is_eof): 處理訊息和是否到達文件結尾
        """
        # 初始化樣式
        self._reset_cell_style(cell)

        cell_value = cell.value

        jin_kang_piau_im = cell.offset(-2, 0).value  # 人工標音
        if jin_kang_piau_im and str(jin_kang_piau_im).strip() != "":
            msg, sing_kong =  self._process_jin_kang_piau_im(jin_kang_piau_im, cell, row, col)
            if sing_kong:
                return (msg, False)
            else:
                return (msg, False)

        # 檢查特殊字元
        if cell_value == 'φ':
            return "【文字終結】", True
        elif cell_value == '\n':
            return "【換行】", False
        elif not is_han_ji(cell_value):
            return self._process_non_han_ji(cell_value)
        else:
            return self._process_han_ji(cell_value, cell, row, col)

    def _reset_cell_style(self, cell):
        """重置儲存格樣式"""
        cell.font.color = (0, 0, 0)  # 黑色
        cell.color = None  # 無填滿

    def _jin_kang_piau_im_ca_han_ji_piau_im(self, han_ji: str, jin_kang_piau_im: str, piau_im: PiauIm, piau_im_huat: str):
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
        elif jin_kang_piau_im == '=':
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
            # 將人工輸入的【台語音標】，解構為【聲母】、【韻母】、【聲調】
            tai_gi_im_piau = convert_tl_with_tiau_hu_to_tlpa(jin_kang_piau_im)
            # 依指定之【標音方法】，將【台語音標】轉換成其所需之【漢字標音】
            han_ji_piau_im = tlpa_tng_han_ji_piau_im(
                piau_im=piau_im,
                piau_im_huat=piau_im_huat,
                tai_gi_im_piau=tai_gi_im_piau
            )

        return tai_gi_im_piau, han_ji_piau_im

    def _process_jin_kang_piau_im(self, jin_kang_piau_im, cell, row, col) -> Tuple[str, bool]:
        """處理人工標音內容"""
        # 預設未能依【人工標音】欄，找到對應的【台語音標】和【漢字標音】
        sing_kong = False
        han_ji=cell.value
        jin_kang_piau_im = str(jin_kang_piau_im).strip()

        tai_gi_im_piau, han_ji_piau_im = self._jin_kang_piau_im_ca_han_ji_piau_im(
            han_ji=han_ji,
            jin_kang_piau_im=str(jin_kang_piau_im),
            piau_im=self.piau_im,
            piau_im_huat=self.piau_im_huat,
        )
        if tai_gi_im_piau != '':
            # 依指定之【標音方法】，將【台語音標】轉換成其所需之【漢字標音】
            han_ji_piau_im = tlpa_tng_han_ji_piau_im(
                piau_im=self.piau_im,
                piau_im_huat=self.piau_im_huat,
                tai_gi_im_piau=tai_gi_im_piau
            )
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
        elif  tai_gi_im_piau == '' and jin_kang_piau_im == '=':
            # 【人工標音】欄輸入為【=】，但【人工標音字庫】工作表查無結果，再自【標音字庫】工作表，嚐試查找【台語音標】
            tai_gi_im_piau = self.piau_im_ji_khoo.get_tai_gi_im_piau_by_han_ji(han_ji=han_ji)
            if tai_gi_im_piau != '':
                # 依指定之【標音方法】，將【台語音標】轉換成其所需之【漢字標音】
                han_ji_piau_im = tlpa_tng_han_ji_piau_im(
                    piau_im=self.piau_im,
                    piau_im_huat=self.piau_im_huat,
                    tai_gi_im_piau=tai_gi_im_piau
                )
                cell.offset(-2, 0).value = ''  # 清空【人工標音】欄【=】
            else:
                # 若找不到【人工標音】對應的【台語音標】，則記錄到【缺字表】
                self.khuat_ji_piau_ji_khoo.add_entry(
                    han_ji=han_ji,
                    tai_gi_im_piau='N/A',
                    kenn_ziann_im_piau='N/A',
                    coordinates=(row, col)
                )

        # 寫入儲存格
        cell.offset(-1, 0).value = tai_gi_im_piau  # 上方儲存格：台語音標
        cell.offset(1, 0).value = han_ji_piau_im    # 下方儲存格：漢字標音

        # 依據【人工標音】欄，在【人工標音字庫】工作表找到相對應的【台語音標】和【漢字標音】
        msg = f"{cell.value}： [{tai_gi_im_piau}] /【{han_ji_piau_im}】（人工標音：【{jin_kang_piau_im}】）"
        sing_kong = True
        return msg, sing_kong

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

        # 轉換音標
        tai_gi_im_piau, han_ji_piau_im = self._convert_piau_im(result)

        # 寫入儲存格
        cell.offset(-1, 0).value = tai_gi_im_piau  # 上方儲存格：台語音標
        cell.offset(1, 0).value = han_ji_piau_im    # 下方儲存格：漢字標音

        # 記錄到標音字庫
        self.piau_im_ji_khoo.add_entry(
            han_ji=han_ji,
            tai_gi_im_piau=tai_gi_im_piau,
            kenn_ziann_im_piau='N/A',
            coordinates=(row, col)
        )

        return f"{han_ji}： [{tai_gi_im_piau}] /【{han_ji_piau_im}】", False

    def _convert_piau_im(self, result: list) -> Tuple[str, str]:
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
        tai_gi_im_piau, han_ji_piau_im = ca_ji_kiat_ko_tng_piau_im(
            result=result,
            han_ji_khoo=self.han_ji_khoo,
            piau_im=self.piau_im,
            piau_im_huat=self.piau_im_huat
        )
        return tai_gi_im_piau, han_ji_piau_im

# =========================================================================
# 主要處理函數
# =========================================================================
def extract_and_set_title(wb, file_path):
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            first_line = f.readline().strip()
            match = re.search(r"《(.*?)》", first_line)
            if match:
                title = match.group(1)
                wb.names['TITLE'].refers_to_range.value = title
                logging.info(f"✅ 已將文件標題《{title}》寫入 env 表 TITLE 名稱格。")
            else:
                logging.info("❕ 無《標題》可提取，未更新 TITLE。")
    except Exception as e:
        logging_exc_error("無法讀取或更新 TITLE 名稱。", error=e)


def fill_in_han_ji(wb, text_with_han_ji:list, sheet_name:str='漢字注音', start_row:int=5):
    sheet = wb.sheets[sheet_name]
    sheet.activate()
    sheet.range('A1').select()

    #------------------------------------------------------------------------------
    # 填入【漢字】
    #------------------------------------------------------------------------------
    row_han_ji = start_row      # 漢字位置
    start_col = 4   # 從D欄開始
    max_col = 18    # 最大可填入的欄位（R欄）

    col = start_col

    text = ""
    for han_ji_ku in text_with_han_ji:
        for han_ji in han_ji_ku:
            if col > max_col:
                # 超過欄位，換到下一組行
                row_han_ji += 4
                col = start_col

            text += han_ji
            sheet.cells(row_han_ji, col).value = han_ji
            sheet.cells(row_han_ji, col).select()  # 選取，畫面滾動
            col += 1  # 填入後右移一欄
            # 以下程式碼有假設：每組漢字之結尾，必有標點符號

        # 段落終結處：換下一段落
        if col > max_col:
            # 超過欄位，換到下一組行
            row_han_ji += 4
            col = start_col
        sheet.cells(row_han_ji, col).value = "=CHAR(10)"
        text += "\n"
        row_han_ji += 4
        col = start_col

    # 填入文章終止符號：φ
    # sheet["V3"].value = text
    sheet.cells(row_han_ji, col).value = "φ"
    print(f"已將文章之漢字純文字檔讀入，並填進【{sheet_name}】工作表！")

    return text_with_han_ji


def process(
    han_ji_file: str,
    wb,
    sheet_name: str = '漢字注音',
    cell: str = 'V3',
    ue_im_lui_piat: str = "白話音",
    han_ji_khoo: str = "河洛話",
    new_jin_kang_piau_im_ji_khoo_sheet: bool = True,
    new_piau_im_ji_khoo_sheet: bool = True,
    new_khuat_ji_piau_sheet: bool = True,
) -> int:
    """
    查詢漢字讀音並標注

    Args:
        han_ji_file: 漢字純文字檔路徑
        wb: Excel Workbook 物件
        sheet_name: 工作表名稱
        cell: 起始儲存格
        ue_im_lui_piat: 語音類別（白話音/文言音）
        han_ji_khoo: 漢字庫名稱
        new_jin_kang_piau_im_ji_khoo_sheet: 是否建立新的人工標音字庫表
        new_piau_im_ji_khoo_sheet: 是否建立新的標音字庫表
        new_khuat_ji_piau_sheet: 是否建立新的缺字表

    Returns:
        處理結果代碼
    """
    try:
        # 初始化配置
        config = ProcessConfig(wb)

        # 初始化字典物件
        db_name = DB_HO_LOK_UE if han_ji_khoo == '河洛話' else DB_KONG_UN
        ji_tian = HanJiTian(db_name)
        piau_im = PiauIm(han_ji_khoo=config.han_ji_khoo_name)

        #======================================================================
        # 填入【漢字】：讀取整篇文章之【漢字】純文字檔案；並填入【漢字注音】工作表。
        #======================================================================
        # 讀取漢字檔，並填入 Excel
        text_with_han_ji = read_text_with_han_ji(filename=han_ji_file)
        text_with_han_ji = fill_in_han_ji(wb,
                                          text_with_han_ji,
                                          sheet_name=sheet_name,
                                          start_row=config.start_row)

        # 建漢字清單：將 text_with_han_ji 整編為【漢字清單】
        han_ji_list = []
        for han_ji_ku in text_with_han_ji:
            for han_ji in han_ji_ku:
                han_ji_list.append(han_ji)
            # 段落終結處：換下一段落
            han_ji_list.append("\n")
        # 將漢字檔已讀取之內容，填入【漢字注音】工作表之【V3】儲存格
        wb.sheets['漢字注音'].range('V3').value = ''.join(han_ji_list)

        #======================================================================
        # 建立字庫工作表
        #======================================================================
        jin_kang_piau_im_ji_khoo, piau_im_ji_khoo, khuat_ji_piau_ji_khoo = _initialize_ji_khoo(
            wb=wb,
            new_jin_kang_piau_im_ji_khoo_sheet=new_jin_kang_piau_im_ji_khoo_sheet,
            new_piau_im_ji_khoo_sheet=new_piau_im_ji_khoo_sheet,
            new_khuat_ji_piau_sheet=new_khuat_ji_piau_sheet,
        )

        # 建立儲存格處理器
        processor = CellProcessor(
            ji_tian=ji_tian,
            piau_im=piau_im,
            piau_im_huat=config.piau_im_huat,
            ue_im_lui_piat=ue_im_lui_piat,
            han_ji_khoo=han_ji_khoo,
            jin_kang_piau_im_ji_khoo=jin_kang_piau_im_ji_khoo,
            piau_im_ji_khoo=piau_im_ji_khoo,
            khuat_ji_piau_ji_khoo=khuat_ji_piau_ji_khoo,
        )

        # 處理工作表
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

        logging_process_step("已完成【台語音標】和【漢字標音】標注工作。")
        return EXIT_CODE_SUCCESS

    except Exception as e:
        logging.exception("自動為【漢字】查找【台語音標】作業，發生例外！")
        raise


def _initialize_ji_khoo(
    wb,
    new_jin_kang_piau_im_ji_khoo_sheet: bool,
    new_piau_im_ji_khoo_sheet: bool,
    new_khuat_ji_piau_sheet: bool,
) -> Tuple[JiKhooDict, JiKhooDict]:
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

    EOF = False
    line = 1
    empty_cell_count = 0

    for row in range(config.start_row, config.end_row, config.ROWS_PER_LINE):
        # 設定作用儲存格為列首
        sheet.range((row, 1)).select()

        # 逐欄處理
        for col in range(config.start_col, config.end_col):
            cell = sheet.range((row, col))
            # 設定作用儲存格為目前儲存格
            cell.select()

            # 處理儲存格
            msg, is_eof = processor.process_cell(cell, row, col)

            # 處理空白儲存格計數
            if msg == "【空白】":
                empty_cell_count += 1
                if empty_cell_count >= 2:
                    EOF = True
            else:
                empty_cell_count = 0

            # 顯示處理進度
            col_name = xw.utils.col_name(col)
            print(f"【{col_name}{row}】({row}, {col}) = {msg}")

            # 檢查是否結束
            if is_eof or EOF or msg == "【換行】":
                EOF = is_eof or EOF
                break

        # 檢查是否到達結尾
        if EOF or line > config.TOTAL_LINES:
            break

        # 換行顯示
        if col == config.end_col - 1:
            print('\n')

        line += 1


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


# =========================================================================
# 主程式
# =========================================================================
def main():

    # 使用 argparse 處理命令列參數
    parser = argparse.ArgumentParser(description="根據漢字檔產生台語音標與漢字標音")
    parser.add_argument("han_ji_file", nargs="?", default="_tmp_p1_han_ji.txt", help="漢字純文字檔路徑")
    parser.add_argument("ping_im_file", nargs="?", default="", help="標音檔（可選）")
    parser.add_argument("--reset_wb", action="store_true", help="重置工作表初始狀態")
    parser.add_argument("--peh_ue", action="store_true", help="將語音類型設定為白話音")
    # parser.add_argument("--tiau_hu", action="store_false", help="使用【聲調符號】（不帶調號）的音標（TLPA 拼音）")
    # use_tiau_hu = args.tiau_hu
    # 預設為使用聲調號（tiau_ho = True），若加上 --tiau_hu，則關閉聲調號
    parser.add_argument("--tiau_hu", action="store_false", dest="tiau_ho", help="TLPA音標改【聲調符號】（不帶調號數值）")

    args = parser.parse_args()

    han_ji_file = args.han_ji_file
    reset_wb = args.reset_wb
    ping_im_file = args.ping_im_file
    use_peh_ue = args.peh_ue
    use_tiau_ho = args.tiau_ho

    """主程式 - 從 Excel 呼叫或直接執行"""
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

        # 若使用者指定 --peh_ue 參數，則將【語音類型】設定為【白話音】
        if use_peh_ue:
            try:
                # 設定【語音類型】為【白話音】--peh_ue
                wb.names['語音類型'].refers_to_range.value = '白話音'
                print("已將【env】工作表，之【語音類型】欄，設定為：【白話音】。")
            except Exception as e:
                logging_exc_error("無法設定【env】工作表，之【語音類型】欄，為【白話音】。", error=e)
                return EXIT_CODE_PROCESS_FAILURE

        # 若使用者指定 --tiau_hu 參數，則將【標音方法】設定欄之【閩拼調號】，變更成【閩拼調符】
        if not use_tiau_ho:
            try:
                piau_im_huat = wb.names['標音方法'].refers_to_range.value
                if piau_im_huat == "閩拼調號":
                    # 若原本為【閩拼調號】，則將其變更為【閩拼調符】
                    wb.names['標音方法'].refers_to_range.value = '閩拼調符'
                    print("已將【env】工作表，之【標音方法】欄，設定為：【閩拼調符】。")
            except Exception as e:
                logging_exc_error("無法設定【env】工作表，之【標音方法】欄，為【閩拼調符】。", error=e)
                return EXIT_CODE_PROCESS_FAILURE

        #======================================================================
        # 重置工作表
        #======================================================================
        print("清除儲存格內容...")
        clear_han_ji_kap_piau_im(wb)
        logging.info("儲存格內容清除完畢")

        if reset_wb:
            print("重設儲存格之格式...")
            reset_cells_format_in_sheet(wb)
            logging.info("儲存格格式重設完畢")

        #======================================================================
        # 執行處理
        #======================================================================
        ue_im_lui_piat = get_value_by_name(wb=wb, name='語音類型')
        han_ji_khoo = get_value_by_name(wb=wb, name='漢字庫')
        sheet_name = '漢字注音'
        wb.sheets[sheet_name].activate()
        exit_code = process(
            han_ji_file=han_ji_file,
            wb=wb,
            sheet_name=sheet_name,
            ue_im_lui_piat=ue_im_lui_piat,
            han_ji_khoo=han_ji_khoo,
            new_khuat_ji_piau_sheet=False,
            new_piau_im_ji_khoo_sheet=False,
        )
        if exit_code != EXIT_CODE_SUCCESS:
            logging_process_step("在【漢字注音】工作表，查找【台語音標】與【漢字標音】轉換作業已結束。")
            return exit_code

        #======================================================================
        # 儲存檔案
        #======================================================================
        # 試圖以【漢字】純文字檔案，設定【標題】名稱
        try:
            extract_and_set_title(wb, han_ji_file)
        except Exception as e:
            logging_exc_error("無法讀取或更新 TITLE 名稱。", error=e)

        # 存檔
        try:
            # 要求畫面回到【漢字注音】工作表
            wb.sheets['漢字注音'].activate()
            # 儲存檔案
            file_path = save_as_new_file(wb=wb)
            if not file_path:
                logging.error("儲存檔案失敗！")
                return EXIT_CODE_SAVE_FAILURE    # 作業異當終止：無法儲存檔案
            else:
                logging_process_step(f"儲存檔案至路徑：{file_path}")
        except Exception as e:
            logging_exc_error(msg="儲存檔案失敗！", error=e)
            return EXIT_CODE_SAVE_FAILURE    # 作業異當終止：無法儲存檔案

        return EXIT_CODE_SUCCESS    # 作業成功結束

    except Exception as e:
        logging.exception("程式執行失敗")
        return EXIT_CODE_UNKNOWN_ERROR


def test_han_ji_tian():
    """測試 HanJiTian 類別"""
    print("=" * 70)
    print("測試 HanJiTian 查詢功能")
    print("=" * 70)

    try:
        # 初始化字典
        ji_tian = HanJiTian(DB_HO_LOK_UE)

        # 測試查詢
        test_chars = ["東", "西", "南", "北", "中"]

        for han_ji in test_chars:
            print(f"\n查詢漢字：{han_ji}")
            result = ji_tian.han_ji_ca_piau_im(han_ji, ue_im_lui_piat="白話音")

            if result:
                for item in result:
                    print(f"  台語音標：{item['台語音標']}, 常用度：{item.get('常用度', 'N/A')}, 說明：{item.get('摘要說明', 'N/A')}")
            else:
                print(f"  查無資料")

        print("\n" + "=" * 70)
        print("測試完成")
        print("=" * 70)

    except Exception as e:
        print(f"測試失敗：{e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    import sys

    # 檢查是否有命令列參數
    if len(sys.argv) > 1 and sys.argv[1] == "test":
        # 執行測試
        test_han_ji_tian()
    else:
        # 從 Excel 呼叫
        main()
