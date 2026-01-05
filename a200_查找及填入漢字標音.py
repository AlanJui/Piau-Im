# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import logging
import os
from pathlib import Path
from typing import Optional, Tuple

# 載入第三方套件
import xlwings as xw
from dotenv import load_dotenv

# 載入自訂模組
from mod_BP_tng_huan_ping_im import convert_TLPA_to_BP
from mod_ca_ji_tian import HanJiTian  # 新的查字典模組
from mod_excel_access import delete_sheet_by_name, get_value_by_name
from mod_字庫 import JiKhooDict
from mod_帶調符音標 import is_han_ji, kam_si_u_tiau_hu
from mod_標音 import (
    PiauIm,
    ca_ji_kiat_ko_tng_piau_im,
    ca_ji_tng_piau_im,
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

    def __init__(self, wb, new_piau_im_sheets: bool = False, hanji_piau_im_sheet: str = '漢字注音'):
        self.wb = wb
        self.new_piau_im_sheets = new_piau_im_sheets
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

        if '〔' in jin_kang_piau_im and '〕' in jin_kang_piau_im:   # 〔台語音標/台羅拼音〕
            # 將人工輸入的〔台語音標〕轉換成【方音符號】
            im_piau = jin_kang_piau_im.split('〔')[1].split('〕')[0].strip()
            tai_gi_im_piau = convert_tl_with_tiau_hu_to_tlpa(im_piau)
            # 依使用者指定之【標音方法】，將【台語音標】轉換成其所需之【漢字標音】
            han_ji_piau_im = tlpa_tng_han_ji_piau_im(
                piau_im=piau_im,
                piau_im_huat=piau_im_huat,
                tai_gi_im_piau=tai_gi_im_piau
            )
        elif '【' in jin_kang_piau_im and '】' in jin_kang_piau_im:  # 【方音符號/注音符號】
            # 將人工輸入的【方音符號】轉換成【台語音標】
            han_ji_piau_im = jin_kang_piau_im.split('【')[1].split('】')[0].strip()
            siann, un, tiau = split_hong_im_hu_ho(han_ji_piau_im)
            # 依使用者指定之【標音方法】，將【台語音標】轉換成其所需之【漢字標音】
            tai_gi_im_piau = piau_im.hong_im_tng_tai_gi_im_piau(
                siann=siann,
                un=un,
                tiau=tiau)['台語音標']
        else:    # 直接輸入【人工標音】
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

    def new_entry_in_ji_khoo_dict(self,
            han_ji: str, tai_gi_im_piau: str, kenn_ziann_im_piau: str, row: int, col: int):
        """更新字典內容"""
        self.piau_im_ji_khoo.add_or_update_entry(
            han_ji=han_ji,
            tai_gi_im_piau=tai_gi_im_piau,
            kenn_ziann_im_piau=kenn_ziann_im_piau,
            coordinates=(row, col)
        )

    def update_entry_in_ji_khoo_dict(self, wb,
            ji_khoo: JiKhooDict,
            han_ji: str, tai_gi_im_piau: str, kenn_ziann_im_piau: str, row: int, col: int):
        """更新字典內容"""
        # ji_khoo_name = '標音字庫'
        ji_khoo_name = ji_khoo.name if hasattr(ji_khoo, 'name') else '標音字庫'
        target = f"（{row}, {col}）：{han_ji}【{tai_gi_im_piau}】"
        print(f"更新【{ji_khoo_name}】：{target}")
        # 取得該筆資料在【標音字庫】的 Row 號
        piau_im_ji_khoo_dict = ji_khoo
        row_no = piau_im_ji_khoo_dict.get_row_by_han_ji_and_tai_gi_im_piau(
            han_ji=han_ji,
            tai_gi_im_piau=tai_gi_im_piau
        )
        print(f"{target}落在【標音字庫】工作表的列號：{row_no}")
        # 依【漢字】與【台語音標】取得在【標音字庫】工作表中的【座標】清單
        coord_list = piau_im_ji_khoo_dict.get_coordinates_by_han_ji_and_tai_gi_im_piau(
            han_ji=han_ji,
            tai_gi_im_piau=tai_gi_im_piau
        )
        print(f"{target}對映的座標清單：{coord_list}")
        # 自座標清單中，移除目前處理的座標
        coord_to_remove = (row, col)
        if coord_to_remove in coord_list:
            # (row, col)座標落在座標清單中
            print(f"座標 {coord_to_remove} 有在座標清單之中。")
            # 自座標清單中移除該座標
            piau_im_ji_khoo_dict.remove_coordinate(
                han_ji=han_ji,
                tai_gi_im_piau=tai_gi_im_piau,
                coordinate=coord_to_remove
            )
            print(f"{target}已自座標清單中移除。")
            # 儲存回 Excel
            print(f"將更新後的【標音字庫】寫回 Excel 工作表...")
            piau_im_ji_khoo_dict.write_to_excel_sheet(
                wb=wb,
                sheet_name='標音字庫'
            )
        else:
            print(f"座標 {coord_to_remove} 不在座標清單之中。")
        return

    def _process_jin_kang_piau_im(self, jin_kang_piau_im, cell, row, col):
        """處理人工標音內容"""
        # 預設未能依【人工標音】欄，找到對應的【台語音標】和【漢字標音】
        original_tai_gi_im_piau = cell.offset(-1, 0).value
        han_ji=cell.value
        sing_kong = False

        # 判斷【人工標音】是要【引用既有標音】還是【手動輸入標音】
        if  jin_kang_piau_im == '=':    # 引用既有的人工標音
            # 【人工標音】欄輸入為【=】，但【人工標音字庫】工作表查無結果，再自【標音字庫】工作表，嚐試查找【台語音標】
            tai_gi_im_piau = self.jin_kang_piau_im_ji_khoo.get_tai_gi_im_piau_by_han_ji(han_ji=han_ji)
            if tai_gi_im_piau != '':
                row_no = self.jin_kang_piau_im_ji_khoo.get_row_by_han_ji_and_tai_gi_im_piau(
                    han_ji=han_ji,
                    tai_gi_im_piau=tai_gi_im_piau)
                # 依指定之【標音方法】，將【台語音標】轉換成【漢字標音】
                han_ji_piau_im = tlpa_tng_han_ji_piau_im(
                    piau_im=self.piau_im,
                    piau_im_huat=self.piau_im_huat,
                    tai_gi_im_piau=tai_gi_im_piau
                )
                # 顯示處理訊息
                target = f"（{row}, {col}）：{han_ji}【{tai_gi_im_piau}】，【人工標音】：{jin_kang_piau_im}"
                print(f"{target}，在【人工標音字庫】工作表 row：{row_no}。")
                # 因【人工標音】欄為【=】，故而在【標音字庫】工作表中的紀錄，需自原有的【座標】欄位移除目前處理的座標除
                self.update_entry_in_ji_khoo_dict(
                    wb=self.config.wb,
                    ji_khoo=self.piau_im_ji_khoo,
                    han_ji=han_ji,
                    tai_gi_im_piau=tai_gi_im_piau,
                    kenn_ziann_im_piau='N/A',
                    row=row,
                    col=col
                )
                # 記錄到人工標音字庫
                self.jin_kang_piau_im_ji_khoo.add_entry(
                    han_ji=han_ji,
                    tai_gi_im_piau=tai_gi_im_piau,
                    kenn_ziann_im_piau='N/A',
                    coordinates=(row, col)
                )
                sing_kong = True
            else:   # 若在【人工標音字庫】找不到【人工標音】對應的【台語音標】，則自【標音字庫】工作表查找
                cell.offset(-2, 0).value = ''  # 清空【人工標音】欄【=】
                tai_gi_im_piau = self.piau_im_ji_khoo.get_tai_gi_im_piau_by_han_ji(han_ji=han_ji)
                if tai_gi_im_piau != '':
                    row_no = self.piau_im_ji_khoo.get_row_by_han_ji_and_tai_gi_im_piau(
                        han_ji=han_ji,
                        tai_gi_im_piau=tai_gi_im_piau)
                    # 依指定之【標音方法】，將【台語音標】轉換成【漢字標音】
                    han_ji_piau_im = tlpa_tng_han_ji_piau_im(
                        piau_im=self.piau_im,
                        piau_im_huat=self.piau_im_huat,
                        tai_gi_im_piau=tai_gi_im_piau
                    )
                    # 記錄到標音字庫
                    self.piau_im_ji_khoo.add_entry(
                        han_ji=han_ji,
                        tai_gi_im_piau=tai_gi_im_piau,
                        kenn_ziann_im_piau='N/A',
                        coordinates=(row, col)
                    )
                    # 顯示處理訊息
                    target = f"（{row}, {col}）：{han_ji}【{tai_gi_im_piau}】，【人工標音】：{jin_kang_piau_im}"
                    print(f"{target}的【人工標音】，在【人工標音字庫】找不到，改用【標音字庫】（row：{row_no}）中的【台語音標】替代。")
                    sing_kong = True
                else:
                    # 若找不到【人工標音】對應的【台語音標】，則記錄到【缺字表】
                    self.khuat_ji_piau_ji_khoo.add_entry(
                        han_ji=han_ji,
                        tai_gi_im_piau='N/A',
                        kenn_ziann_im_piau='N/A',
                        coordinates=(row, col)
                    )
        else:   # 手動輸入人工標音
            # 自【人工標音】儲存格，取出【人工標音】
            tai_gi_im_piau, han_ji_piau_im = self._cu_jin_kang_piau_im(
                han_ji=han_ji,
                jin_kang_piau_im=str(jin_kang_piau_im),
                piau_im=self.piau_im,
                piau_im_huat=self.piau_im_huat,
            )
            if tai_gi_im_piau != '' and han_ji_piau_im != '':
                # 自【標音字庫】工作表，移除目前處理的座標
                self.update_entry_in_ji_khoo_dict(
                    wb=self.config.wb,
                    ji_khoo=self.piau_im_ji_khoo,
                    han_ji=han_ji,
                    tai_gi_im_piau=original_tai_gi_im_piau,
                    kenn_ziann_im_piau='N/A',
                    row=row,
                    col=col
                )
                # 在【人工標音字庫】新增一筆紀錄
                self.jin_kang_piau_im_ji_khoo.add_entry(
                    han_ji=han_ji,
                    tai_gi_im_piau=tai_gi_im_piau,
                    kenn_ziann_im_piau='N/A',
                    coordinates=(row, col)
                )
                row_no = self.jin_kang_piau_im_ji_khoo.get_row_by_han_ji_and_tai_gi_im_piau(
                    han_ji=han_ji,
                    tai_gi_im_piau=tai_gi_im_piau)
                # 顯示處理訊息
                target = f"（{row}, {col}）：{han_ji}【{tai_gi_im_piau}】，【人工標音】：{jin_kang_piau_im}"
                print(f"{target}，已記錄到【人工標音字庫】工作表的 row：{row_no}）。")
                sing_kong = True

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
        #-------------------------------------------
        # 顯示漢字庫查找結果的單一讀音選項
        #-------------------------------------------
        def _process_one_entry(cell, entry):
            # 轉換音標
            tai_gi_im_piau, han_ji_piau_im = self._convert_piau_im(entry)

            # 寫入儲存格
            cell.offset(-1, 0).value = tai_gi_im_piau  # 上方儲存格：台語音標
            cell.offset(1, 0).value = han_ji_piau_im    # 下方儲存格：漢字標音

            # 在【標音字庫】新增一筆紀錄
            self.piau_im_ji_khoo.add_entry(
                han_ji=han_ji,
                tai_gi_im_piau=tai_gi_im_piau,
                kenn_ziann_im_piau='N/A',
                coordinates=(row, col)
            )

            # 顯示處理進度
            han_ji_thok_im = f" [{tai_gi_im_piau}] /【{han_ji_piau_im}】"

            # 結束處理
            return han_ji_thok_im

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

        # 顯示所有讀音選項
        # excel_address = f"{xw.utils.col_name(col)}{row}"
        # print(f"漢字儲存格：{excel_address}（{row}, {col}）：【{han_ji}】有 {len(result)} 個讀音...")
        # for idx, entry in enumerate(cell, result):
        #     han_ji_thok_im = _process_one_entry(cell, entry)
        #     print(f"{idx + 1}. 【{han_ji}】：{han_ji_thok_im}")

        # 預設只處理第一個讀音選項
        excel_address = f"{xw.utils.col_name(col)}{row}"
        print(f"漢字儲存格：{excel_address}（{row}, {col}）的讀音為...")
        han_ji_thok_im = _process_one_entry(cell, result[0])
        print(f"【{han_ji}】：{han_ji_thok_im}")


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

        # 若【人工標音】欄位有值，且【漢字】欄位有【漢字】，則以【人工標音】求取【台語音標】及【漢字標音】
        jin_kang_piau_im = cell.offset(-2, 0).value  # 人工標音
        if jin_kang_piau_im and is_han_ji(cell_value):
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
# 主要處理函數
# =========================================================================
def process(wb, new_piau_im_sheets: bool = False) -> int:
    """
    查詢漢字讀音並標注

    Args:
        wb: Excel Workbook 物件

    Returns:
        處理結果代碼
    """
    #--------------------------------------------------------------------------
    # 作業初始化
    #--------------------------------------------------------------------------
    logging_process_step("<----------- 作業開始！---------->")

    try:
        #--------------------------------------------------------------------------
        # 初始化 process config
        #--------------------------------------------------------------------------
        config = ProcessConfig(wb, new_piau_im_sheets=new_piau_im_sheets, hanji_piau_im_sheet='漢字注音')

        # 建立字庫工作表
        create = new_piau_im_sheets
        if create:
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

        # 處理工作表
        sheet_name = '漢字注音'
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
            jin_kang_piau_im_ji_khoo=jin_kang_piau_im_ji_khoo_dict,
            piau_im_ji_khoo=piau_im_ji_khoo_dict,
            khuat_ji_piau_ji_khoo=khuat_ji_piau_ji_khoo_dict,
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

    # 缺字表
    khuat_ji_piau_name = '缺字表'
    if new_khuat_ji_piau_sheet:
        delete_sheet_by_name(wb=wb, sheet_name=khuat_ji_piau_name)
    khuat_ji_piau_ji_khoo = JiKhooDict.create_ji_khoo_dict_from_sheet(
        wb=wb,
        sheet_name=khuat_ji_piau_name
    )

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

    return jin_kang_piau_im_ji_khoo, piau_im_ji_khoo, khuat_ji_piau_ji_khoo


def _process_sheet(sheet, config: ProcessConfig, processor: CellProcessor):
    """處理整個工作表"""
    # 處理所有的儲存格
    active_cell = sheet.range(f'{xw.utils.col_name(config.start_col)}{config.line_start_row}')
    active_cell.select()

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


# =========================================================================
# 主程式
# =========================================================================
def main(new_piau_im_sheets: bool = False) -> int:
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

        # 執行處理
        exit_code = process(wb=wb, new_piau_im_sheets=new_piau_im_sheets)

        return exit_code

    except Exception as e:
        logging.exception("程式執行失敗")
        return EXIT_CODE_UNKNOWN_ERROR


def test_01():
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
    import argparse
    import sys

    # 解析命令行參數
    parser = argparse.ArgumentParser(
        description='依【漢字】查找【台語音標】並轉換成【漢字標音】',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
使用範例：
  python a200_查找及填入漢字標音.py          # 執行一般模式
  python a200_查找及填入漢字標音.py -new     # 建立新的字庫工作表
  python a200_查找及填入漢字標音.py -test    # 執行測試模式
'''
        )
    parser.add_argument(
        '-new',
        action='store_true',
        help='建立新的字庫工作表',
    )
    parser.add_argument(
        '-test',
        action='store_true',
        help='執行測試模式',
    )
    args = parser.parse_args()
    new_piau_im_sheets = args.new
    test_mode = args.test

    if test_mode:
        # 執行測試
        sys.exit(test_01())
    else:
        # 從 Excel 呼叫
        sys.exit(main(new_piau_im_sheets))