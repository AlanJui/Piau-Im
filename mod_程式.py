# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import logging
import os
import sys
from pathlib import Path
from typing import Tuple

# 載入第三方套件
import xlwings as xw
from dotenv import load_dotenv

# 載入自訂模組/函式
from mod_ca_ji_tian import HanJiTian
from mod_database import DatabaseManager
from mod_excel_access import delete_sheet_by_name, get_value_by_name, save_as_new_file
from mod_字庫 import JiKhooDict
from mod_帶調符音標 import is_han_ji, kam_si_u_tiau_hu, tng_im_piau, tng_tiau_ho
from mod_標音 import (  # 台語音標轉台語音標; 漢字標音物件
    PiauIm,
    ca_ji_tng_piau_im,
    convert_tl_with_tiau_hu_to_tlpa,
    convert_tlpa_to_tl,
    is_punctuation,
    split_hong_im_hu_ho,
    tlpa_tng_han_ji_piau_im,
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
class Program:
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
class ExcelCell:
    """儲存格處理器"""

    def __init__(
        self,
        program: Program,
        new_jin_kang_piau_im_ji_khoo_sheet: bool = False,
        new_piau_im_ji_khoo_sheet: bool = False,
        new_khuat_ji_piau_sheet: bool = False,
    ):
        self.program = program
        # 初始化資料庫管理器
        self.db_manager = DatabaseManager()
        self.db_manager.connect(program.db_name)
        #---------------------------------------------------------------------------
        # 初始化標音字庫
        #---------------------------------------------------------------------------
        # 人工標音字庫
        self.jin_kang_piau_im_ji_khoo_dict = self._initialize_ji_khoo(
            sheet_name='人工標音字庫',
            new_sheet=new_jin_kang_piau_im_ji_khoo_sheet,
        )
        # 標音字庫
        self.piau_im_ji_khoo_dict = self._initialize_ji_khoo(
            sheet_name='標音字庫',
            new_sheet=new_piau_im_ji_khoo_sheet,
        )
        # 缺字表
        self.khuat_ji_piau_ji_khoo_dict = self._initialize_ji_khoo(
            sheet_name='缺字表',
            new_sheet=new_khuat_ji_piau_sheet,
        )

    def _initialize_ji_khoo(
        self,
        sheet_name: str,
        new_sheet: bool,
    ) -> tuple[JiKhooDict]:
        """
        初始化字庫工作表

        :param sheet_name: 工作表名稱
        :param new_ji_khoo_sheet: 是否建立新的字庫工作表

        :return: JiKhooDict 物件
        """
        # 標音字庫
        if new_sheet:
            delete_sheet_by_name(wb=self.program.wb, sheet_name=sheet_name)
        ji_khoo_dict = JiKhooDict.create_ji_khoo_dict_from_sheet(
            wb=self.program.wb,
            sheet_name=sheet_name
        )

        return ji_khoo_dict

    def initialize_all_piau_im_ji_khoo_dict(
        self,
        new_jin_kang_piau_im_ji_khoo_sheet: bool,
        new_piau_im_ji_khoo_sheet: bool,
        new_khuat_ji_piau_sheet: bool,
    ) -> tuple[JiKhooDict, JiKhooDict, JiKhooDict]:
        """初始化字庫工作表"""
        # 人工標音字庫
        jin_kang_piau_im_ji_khoo_dict =  self._initialize_ji_khoo(
            sheet_name='人工標音字庫',
            new_sheet=new_jin_kang_piau_im_ji_khoo_sheet,
        )
        # 標音字庫
        piau_im_ji_khoo_dict = self._initialize_ji_khoo(
            sheet_name='標音字庫',
            new_sheet=new_piau_im_ji_khoo_sheet,
        )
        # 缺字表
        khuat_ji_piau_ji_khoo_dict = self._initialize_ji_khoo(
            sheet_name='缺字表',
            new_sheet=new_khuat_ji_piau_sheet,
        )

        self.jin_kang_piau_im_ji_khoo_dict = jin_kang_piau_im_ji_khoo_dict
        self.piau_im_ji_khoo_dict = piau_im_ji_khoo_dict
        self.khuat_ji_piau_ji_khoo_dict = khuat_ji_piau_ji_khoo_dict
        return jin_kang_piau_im_ji_khoo_dict, piau_im_ji_khoo_dict, khuat_ji_piau_ji_khoo_dict

    def save_all_piau_im_ji_khoo_dict(self):
        """將【字庫 Dict】存到 Excel 工作表中"""
        wb = self.program.wb
        self.jin_kang_piau_im_ji_khoo_dict.write_to_excel_sheet(wb=wb, sheet_name='人工標音字庫')
        self.piau_im_ji_khoo_dict.write_to_excel_sheet(wb=wb, sheet_name='標音字庫')
        self.khuat_ji_piau_ji_khoo_dict.write_to_excel_sheet(wb=wb, sheet_name='缺字表')

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
        self.piau_im_ji_khoo_dict.add_or_update_entry(
            han_ji=han_ji,
            tai_gi_im_piau=tai_gi_im_piau,
            kenn_ziann_im_piau=kenn_ziann_im_piau,
            coordinates=(row, col)
        )

    def update_entry_in_ji_khoo_dict(self, wb,
            ji_khoo_dict: JiKhooDict,
            han_ji: str, tai_gi_im_piau: str, kenn_ziann_im_piau: str, row: int, col: int):
        """更新字典內容"""
        ji_khoo_name = ji_khoo_dict.name if hasattr(ji_khoo_dict, 'name') else '標音字庫'
        target = f"（{row}, {col}）：{han_ji}【{tai_gi_im_piau}】"
        print(f"更新【{ji_khoo_name}】：{target}")
        # 取得該筆資料在【標音字庫】的 Row 號
        piau_im_ji_khoo_dict = ji_khoo_dict
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

    def _process_jin_kang_piau_im(self, jin_kang_piau_im: str, cell, row: int, col: int):
        """處理人工標音內容"""
        self.jin_kang_piau_im_ji_khoo_dict = self.jin_kang_piau_im_ji_khoo_dict
        # 預設未能依【人工標音】欄，找到對應的【台語音標】和【漢字標音】
        original_tai_gi_im_piau = cell.offset(-1, 0).value
        han_ji=cell.value
        sing_kong = False

        # 判斷【人工標音】是要【引用既有標音】還是【手動輸入標音】
        if  jin_kang_piau_im == '=':    # 引用既有的人工標音
            # 【人工標音】欄輸入為【=】，但【人工標音字庫】工作表查無結果，再自【標音字庫】工作表，嚐試查找【台語音標】
            tai_gi_im_piau = self.jin_kang_piau_im_ji_khoo_dict.get_tai_gi_im_piau_by_han_ji(han_ji=han_ji)
            if tai_gi_im_piau != '':
                row_no = self.jin_kang_piau_im_ji_khoo_dict.get_row_by_han_ji_and_tai_gi_im_piau(
                    han_ji=han_ji,
                    tai_gi_im_piau=tai_gi_im_piau)
                # 依指定之【標音方法】，將【台語音標】轉換成【漢字標音】
                han_ji_piau_im = tlpa_tng_han_ji_piau_im(
                    piau_im=self.program.piau_im,
                    piau_im_huat=self.program.piau_im_huat,
                    tai_gi_im_piau=tai_gi_im_piau
                )
                # 顯示處理訊息
                target = f"（{row}, {col}）：{han_ji}【{tai_gi_im_piau}】，【人工標音】：{jin_kang_piau_im}"
                print(f"{target}，在【人工標音字庫】工作表 row：{row_no}。")
                # 因【人工標音】欄為【=】，故而在【標音字庫】工作表中的紀錄，需自原有的【座標】欄位移除目前處理的座標除
                self.jin_kang_piau_im_ji_khoo_dict.update_entry_in_ji_khoo_dict(
                    wb=self.program.wb,
                    ji_khoo=self.program.han_ji_khoo_name,
                    han_ji=han_ji,
                    tai_gi_im_piau=tai_gi_im_piau,
                    kenn_ziann_im_piau='N/A',
                    row=row,
                    col=col
                )
                # 記錄到人工標音字庫
                self.jin_kang_piau_im_ji_khoo_dict.add_entry(
                    han_ji=han_ji,
                    tai_gi_im_piau=tai_gi_im_piau,
                    kenn_ziann_im_piau='N/A',
                    coordinates=(row, col)
                )
                sing_kong = True
            else:   # 若在【人工標音字庫】找不到【人工標音】對應的【台語音標】，則自【標音字庫】工作表查找
                cell.offset(-2, 0).value = ''  # 清空【人工標音】欄【=】
                tai_gi_im_piau = self.jin_kang_piau_im_ji_khoo_dict.get_tai_gi_im_piau_by_han_ji(han_ji=han_ji)
                if tai_gi_im_piau != '':
                    row_no = self.jin_kang_piau_im_ji_khoo_dict.get_row_by_han_ji_and_tai_gi_im_piau(
                        han_ji=han_ji,
                        tai_gi_im_piau=tai_gi_im_piau)
                    # 依指定之【標音方法】，將【台語音標】轉換成【漢字標音】
                    han_ji_piau_im = tlpa_tng_han_ji_piau_im(
                        piau_im=self.program.piau_im,
                        piau_im_huat=self.program.piau_im_huat,
                        tai_gi_im_piau=tai_gi_im_piau
                    )
                    # 記錄到標音字庫
                    self.jin_kang_piau_im_ji_khoo_dict.add_entry(
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
                    self.jin_kang_piau_im_ji_khoo_dict.khuat_ji_piau_ji_khoo.add_entry(
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
                piau_im=self.program.piau_im,
                piau_im_huat=self.program.piau_im_huat,
            )
            if tai_gi_im_piau != '' and han_ji_piau_im != '':
                # 自【標音字庫】工作表，移除目前處理的座標
                self.jin_kang_piau_im_ji_khoo_dict.update_entry_in_ji_khoo_dict(
                    wb=self.program.wb,
                    ji_khoo=self.program.han_ji_khoo_name,
                    han_ji=han_ji,
                    tai_gi_im_piau=original_tai_gi_im_piau,
                    kenn_ziann_im_piau='N/A',
                    row=row,
                    col=col
                )
                # 在【人工標音字庫】新增一筆紀錄
                self.jin_kang_piau_im_ji_khoo_dict.add_entry(
                    han_ji=han_ji,
                    tai_gi_im_piau=tai_gi_im_piau,
                    kenn_ziann_im_piau='N/A',
                    coordinates=(row, col)
                )
                row_no = self.jin_kang_piau_im_ji_khoo_dict.get_row_by_han_ji_and_tai_gi_im_piau(
                    han_ji=han_ji,
                    tai_gi_im_piau=tai_gi_im_piau)
                # 顯示處理訊息
                target = f"（【{han_ji}】[{tai_gi_im_piau}]／【{han_ji_piau_im}】，【人工標音】：{jin_kang_piau_im}"
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
            msg = f"【標點符號】"
        elif isinstance(cell_value, float) and cell_value.is_integer():
            msg = f"【英/數半形字元】（{int(cell_value)}）"
        else:
            msg = f"【非漢字之其餘字元】"

        print(f"【{cell_value}】：{msg}。")
        return

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
            han_ji_khoo=self.program.han_ji_khoo_name,
            piau_im=self.program.piau_im,
            piau_im_huat=self.program.piau_im_huat
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
            self.program.piau_im_ji_khoo_dict.add_entry(
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
        result = self.program.ji_tian.han_ji_ca_piau_im(
            han_ji=han_ji,
            ue_im_lui_piat=self.program.ue_im_lui_piat
        )

        # 查無此字
        if not result:
            self.program.khuat_ji_piau_ji_khoo_dict.add_entry(
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
            print(f"【{cell_value}】：【文章結束】結束行處理作業。")
            return True, True
        elif cell_value == '\n':
            #【換行】
            print(f"【換行】：結束行中各欄處理作業。")
            return False, True
        elif not is_han_ji(cell_value):
            # 處理【標點符號】、【英數字元】、【其他字元】
            self._process_non_han_ji(cell_value)
            return False, False
        else:
            self._process_han_ji(cell_value, cell, row, col)
            return False, False

    def insert_or_update_to_db(self, table_name: str, han_ji: str, tai_gi_im_piau: str, piau_im_huat: str) -> None:
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
        self.db_manager.execute(f"""
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
        row = self.db_manager.fetchone(
            f"SELECT 識別號 FROM {table_name} WHERE 漢字 = ? AND 台羅音標 = ?",
            (han_ji, tai_gi_im_piau)
        )

        siong_iong_too = 0.8 if piau_im_huat == "文讀音" else 0.6

        try:
            with self.db_manager.transaction():
                if row:
                    # 更新資料
                    from datetime import datetime
                    self.db_manager.execute(f"""
                    UPDATE {table_name}
                    SET 常用度 = ?, 更新時間 = ?
                    WHERE 識別號 = ?;
                    """, (siong_iong_too, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), row[0]))
                    print(f"  ✅ 已更新：{han_ji} - {tai_gi_im_piau}")
                else:
                    # 新增資料
                    self.db_manager.execute(f"""
                    INSERT INTO {table_name} (漢字, 台羅音標, 常用度, 摘要說明)
                    VALUES (?, ?, ?, NULL);
                    """, (han_ji, tai_gi_im_piau, siong_iong_too))
                    print(f"  ✅ 已新增：{han_ji} - {tai_gi_im_piau}")
        except Exception as e:
            print(f"  ❌ 資料庫操作失敗：{han_ji} - {tai_gi_im_piau}，錯誤：{e}")
            raise

    def khuat_ji_piau_poo_im_piau(self) -> int:
        """
        讀取 Excel 的【缺字表】工作表，並將資料回填至 SQLite 資料庫。
        """
        sheet_name = "缺字表"
        sheet = self.program.wb.sheets[sheet_name]
        piau_im_huat = self.program.piau_im_huat
        hue_im = self.program.ue_im_lui_piat
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
                self.insert_or_update_to_db(
                    table_name,
                    han_ji,
                    tai_gi_im_piau,
                    piau_im_huat
                )
                idx += 1

        logging_process_step(f"\n【缺字表】中的資料已成功回填至資料庫： {self.program.db_name} 的【{table_name}】資料表中。")
        return EXIT_CODE_SUCCESS

    def tiau_zing_piau_im_ji_khoo_dict(self, han_ji:str, tai_gi_im_piau:str, row:int, col:int) -> bool:
        """
        重整【標音字庫】查詢表：重整【標音字庫】工作表使用之 Dict
        依據【缺字表】工作表之【漢字】+【台語音標】資料，在【標音字庫】工作表【添增】此筆資料紀錄

        Args:
            han_ji (str): 漢字
            tai_gi_im_piau (str): 台語音標
            row (int): 儲存格列號
            col (int): 儲存格欄號
        Returns:
            bool: 是否在【標音字庫】找到該筆資料並移除
        """
        piau_im_ji_khoo_dict: JiKhooDict = self.program.piau_im_ji_khoo_dict

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

        return entry_found


# =========================================================================
# 作業處理函數
# =========================================================================

def process_sheet(sheet, program: Program, xls_cell: ExcelCell):
    """處理整個工作表"""

    # 處理所有的儲存格
    active_cell = sheet.range(f'{xw.utils.col_name(program.start_col)}{program.line_start_row}')
    active_cell.select()

    # 調整 row 值至【漢字】列（每 4 列為一組【列群】，漢字在第 3 列：5, 9, 13, ... ）
    is_eof = False
    for r in range(1, program.TOTAL_LINES + 1):
        if is_eof: break
        line_no = r
        print('=' * 80)
        print(f"處理第 {line_no} 行...")
        row = program.line_start_row + (r - 1) * program.ROWS_PER_LINE + program.han_ji_row_offset
        new_line = False
        for c in range(program.start_col, program.end_col + 1):
            if is_eof: break
            row = row
            col = c
            active_cell = sheet.range((row, col))
            active_cell.select()
            # 處理儲存格
            print('-' * 80)
            print(f"儲存格：{xw.utils.col_name(col)}{row}（{row}, {col}）")
            is_eof, new_line = xls_cell.process_cell(active_cell, row, col)
            if new_line: break  # noqa: E701
            if is_eof: break  # noqa: E701

# =========================================================================
# 本程式主要處理作業程序
# =========================================================================
def _process_sheet(sheet, program: Program, xls_cell: ExcelCell):
    """處理整個工作表"""

    # 處理所有的儲存格
    active_cell = sheet.range(f'{xw.utils.col_name(program.start_col)}{program.line_start_row}')
    active_cell.select()

    # 調整 row 值至【漢字】列（每 4 列為一組【列群】，漢字在第 3 列：5, 9, 13, ... ）
    is_eof = False
    for r in range(1, program.TOTAL_LINES + 1):
        if is_eof: break
        line_no = r
        print('=' * 80)
        print(f"處理第 {line_no} 行...")
        row = program.line_start_row + (r - 1) * program.ROWS_PER_LINE + program.han_ji_row_offset
        new_line = False
        for c in range(program.start_col, program.end_col + 1):
            if is_eof: break  # noqa: E701
            row = row
            col = c
            active_cell = sheet.range((row, col))
            active_cell.select()
            # 處理儲存格
            print('-' * 60)
            print(f"儲存格：{xw.utils.col_name(col)}{row}（{row}, {col}）")
            is_eof, new_line = xls_cell.process_cell(active_cell, row, col)
            if new_line: break  # noqa: E701
            if is_eof: break  # noqa: E701


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
        program = Program(wb, args, hanji_piau_im_sheet='漢字注音')

        # 建立儲存格處理器
        # xls_cell = ExcelCell(program=program)
        xls_cell = ExcelCell(
            program=program,
            new_jin_kang_piau_im_ji_khoo_sheet=False,
            new_piau_im_ji_khoo_sheet=False,
            new_khuat_ji_piau_sheet=False,
        )

        #--------------------------------------------------------------------------
        # 處理作業開始
        #--------------------------------------------------------------------------
        # 處理工作表
        sheet_name = '漢字注音'
        sheet = wb.sheets[sheet_name]
        sheet.activate()

        # 逐列處理
        _process_sheet(
            sheet=sheet,
            program=program,
            xls_cell=xls_cell,
        )

        # 寫回字庫到 Excel
        xls_cell.save_all_piau_im_ji_khoo_dict()

        #--------------------------------------------------------------------------
        # 處理作業結束
        #--------------------------------------------------------------------------
        print('\n')
        logging_process_step("<=========== 作業結束！==========>")
        return EXIT_CODE_SUCCESS

    except Exception as e:
        msg=f"處理作業，發生異常！ ==> error = {e}"
        logging.exception(msg)
        raise


# =========================================================================
# 主程式
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
        msg = "無法找到作用中的 Excel 工作簿！"
        logging_exception(msg=msg, error=e)
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
        logging_exception(msg=msg, error=e)
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
            logging_exc_error(msg="儲存檔案失敗！", error=None)
            return EXIT_CODE_SAVE_FAILURE    # 作業異當終止：無法儲存檔案
        else:
            logging_process_step(f"儲存檔案至路徑：{file_path}")
    except Exception as e:
        logging_exception(msg="儲存檔案失敗！", error=e)
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
  python a000.py          # 執行一般模式
  python a000.py -new     # 建立新的字庫工作表
  python a000.py -test    # 執行測試模式
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