import logging
import os
import sys
from typing import Tuple

import xlwings as xw
from dotenv import load_dotenv

from mod_ca_ji_tian import HanJiTian
from mod_excel_access import delete_sheet_by_name, ensure_sheet_exists
from mod_logging import logging_exception
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

class JiKhooDict:
    def __init__(self):
        self.ji_khoo_dict = {}

    def items(self):
        for han_ji, entries in self.ji_khoo_dict.items():
            for entry in entries:
                yield (han_ji, entry)

    def add_entry(self, han_ji: str, tai_gi_im_piau: str, kenn_ziann_im_piau: str, coordinates: tuple):
        if not tai_gi_im_piau:
            tai_gi_im_piau = "N/A"
        if not kenn_ziann_im_piau:
            kenn_ziann_im_piau = "N/A"

        entry = {
            "tai_gi_im_piau": tai_gi_im_piau,
            "kenn_ziann_im_piau": kenn_ziann_im_piau,
            "coordinates": [coordinates]
        }

        if han_ji not in self.ji_khoo_dict:
            self.ji_khoo_dict[han_ji] = [entry]
        else:
            for existing in self.ji_khoo_dict[han_ji]:
                if existing["tai_gi_im_piau"] == tai_gi_im_piau:
                    if coordinates not in existing["coordinates"]:
                        existing["coordinates"].append(coordinates)
                    return
            self.ji_khoo_dict[han_ji].append(entry)

    def update_entry(self, han_ji: str, tai_gi_im_piau: str, kenn_ziann_im_piau: str, coordinates: tuple):
        if han_ji not in self.ji_khoo_dict:
            raise ValueError(f"漢字 '{han_ji}' 不存在，請先使用 add_entry 方法新增資料。")

        for existing in self.ji_khoo_dict[han_ji]:
            if existing["tai_gi_im_piau"] == tai_gi_im_piau:
                if kenn_ziann_im_piau:
                    existing["kenn_ziann_im_piau"] = kenn_ziann_im_piau
                if coordinates not in existing["coordinates"]:
                    existing["coordinates"].append(coordinates)
                return

        self.add_entry(han_ji, tai_gi_im_piau, kenn_ziann_im_piau, coordinates)

    def add_or_update_entry(self, han_ji, tai_gi_im_piau, kenn_ziann_im_piau, coordinates):
        self.add_entry(han_ji, tai_gi_im_piau, kenn_ziann_im_piau, coordinates)

    def get_entry(self, han_ji: str):
        if han_ji in self.ji_khoo_dict:
            return self.ji_khoo_dict[han_ji]
        else:
            raise ValueError(f"漢字 '{han_ji}' 不存在於字典中。")

    def get_value_by_key(self, han_ji: str, tai_gi_im_piau: str, key: str):
        if han_ji in self.ji_khoo_dict:
            for entry in self.ji_khoo_dict[han_ji]:
                if entry["tai_gi_im_piau"] == tai_gi_im_piau:
                    return entry.get(key)
            raise ValueError(f"漢字 '{han_ji}' 中找不到音標 '{tai_gi_im_piau}' 對應的欄位 '{key}'。")
        else:
            raise ValueError(f"漢字 '{han_ji}' 不存在於字典中。")


    def update_value_by_key(self, han_ji: str, tai_gi_im_piau: str, key: str, value):
        if han_ji in self.ji_khoo_dict:
            for entry in self.ji_khoo_dict[han_ji]:
                if entry["tai_gi_im_piau"] == tai_gi_im_piau:
                    if key in entry:
                        entry[key] = value
                        return
                    else:
                        raise ValueError(f"欄位名稱 '{key}' 無效。")
            raise ValueError(f"找不到對應音標 '{tai_gi_im_piau}' 的資料。")
        else:
            raise ValueError(f"漢字 '{han_ji}' 不存在於字典中。")

    def get_tai_gi_im_piau_by_han_ji(self, han_ji: str) -> str:
        """
        根據漢字查詢台語音標
        若該漢字有多個音標，返回第一個
        若查無結果，返回空字串

        Args:
            han_ji: 要查詢的漢字

        Returns:
            str: 台語音標，若無則返回空字串
        """
        if han_ji in self.ji_khoo_dict:
            entries = self.ji_khoo_dict[han_ji]
            if entries and len(entries) > 0:
                # 返回第一個音標
                tai_gi_im_piau = entries[0].get("tai_gi_im_piau", "")
                # 若該音標為 N/A 則返回空字串
                if tai_gi_im_piau == "N/A":
                    return ""
                return tai_gi_im_piau
        return ""

    def update_kau_ziang_im_piau(self, han_ji: str, tai_gi_im_piau: str, kenn_ziann_im_piau: str, coordinates: tuple):
        """
        將人工標音或校正音標更新至字典。
        如果該漢字、音標已存在則更新校正音標與座標。
        若尚未記錄該音標，則新增一筆。
        """
        if han_ji in self.ji_khoo_dict:
            for entry in self.ji_khoo_dict[han_ji]:
                if entry["tai_gi_im_piau"] == tai_gi_im_piau:
                    entry["kenn_ziann_im_piau"] = kenn_ziann_im_piau
                    if coordinates not in entry["coordinates"]:
                        entry["coordinates"].append(coordinates)
                    return
        # 若找不到，則新增新項目
        self.add_entry(han_ji, tai_gi_im_piau, kenn_ziann_im_piau, coordinates)

    def update_by_piau_im_ji_khoo(self, wb, sheet_name: str, piau_im, piau_im_huat: str):
        """
        依【標音字庫】中的【校正音標】欄位進行更新，並將【校正音標】覆蓋至原【台語音標】。
        """
        try:
            han_ji_piau_im_sheet_name = '漢字注音'
            ensure_sheet_exists(wb, han_ji_piau_im_sheet_name)
            han_ji_piau_im_sheet = wb.sheets[han_ji_piau_im_sheet_name]

            piau_im_sheet_name = '標音字庫'
            piau_im_ji_khoo_dict = self.create_ji_khoo_dict_from_sheet(wb, piau_im_sheet_name)
        except Exception as e:
            raise ValueError(f"無法找到或建立工作表 '{sheet_name}'：{e}")

        try:
            for han_ji, entries in piau_im_ji_khoo_dict.ji_khoo_dict.items():
                if not isinstance(entries, list):
                    continue
                for entry in entries:
                    if not isinstance(entry, dict):
                        continue
                    tai_gi_im_piau = entry.get("tai_gi_im_piau", "")
                    kau_ziann_im_piau = entry.get("kenn_ziann_im_piau", "")
                    coordinates = entry.get("coordinates", [])

                    if not kau_ziann_im_piau or kau_ziann_im_piau == 'N/A':
                        if coordinates:
                            row_no, col_no = coordinates[0]
                            msg = f"{han_ji} [{tai_gi_im_piau}] / [{kau_ziann_im_piau}]"
                            print(f"({row_no}, {col_no}) = {msg}")
                        continue

                    for row, col in coordinates:
                        han_ji_piau_im_sheet.activate()
                        han_ji_piau_im_sheet.range((row, col)).select()
                        han_ji_cell = han_ji_piau_im_sheet.range((row, col))
                        tai_gi_im_piau_cell = han_ji_piau_im_sheet.range((row - 1, col))
                        han_ji_piau_im_cell = han_ji_piau_im_sheet.range((row + 1, col))

                        tai_gi_im_piau_cell.value = kau_ziann_im_piau
                        han_ji_piau_im_cell.value = tlpa_tng_han_ji_piau_im(
                            piau_im=piau_im,
                            piau_im_huat=piau_im_huat,
                            tai_gi_im_piau=kau_ziann_im_piau
                        )
                        han_ji_cell.color = (0, 255, 255)
                        han_ji_cell.font.color = (255, 0, 0)

                        msg = f"{han_ji} [{tai_gi_im_piau}] / [{kau_ziann_im_piau}]"
                        print(f"({row}, {col}) = {msg}")

        except Exception as e:
            logging_exception(msg=f"使用【標音字庫】之【校正音標】，改正【漢字注音】之【台語音標】作業異常！", error=e)
            raise

        try:
            piau_im_ji_khoo_dict.write_to_excel_sheet(wb=wb, sheet_name=piau_im_sheet_name)
        except Exception as e:
            logging_exception(msg=f"將【字典】存放之資料，更新工作表作業異常！", error=e)
            raise

        han_ji_piau_im_sheet.range('A1').select()
        return EXIT_CODE_SUCCESS


    def write_to_excel_sheet(self, wb, sheet_name: str) -> int:
        try:
            sheet = wb.sheets[sheet_name]
        except Exception:
            sheet = wb.sheets.add(sheet_name)

        sheet.clear()
        headers = ["漢字", "台語音標", "校正音標", "座標"]
        sheet.range("A1").value = headers

        data = []
        for han_ji, entries in self.ji_khoo_dict.items():
            for entry in entries:
                if not entry["coordinates"]:  # 若座標為空，跳過不寫入
                    continue
                coord_str = "; ".join(f"({r}, {c})" for r, c in entry["coordinates"])
                data.append([han_ji, entry["tai_gi_im_piau"], entry["kenn_ziann_im_piau"], coord_str])

        sheet.range("A2").value = data
        return 0


    def write_to_han_ji_zu_im_sheet(self, wb, sheet_name: str):
        from mod_excel_access import ensure_sheet_exists

        try:
            ensure_sheet_exists(wb, sheet_name)
            sheet = wb.sheets[sheet_name]
        except Exception as e:
            raise ValueError(f"無法找到或建立工作表 '{sheet_name}'：{e}")

        for han_ji, entry in self.items():
            for row, col in entry["coordinates"]:
                sheet.range((row, col)).value = han_ji
                sheet.range((row - 1, col)).value = entry["tai_gi_im_piau"]

        self.write_to_excel_sheet(wb, "缺字表")
        print(f"已成功將字典資料寫入工作表 '{sheet_name}'。")

    def remove_coordinate(self, han_ji: str, tai_gi_im_piau: str, coordinate: tuple):
        """
        移除指定漢字與音標下的某個座標；若座標清空則移除整筆項目。
        """
        if han_ji not in self.ji_khoo_dict:
            return

        entries = self.ji_khoo_dict[han_ji]
        to_delete = None
        for entry in entries:
            if entry["tai_gi_im_piau"] == tai_gi_im_piau:
                if coordinate in entry["coordinates"]:
                    entry["coordinates"].remove(coordinate)
                if len(entry["coordinates"]) == 0:
                    to_delete = entry
                break

        if to_delete:
            entries.remove(to_delete)

    @classmethod
    def create_ji_khoo_dict_from_sheet(cls, wb, sheet_name: str):
        from mod_excel_access import ensure_sheet_exists

        if not ensure_sheet_exists(wb, sheet_name):
            raise ValueError(f"無法找到工作表 '{sheet_name}'。")

        try:
            sheet = wb.sheets[sheet_name]
        except Exception as e:
            raise ValueError(f"無法找到工作表 '{sheet_name}'：{e}")

        data = sheet.range("A2").expand("table").value
        ji_khoo = cls()

        if data is None:
            return ji_khoo
        if not isinstance(data[0], list):
            data = [data]

        for row in data:
            han_ji = row[0] or ""
            tai_gi_im_piau = row[1] or "N/A"
            kenn_ziann_im_piau = row[2] or "N/A"
            coords_str = row[3] or ""

            coordinates = []
            if coords_str:
                coords_list = coords_str.split("; ")
                for coord in coords_list:
                    coord = coord.strip("()")
                    row_col = tuple(map(int, coord.split(", ")))
                    coordinates.append(row_col)

            for coord in coordinates:
                ji_khoo.add_entry(han_ji, tai_gi_im_piau, kenn_ziann_im_piau, coord)

        return ji_khoo

    # def remove_coordinate(self, han_ji: str, tai_gi_im_piau: str, coordinate: tuple):
    #     """
    #     移除指定【漢字】與【台語音標】對應項目中的【座標】。
    #     若該項目座標清單為空，則整筆項目從字典中移除。
    #     """
    #     if han_ji not in self.ji_khoo_dict:
    #         return

    #     entries = self.ji_khoo_dict[han_ji]
    #     for entry in entries:
    #         if entry["tai_gi_im_piau"] == tai_gi_im_piau:
    #             if coordinate in entry["coordinates"]:
    #                 entry["coordinates"].remove(coordinate)
    #             if len(entry["coordinates"]) == 0:
    #                 entries.remove(entry)
    #             break

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
            han_ji_piau_im = tlpa_tng_han_ji_piau_im(
                piau_im=self.piau_im,
                piau_im_huat=self.piau_im_huat,
                tai_gi_im_piau=tai_gi_im_piau
            )
            # 記錄到人工標音字庫
            self.jin_kang_piau_im_ji_khoo.add_entry(
                han_ji=cell.value,
                tai_gi_im_piau=tai_gi_im_piau,
                kenn_ziann_im_piau='N/A',
                coordinates=(row, col)
            )
            # 記錄到標音字庫
            self.piau_im_ji_khoo.add_entry(
                han_ji=cell.value,
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
            msg = f"{cell.value}： [{tai_gi_im_piau}] /【{han_ji_piau_im}】（人工標音：【{jin_kang_piau_im}】）"
        else:
            msg = f"找不到【{cell.value}】此字的【台語音標】！"

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
            is_eof: 處理訊息和是否到達文件結尾
        """
        # 初始化樣式
        self._reset_cell_style(cell)

        cell_value = cell.value

        jin_kang_piau_im = cell.offset(-2, 0).value  # 人工標音
        if jin_kang_piau_im and str(jin_kang_piau_im).strip() != "":
            self._process_jin_kang_piau_im(jin_kang_piau_im, cell, row, col)
            return

        # 檢查特殊字元
        if cell_value == 'φ':
            # 【文字終結】
            return True
        elif cell_value == '\n':
            #【換行】
            return False
        elif not is_han_ji(cell_value):
            self._process_non_han_ji(cell_value)
            return False
        else:
            self._process_han_ji(cell_value, cell, row, col)
            return False


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
    end_of_process = False
    for r in range(1, config.TOTAL_LINES + 1):
        if end_of_process:
            break
        line_no = r
        print('-' * 60)
        print(f"處理第 {line_no} 行...")
        row = config.line_start_row + (r - 1) * config.ROWS_PER_LINE + config.han_ji_row_offset
        for c in range(config.start_col, config.end_col + 1):
            row = row
            col = c
            active_cell = sheet.range((row, col))
            active_cell.select()
            # 處理儲存格
            print(f"儲存格：{xw.utils.col_name(col)}{row}（{row}, {col}）")
            end_of_process = processor.process_cell(active_cell, row, col)
            if end_of_process:
                break


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

        # 處理工作表
        sheet = wb.sheets[config.hanji_piau_im_sheet]
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