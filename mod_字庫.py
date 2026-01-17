import logging
import os

from dotenv import load_dotenv

from mod_excel_access import ensure_sheet_exists
from mod_logging import logging_exception
from mod_標音 import (
    tlpa_tng_han_ji_piau_im,
)

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
# 載入環境變數
# =========================================================================
load_dotenv()

DB_HO_LOK_UE = os.getenv('DB_HO_LOK_UE', 'Ho_Lok_Ue.db')
DB_KONG_UN = os.getenv('DB_KONG_UN', 'Kong_Un.db')

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
    def __init__(self, name: str = ''):
        self.name = name
        self.ji_khoo_dict = {}

    @classmethod
    def create_ji_khoo_dict_from_sheet(cls, wb, sheet_name: str):
        """_summary_
        cls: 指 class JiKhooDict 本身
        Args:
            wb (_type_): _description_
            sheet_name (str): _description_

        Raises:
            ValueError: _description_
            ValueError: _description_

        Returns:
            _type_: _description_
        """
        from mod_excel_access import ensure_sheet_exists

        if not ensure_sheet_exists(wb, sheet_name):
            raise ValueError(f"無法找到工作表 '{sheet_name}'。")

        try:
            sheet = wb.sheets[sheet_name]
            sheet.activate()
            sheet.range('A1').value = '漢字'
            sheet.range('B1').value = '台語音標'
            sheet.range('C1').value = '校正音標'
            sheet.range('D1').value = '座標'
        except Exception as e:
            raise ValueError(f"無法找到工作表 '{sheet_name}'：{e}")

        data = sheet.range("A2").expand("table").value
        ji_khoo = cls(sheet_name)

        if data is None:
            return ji_khoo
        if not isinstance(data[0], list):
            data = [data]

        for row in data:
            han_ji = row[0] or ""
            tai_gi_im_piau = row[1] or "N/A"
            hau_ziann_im_piau = row[2] or "N/A"
            coords_str = row[3] or ""

            coordinates = []
            if coords_str:
                coords_list = coords_str.split("; ")
                for coord in coords_list:
                    coord = coord.strip("()")
                    row_col = tuple(map(int, coord.split(", ")))
                    coordinates.append(row_col)

            for coord in coordinates:
                ji_khoo.add_entry(han_ji, tai_gi_im_piau, hau_ziann_im_piau, coord)

        return ji_khoo

    def items(self):
        for han_ji, entries in self.ji_khoo_dict.items():
            for entry in entries:
                yield (han_ji, entry)

    def add_entry(
        self,
        han_ji: str,
        tai_gi_im_piau: str,
        hau_ziann_im_piau: str,
        coordinates: tuple[int, int]
    ):
        if not tai_gi_im_piau:
            tai_gi_im_piau = "N/A"
        if not hau_ziann_im_piau:
            hau_ziann_im_piau = "N/A"

        entry = {
            "tai_gi_im_piau": tai_gi_im_piau,
            "hau_ziann_im_piau": hau_ziann_im_piau,
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

    def update_entry(
        self,
        han_ji: str,
        tai_gi_im_piau: str,
        hau_ziann_im_piau: str,
        coordinates: tuple[int, int]
    ):
        if han_ji not in self.ji_khoo_dict:
            raise ValueError(f"漢字 '{han_ji}' 不存在，請先使用 add_entry 方法新增資料。")

        for existing in self.ji_khoo_dict[han_ji]:
            if existing["tai_gi_im_piau"] == tai_gi_im_piau:
                if hau_ziann_im_piau:
                    existing["hau_ziann_im_piau"] = hau_ziann_im_piau
                if coordinates not in existing["coordinates"]:
                    existing["coordinates"].append(coordinates)
                return

        self.add_entry(han_ji, tai_gi_im_piau, hau_ziann_im_piau, coordinates)

    def get_row_by_han_ji_and_coordinate(self, han_ji: str, coordinate: tuple[int, int]) -> int:
        """
        根據【漢字】欄和【座標】欄，查詢【工作表】所在【列號】
        若查無結果，返回 -1

        Args:
            han_ji: 要查詢的漢字
            coordinate: 要查詢的座標

        Returns:
            int: 工作表的列號，若無則返回 -1
        """
        #  列號從 2 開始（第1列是標題）
        row_no = 2

        # 遍歷所有漢字及其對應的音標項目
        for current_han_ji, entries in self.ji_khoo_dict.items():
            for entry in entries:
                # 跳過沒有座標的項目（這些不會寫入 Excel）
                if not entry.get("coordinates"):
                    continue

                # 檢查是否匹配目標漢字和音標
                if current_han_ji == han_ji and coordinate in entry.get("coordinates", []):
                    return row_no

                # 每個有效項目佔一行
                row_no += 1

        # 找不到匹配項目
        return -1

    def add_or_update_entry(
        self,
        han_ji: str,
        tai_gi_im_piau: str,
        hau_ziann_im_piau: str,
        coordinates: tuple[int, int]
    ):
        """新增或更新【字典】中的【漢字】項目
        若【漢字】與【台語音標】已存在，則更新【校正音標】與【座標】
        否則新增一筆新項目
        Args:
            han_ji: 要新增或更新的漢字
            tai_gi_im_piau: 要新增或更新的台語音標
            hau_ziann_im_piau: 要新增或更新的校正音標
            coordinates: 要新增或更新的座標
        """
        row_no = self.get_row_by_han_ji_and_coordinate(han_ji=han_ji, coordinate=coordinates)
        if row_no != -1:
            self.update_entry(
                han_ji=han_ji,
                tai_gi_im_piau=tai_gi_im_piau,
                hau_ziann_im_piau=hau_ziann_im_piau,
                coordinates=coordinates
            )
        else:
            self.add_entry(
                han_ji=han_ji,
                tai_gi_im_piau=tai_gi_im_piau,
                hau_ziann_im_piau=hau_ziann_im_piau,
                coordinates=coordinates
            )

    def add_or_update_entry_by_coordinate(
        self,
        han_ji: str,
        tai_gi_im_piau: str,
        hau_ziann_im_piau: str,
        coordinates: tuple[int, int]
    ):
        return self.add_or_update_entry(
            han_ji=han_ji,
            tai_gi_im_piau=tai_gi_im_piau,
            hau_ziann_im_piau=hau_ziann_im_piau,
            coordinates=coordinates
        )

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

    def get_coordinates_by_han_ji_and_tai_gi_im_piau(self, han_ji: str, tai_gi_im_piau: str) -> list:
        """
        根據漢字與台語音標查詢工作表中的所有座標列表
        若查無結果，返回空列表

        Args:
            han_ji: 要查詢的漢字
            tai_gi_im_piau: 要查詢的台語音標

        Returns:
            list: 座標列表，若無則返回空列表
        """
        if han_ji in self.ji_khoo_dict:
            for entry in self.ji_khoo_dict[han_ji]:
                if entry["tai_gi_im_piau"] == tai_gi_im_piau:
                    return entry.get("coordinates", [])
        return []

    def get_row_by_han_ji_and_tai_gi_im_piau(self, han_ji: str, tai_gi_im_piau: str) -> int:
        """
        根據漢字與台語音標查詢工作表所在列號
        若查無結果，返回 -1

        Args:
            han_ji: 要查詢的漢字
            tai_gi_im_piau: 要查詢的台語音標

        Returns:
            int: 工作表的列號，若無則返回 -1
        """
        #  列號從 2 開始（第1列是標題）
        row_no = 2

        # 遍歷所有漢字及其對應的音標項目
        for current_han_ji, entries in self.ji_khoo_dict.items():
            for entry in entries:
                # 跳過沒有座標的項目（這些不會寫入 Excel）
                if not entry.get("coordinates"):
                    continue

                # 檢查是否匹配目標漢字和音標
                if current_han_ji == han_ji and entry.get("tai_gi_im_piau", "") == tai_gi_im_piau:
                    return row_no

                # 每個有效項目佔一行
                row_no += 1

        # 找不到匹配項目
        return -1

    def get_entry_by_han_ji_and_coordinate(self, han_ji: str, coordinate: tuple[int, int]) -> dict:
        """
        根據漢字與座標查詢對應的音標項目
        若查無結果，返回 None

        Args:
            han_ji: 要查詢的漢字
            coordinate: 要查詢的座標

        Returns:
            dict: 音標項目字典，若無則返回 None
        """
        if han_ji in self.ji_khoo_dict:
            entries = self.ji_khoo_dict[han_ji]
            for entry in entries:
                if coordinate in entry.get("coordinates", []):
                    return entry
        return None

    def get_entry_by_coordinate(self, coordinate: tuple[int, int]) -> tuple[str, dict]:
        """
        根據工作表座標查詢對應的漢字及其音標項目
        若查無結果，返回 None

        Args:
            coordinate: 要查詢的工作表座標

        Returns:
            tuple: (漢字, 音標項目字典)，若無則返回 None
        """
        for han_ji, entries in self.ji_khoo_dict.items():
            for entry in entries:
                # 跳過沒有座標的項目（這些不會寫入 Excel）
                if not entry.get("coordinates"):
                    continue

                if coordinate in entry.get("coordinates", []):
                    return han_ji, entry

        return None

    def get_entry_by_row_no(self, row_no: int) -> tuple[str, dict]:
        """
        根據工作表列號查詢對應的漢字及其音標項目
        若查無結果，返回 None

        Args:
            row_no: 要查詢的工作表列號

        Returns:
            tuple: (漢字, 音標項目字典)，若無則返回 None
        """
        current_row = 2  # 列號從 2 開始（第1列是標題）

        for han_ji, entries in self.ji_khoo_dict.items():
            for entry in entries:
                # 跳過沒有座標的項目（這些不會寫入 Excel）
                if not entry.get("coordinates"):
                    continue

                if current_row == row_no:
                    return han_ji, entry

                current_row += 1

        return None

    def get_tai_gi_im_piau_by_han_ji_and_coordinate(self, han_ji: str, coordinate: tuple[int, int]) -> str:
        """
        根據漢字與座標查詢台語音標
        若該漢字有多個音標，返回第一個符合座標的音標
        若查無結果，返回空字串

        Args:
            han_ji: 要查詢的漢字
            coordinate: 要查詢的座標

        Returns:
            str: 台語音標，若無則返回空字串
        """
        if han_ji in self.ji_khoo_dict:
            entries = self.ji_khoo_dict[han_ji]
            for entry in entries:
                if coordinate in entry.get("coordinates", []):
                    tai_gi_im_piau = entry.get("tai_gi_im_piau", "")
                    # 若該音標為 N/A 則返回空字串
                    if tai_gi_im_piau == "N/A":
                        return ""
                    return tai_gi_im_piau
        return ""

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

    # def update_kau_ziang_im_piau(self, han_ji: str, tai_gi_im_piau: str, hau_ziann_im_piau: str, coordinates: tuple):
    def update_hau_ziann_im_piau(self, han_ji: str, tai_gi_im_piau: str, hau_ziann_im_piau: str, coordinates: tuple):
        """
        將人工標音或校正音標更新至字典。
        如果該漢字、音標已存在則更新校正音標與座標。
        若尚未記錄該音標，則新增一筆。
        """
        if han_ji in self.ji_khoo_dict:
            for entry in self.ji_khoo_dict[han_ji]:
                if entry["tai_gi_im_piau"] == tai_gi_im_piau:
                    entry["hau_ziann_im_piau"] = hau_ziann_im_piau
                    if coordinates not in entry["coordinates"]:
                        entry["coordinates"].append(coordinates)
                    return
        # 若找不到，則新增新項目
        self.add_entry(han_ji, tai_gi_im_piau, hau_ziann_im_piau, coordinates)

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
                    kau_ziann_im_piau = entry.get("hau_ziann_im_piau", "")
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
            logging_exception(msg="使用【標音字庫】之【校正音標】，改正【漢字注音】之【台語音標】作業異常！", error=e)
            raise

        try:
            piau_im_ji_khoo_dict.write_to_excel_sheet(wb=wb, sheet_name=piau_im_sheet_name)
        except Exception as e:
            logging_exception(msg="將【字典】存放之資料，更新工作表作業異常！", error=e)
            raise

        han_ji_piau_im_sheet.range('A1').select()
        return EXIT_CODE_SUCCESS

    def save_to_sheet(self, wb, sheet_name: str) -> int:
        try:
            sheet_name_to_use = self.name if sheet_name == "" else sheet_name
            ensure_sheet_exists(wb, sheet_name_to_use)
            self.write_to_excel_sheet(wb, sheet_name_to_use)
            return EXIT_CODE_SUCCESS
        except Exception as e:
            logging_exception(msg="將【字典】存放之資料，更新工作表作業異常！", error=e)
            return EXIT_CODE_PROCESS_FAILURE

    def write_to_excel_sheet(self, wb, sheet_name: str) -> int:
        sheet_name_to_use = self.name if sheet_name == "" else sheet_name
        try:
            sheet = wb.sheets[sheet_name_to_use]
        except Exception:
            sheet = wb.sheets.add(sheet_name_to_use)
        sheet.clear()
        headers = ["漢字", "台語音標", "校正音標", "座標"]
        sheet.range("A1").value = headers

        data = []
        for han_ji, entries in self.ji_khoo_dict.items():
            for entry in entries:
                if not entry["coordinates"]:  # 若座標為空，跳過不寫入
                    continue
                coord_str = "; ".join(f"({r}, {c})" for r, c in entry["coordinates"])
                data.append([han_ji, entry["tai_gi_im_piau"], entry["hau_ziann_im_piau"], coord_str])

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

    def remove_coordinate(
            self,
            han_ji: str,
            tai_gi_im_piau: str,
            coordinate: tuple[int, int],
            entry_to_delete_if_empty: bool = False
        ):
        """
        根據【漢字】與【座標】移除紀錄中，在【座標】欄清單的某【座標】；
        若【座標】欄清空，則移除整筆紀錄。
        """
        if han_ji not in self.ji_khoo_dict:
            return

        entries = self.ji_khoo_dict[han_ji]
        to_delete = None
        for entry in entries:
            if coordinate in entry["coordinates"]:
                entry["coordinates"].remove(coordinate)
            if len(entry["coordinates"]) == 0:
                to_delete = entry
            break

        if to_delete and entry_to_delete_if_empty:
            entries.remove(to_delete)

    def remove_coordinate_by_han_ji_and_tai_gi_im_piau(
            self,
            han_ji: str,
            tai_gi_im_piau: str,
            coordinate: tuple[int, int],
            entry_to_delete_if_empty: bool = False
        ):
        """
        根據【漢字】與【台語音標】移除紀錄中，在【座標】欄清單的某【座標】；
        若【座標】欄清空，則移除整筆紀錄。
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

        if to_delete and entry_to_delete_if_empty:
            entries.remove(to_delete)


    def remove_coordinate_by_han_ji_and_coordinate(self, han_ji: str, coordinate: tuple[int, int]):
        """
        移除指定漢字與音標下的某個座標；若座標清空則移除整筆項目。
        """
        if han_ji not in self.ji_khoo_dict:
            return

        entries = self.ji_khoo_dict[han_ji]
        to_delete = None
        for entry in entries:
            for coord in entry["coordinates"]:
                if coord == coordinate:
                    entry["coordinates"].remove(coord)
                    if len(entry["coordinates"]) == 0:
                        to_delete = entry
                    break

        if to_delete:
            entries.remove(to_delete)

    def remove_coordinate_by_hau_ziann_im_piau(self, han_ji: str, hau_ziann_im_piau: str, coordinate: tuple):
        """
        移除指定漢字與音標下的某個座標；若座標清空則移除整筆項目。
        """
        if han_ji not in self.ji_khoo_dict:
            return

        entries = self.ji_khoo_dict[han_ji]
        to_delete = None
        for entry in entries:
            if entry["hau_ziann_im_piau"] == hau_ziann_im_piau:
                if coordinate in entry["coordinates"]:
                    entry["coordinates"].remove(coordinate)
                if len(entry["coordinates"]) == 0:
                    to_delete = entry
                break

        if to_delete:
            entries.remove(to_delete)

    def remove_entry(
            self,
            han_ji: str,
            tai_gi_im_piau: str,
            hau_ziann_im_piau: str,
            coordinates: tuple[int, int]
        ):
        """
        移除指定漢字與音標下的某個座標；若座標清空則移除整筆項目。
        """
        if han_ji not in self.ji_khoo_dict:
            return

        entries = self.ji_khoo_dict[han_ji]
        to_delete = None
        for entry in entries:
            if entry["tai_gi_im_piau"] == tai_gi_im_piau:
                if entry["hau_ziann_im_piau"] == hau_ziann_im_piau:
                    if coordinates in entry["coordinates"]:
                        entry["coordinates"].remove(coordinates)
                    if len(entry["coordinates"]) == 0:
                        to_delete = entry
                    break

        if to_delete:
            entries.remove(to_delete)
