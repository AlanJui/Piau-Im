# ✅ 修訂後的 mod_字庫.py（完整版本）

import logging
import os
import sys

import xlwings as xw
from dotenv import load_dotenv

from mod_excel_access import ensure_sheet_exists, get_total_rows_in_sheet
from mod_logging import logging_exception
from mod_標音 import tlpa_tng_han_ji_piau_im

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
EXIT_CODE_SUCCESS = 0
EXIT_CODE_FAILURE = 1
EXIT_CODE_NO_FILE = 1
EXIT_CODE_INVALID_INPUT = 2
EXIT_CODE_PROCESS_FAILURE = 3
EXIT_CODE_UNKNOWN_ERROR = 99

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

    def remove_coordinate(self, han_ji: str, tai_gi_im_piau: str, coordinate: tuple):
        """
        移除指定漢字與音標下的某個座標；若座標清空則移除整筆項目。
        """
        if han_ji not in self.ji_khoo_dict:
            return

        entries = self.ji_khoo_dict[han_ji]
        for entry in entries:
            if entry["tai_gi_im_piau"] == tai_gi_im_piau:
                if coordinate in entry["coordinates"]:
                    entry["coordinates"].remove(coordinate)
                if len(entry["coordinates"]) == 0:
                    entries.remove(entry)
                break

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

    # 其他既有的函式 (如 update_kau_ziang_im_piau, write_to_excel_sheet...) 保持不變

# === 其餘原始程式碼依然適用 ===
