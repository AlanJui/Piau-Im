import logging
import os
import sys

import xlwings as xw
from dotenv import load_dotenv

from mod_excel_access import ensure_sheet_exists, get_total_rows_in_sheet

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
        """
        初始化字典數據結構。
        """
        self.ji_khoo_dict = {}


    def _is_han_ji_at_coordinates(self, han_ji: str, coordinates: tuple) -> bool:
        """
        檢查指定座標中是否確實存放了該漢字。

        :param han_ji: 漢字。
        :param coordinates: 座標 (row, col)。
        :return: 如果座標中的漢字與輸入的漢字一致，回傳 True；否則回傳 False。
        """
        # 從字典中獲取該漢字的座標列表
        if han_ji in self.ji_khoo_dict:
            existing_coordinates = self.ji_khoo_dict[han_ji][3]  # 座標列表在字典中的索引為 3
            return coordinates in existing_coordinates
        return False


    def items(self):
        """
        實現 items() 方法，回傳字典的鍵值對。
        """
        return self.ji_khoo_dict.items()


    def add_entry(self, han_ji: str, tai_gi_im_piau: str, kenn_ziann_im_piau: str, coordinates: tuple):
        """
        新建一筆【漢字】的資料。

        :param han_ji: 漢字。
        :param tai_gi_im_piau: 台語音標。
        :param coordinates: 漢字在【漢字注音】工作表中的座標 (row, col)。
        """
        if han_ji not in self.ji_khoo_dict:
            if tai_gi_im_piau is None or tai_gi_im_piau == "":
                tai_gi_im_piau = "N/A"
            # 如果漢字不存在，初始化資料結構
            kenn_ziann_im_piau = "N/A" if kenn_ziann_im_piau is None or kenn_ziann_im_piau == "" else kenn_ziann_im_piau
            self.ji_khoo_dict[han_ji] = [1, tai_gi_im_piau, kenn_ziann_im_piau, [coordinates]]
        else:
            raise ValueError(f"漢字 '{han_ji}' 已存在，請使用 update_entry 方法來更新資料。")


    # def update_entry(self, han_ji: str, kenn_ziann_im_piau: str, coordinates: tuple):
    #     """
    #     使用【漢字】為【總數】欄加一，並新增一個座標。

    #     :param han_ji: 漢字。
    #     :param coordinates: 新的座標 (row, col)。
    #     """
    #     if han_ji in self.ji_khoo_dict:
    #         # 增加總數
    #         self.ji_khoo_dict[han_ji][0] += 1
    #         # 更新校正音標
    #         self.ji_khoo_dict[han_ji][2] = kenn_ziann_im_piau
    #         # 增加新的座標
    #         self.ji_khoo_dict[han_ji][3].append(coordinates)
    #     else:
    #         raise ValueError(f"漢字 '{han_ji}' 不存在，請先使用 add_entry 方法新增資料。")
    def update_entry(self, han_ji: str, kenn_ziann_im_piau: str, coordinates: tuple):
        """
        使用【漢字】為【總數】欄加一，並新增一個座標。

        :param han_ji: 漢字。
        :param kenn_ziann_im_piau: 校正音標。
        :param coordinates: 新的座標 (row, col)。
        """
        if han_ji in self.ji_khoo_dict:
            # 增加總數
            self.ji_khoo_dict[han_ji][0] += 1
            # 更新校正音標（如果提供）
            if kenn_ziann_im_piau:
                self.ji_khoo_dict[han_ji][2] = kenn_ziann_im_piau
            # 增加新的座標
            self.ji_khoo_dict[han_ji][3].append(coordinates)
        else:
            raise ValueError(f"漢字 '{han_ji}' 不存在，請先使用 add_entry 方法新增資料。")

    # def add_or_update_entry(self, han_ji: str, tai_gi_im_piau: str, kenn_ziann_im_piau: str, coordinates: tuple):
    #     """
    #     新增或更新一筆【漢字】的資料。

    #     - 如果漢字已存在，將更新總數並新增座標。
    #     - 如果漢字不存在，將新建一筆資料。

    #     :param han_ji: 漢字。
    #     :param tai_gi_im_piau: 台語音標。
    #     :param coordinates: 漢字在【漢字注音】工作表中的座標 (row, col)。
    #     """
    #     if han_ji in self.ji_khoo_dict:
    #         # 如果漢字已存在，使用 update_entry 更新
    #         self.update_entry(han_ji, kenn_ziann_im_piau, coordinates)
    #     else:
    #         # 如果漢字不存在，使用 add_entry 新增
    #         self.add_entry(han_ji, tai_gi_im_piau, kenn_ziann_im_piau, coordinates)
    def add_or_update_entry(self, han_ji: str, tai_gi_im_piau: str, kenn_ziann_im_piau: str, coordinates: tuple):
        """
        新增或更新一筆【漢字】的資料。

        - 如果漢字已存在，並且在指定的座標中確實存放了該漢字，則更新總數並新增座標。
        - 如果漢字不存在，則新建一筆資料。

        :param han_ji: 漢字。
        :param tai_gi_im_piau: 台語音標。
        :param kenn_ziann_im_piau: 校正音標。
        :param coordinates: 漢字在【漢字注音】工作表中的座標 (row, col)。
        """
        # 檢查漢字是否已存在於字典中
        if han_ji in self.ji_khoo_dict:
            # 檢查該漢字是否確實存在於指定的座標中
            if self._is_han_ji_at_coordinates(han_ji, coordinates):
                # 如果漢字已存在且座標正確，使用 update_entry 更新
                self.update_entry(han_ji, kenn_ziann_im_piau, coordinates)
            else:
                # 如果漢字存在但座標不正確，視為新增一筆資料
                self.add_entry(han_ji, tai_gi_im_piau, kenn_ziann_im_piau, coordinates)
        else:
            # 如果漢字不存在，使用 add_entry 新增
            self.add_entry(han_ji, tai_gi_im_piau, kenn_ziann_im_piau, coordinates)


    def get_entry(self, han_ji: str):
        """
        使用【漢字】取用其【台語音標】、【總數】、【座標】欄的值。

        :param han_ji: 漢字。
        :return: 包含台語音標、總數和座標的列表 [台語音標, 總數, 座標列表]。
        """
        if han_ji in self.ji_khoo_dict:
            return self.ji_khoo_dict[han_ji]
        else:
            raise ValueError(f"漢字 '{han_ji}' 不存在於字典中。")


    def get_value_by_key(self, han_ji: str, key: str):
        """
        使用【漢字】取用其【台語音標】、【總數】、【座標】欄的值。

        :param han_ji: 漢字。
        :param key: 欄位名稱。
        :return: 欄位值。
        """
        if han_ji in self.ji_khoo_dict:
            if key == "台語音標":
                return self.ji_khoo_dict[han_ji][1]
            elif key == "校正音標":
                return self.ji_khoo_dict[han_ji][2]
            elif key == "總數":
                return self.ji_khoo_dict[han_ji][0]
            elif key == "座標":
                return self.ji_khoo_dict[han_ji][3]
            else:
                raise ValueError(f"無法識別的欄位名稱 '{key}'。")
        else:
            raise ValueError(f"漢字 '{han_ji}' 不存在於字典中。")


    def update_value_by_key(self, han_ji: str, key: str, value):
        """
        使用【漢字】更新其【台語音標】、【總數】、【座標】欄的值。

        :param han_ji: 漢字。
        :param key: 欄位名稱。
        :param value: 新的欄位值。
        """
        if han_ji in self.ji_khoo_dict:
            if key == "台語音標":
                self.ji_khoo_dict[han_ji][1] = value
            elif key == "校正音標":
                self.ji_khoo_dict[han_ji][2] = value
            elif key == "總數":
                self.ji_khoo_dict[han_ji][0] = value
            elif key == "座標":
                self.ji_khoo_dict[han_ji][3] = value
            else:
                raise ValueError(f"無法識別的欄位名稱 '{key}'。")
        else:
            raise ValueError(f"漢字 '{han_ji}' 不存在於字典中。")


    def write_to_excel_sheet(self, wb, sheet_name: str) -> int:
        """
        將【字典】寫入 Excel 工作表。

        :param wb: Excel 活頁簿物件。
        :param sheet_name: 工作表名稱。
        :return: 狀態碼（0 表成功，1 表失敗）。
        """
        try:
            sheet = wb.sheets[sheet_name]
        except Exception:
            sheet = wb.sheets.add(sheet_name)

        # 清空工作表內容
        sheet.clear()

        # 寫入標題列
        headers = ["漢字", "總數", "台語音標", "校正音標", "座標"]
        sheet.range("A1").value = headers

        # 寫入字典內容
        data = []
        for han_ji, (total_count, tai_gi_im_piau, kenn_ziann_im_piau, coordinates) in self.ji_khoo_dict.items():
            if total_count >= 2:
                coords_str = "; ".join([f"({row}, {col})" for row, col in coordinates])
            else:
                coords_str = f"{coordinates[0]}"
            data.append([han_ji, total_count, tai_gi_im_piau, kenn_ziann_im_piau, coords_str])

        sheet.range("A2").value = data
        return 0


    def write_ji_khoo_dict_to_sheet(self, wb, sheet_name: str, ji_khoo_dict: dict):
        """
        將 khuat_ji_piau 字典的資料寫回【缺字表】工作表。

        :param wb: Excel 活頁簿物件。
        :param sheet_name: 工作表名稱（例如「缺字表」）。
        :param khuat_ji_piau: 基於【缺字表】工作表建置的字典。
        """
        try:
            # 確保工作表存在
            ensure_sheet_exists(wb, sheet_name)
            sheet = wb.sheets[sheet_name]
        except Exception as e:
            raise ValueError(f"無法找到或建立工作表 '{sheet_name}'：{e}")

        # 清空工作表內容
        sheet.clear()

        # 寫入標題列
        headers = ["漢字", "總數", "台語音標", "校正音標", "座標"]
        sheet.range("A1").value = headers

        # 寫入字典內容
        data = []
        for han_ji, (total_count, tai_gi_im_piau, kenn_ziann_im_piau, coordinates) in ji_khoo_dict.items():
            coords_str = "; ".join([f"({row}, {col})" for row, col in coordinates])
            data.append([han_ji, total_count, tai_gi_im_piau, kenn_ziann_im_piau, coords_str])

        sheet.range("A2").value = data


    def write_to_han_ji_zu_im_sheet(self, wb, sheet_name: str, ji_khoo_dict: dict):
        """
        將字典中的所有漢字資料寫入 Excel 的「漢字注音」工作表。

        :param wb: Excel 活頁簿物件。
        :param sheet_name: 工作表名稱（例如「漢字注音」）。
        """
        try:
            # 確保工作表存在
            ensure_sheet_exists(wb, sheet_name)
            sheet = wb.sheets[sheet_name]
        except Exception as e:
            raise ValueError(f"無法找到或建立工作表 '{sheet_name}'：{e}")

        # 遍歷字典中的每個漢字
        # for han_ji, (total_count, tai_gi_im_piau, kenn_ziann_im_piau, coordinates) in self.ji_khoo_dict.items():
        for han_ji, (total_count, tai_gi_im_piau, kenn_ziann_im_piau, coordinates) in ji_khoo_dict.items():
            # 遍歷每個座標
            for row, col in coordinates:
                # 將漢字和台語音標寫入指定座標
                sheet.range((row, col)).select()
                # sheet.range((row, col)).value = han_ji
                sheet.range((row-1, col)).value = tai_gi_im_piau
                # 每寫入一次，total_count 減 1
                # self.ji_khoo_dict[han_ji][0] -= 1
                ji_khoo_dict[han_ji][0] -= 1

        # 將 khuat_ji_piau 字典寫回【缺字表】工作表
        self.write_ji_khoo_dict_to_sheet(wb, "缺字表", ji_khoo_dict)

        print(f"已成功將字典資料寫入工作表 '{sheet_name}'。")


    @classmethod
    def create_ji_khoo_dict_from_sheet(cls, wb, sheet_name: str):
        """
        根據輸入的【工作表名稱】，建立並回傳一個 JiKhooDict 物件。

        :param wb: Excel 活頁簿物件。
        :param sheet_name: 工作表名稱。
        :return: JiKhooDict 物件。
        """
        # 確保工作表存在
        if not ensure_sheet_exists(wb, sheet_name):
            raise ValueError(f"無法找到工作表 '{sheet_name}'。")

        try:
            sheet = wb.sheets[sheet_name]
        except Exception as e:
            raise ValueError(f"無法找到工作表 '{sheet_name}'：{e}")

        # 讀取工作表內容
        data = sheet.range("A2").expand("table").value

        # 初始化 JiKhooDict
        ji_khoo = cls()

        # 檢查 data 是否為 None 或空列表
        if data is None:
            return ji_khoo  # 返回空的 JiKhooDict 物件

        # 確保資料為 2D 列表
        if not isinstance(data[0], list):
            data = [data]

        # 將工作表內容轉換為字典結構
        for row in data:
            han_ji = row[0] or ""
            total_count = int(row[1]) if isinstance(row[1], (int, float)) else 0
            tai_gi_im_piau = row[2] or ""
            kenn_ziann_im_piau = row[3] or ""
            coords_str = row[4] or ""

            # 解析座標字串
            coordinates = []
            if coords_str:
                coords_list = coords_str.split("; ")
                for coord in coords_list:
                    coord = coord.strip("()")
                    row_col = tuple(map(int, coord.split(", ")))
                    coordinates.append(row_col)

            # 新增至字典
            ji_khoo.ji_khoo_dict[han_ji] = [total_count, tai_gi_im_piau, kenn_ziann_im_piau, coordinates]

        return ji_khoo


    def __getitem__(self, han_ji: str):
        """
        支持通過下標訪問漢字的資料。
        """
        if han_ji in self.ji_khoo_dict:
            return self.ji_khoo_dict[han_ji]
        else:
            raise KeyError(f"漢字 '{han_ji}' 不存在於字典中。")


    def __repr__(self):
        """
        顯示整個字典的內容，用於調試。
        """
        return repr(self.ji_khoo_dict)


def ut01():
    # 初始化 JiKhooDict
    ji_khoo = JiKhooDict()

    han_ji = "慶"
    result = ji_khoo.get_entry(han_ji)
    print(result)

    print(f'漢字：{han_ji}')
    print(f'台語音標：{result[0]}')
    print(f'總數：{result[1]}')
    print(f'座標：{result[2]}')

    # 顯示所有座標
    # for row, col in result[2]:
    #     print(f'座標：({row}, {col})')
    for idx, (row, col) in enumerate(result[2], start=1):
        print(f"座標{idx}：({row}, {col})")

    # 取得第三個座標
    sn = 3
    print(f"\n座標{sn}：({result[2][sn-1][0]}, {result[2][sn-1][1]})")


def ut02():
    import xlwings as xw

    # 測試用 Excel 活頁簿
    wb = xw.Book()

    # 初始化 JiKhooDict
    ji_khoo = JiKhooDict()

    # 新增資料
    ji_khoo.add_entry("慶", "khing3", (5, 3))
    ji_khoo.add_entry("人", "jin5", (5, 6))

    # 更新資料
    ji_khoo.update_entry("慶", (133, 11))
    ji_khoo.update_entry("慶", (145, 7))
    ji_khoo.update_entry("人", (97, 9))

    # 寫入 Excel
    ji_khoo.write_to_excel_sheet(wb, "漢字庫")

    # 從 Excel 建立字典
    new_ji_khoo = JiKhooDict.create_ji_khoo_dict(wb, "漢字庫")

    # 查看整個字典
    print(new_ji_khoo)

    # 取得第三個座標
    # sn = 3
    # print(f"\n座標{sn}：({new_ji_khoo[2][sn-1][0]}, {new_ji_khoo[2][sn-1][1]})")
    entry = new_ji_khoo["慶"]  # 獲取 '慶' 的資料
    third_coordinate = entry[3][2]  # 取得第三個座標
    print(f"座標3：({third_coordinate[0]}, {third_coordinate[1]})")

    entry2 = new_ji_khoo.get_entry("慶")  # 獲取 '不存在的漢字' 的資料
    print(f'entry2: {entry2}')
    print(f'entry2[1]: {entry2[1]}')

    # entry3 = new_ji_khoo["動"]  # 獲取 '不存在的漢字' 的資料
    entry3 = new_ji_khoo.get_entry("動")  # 獲取 '不存在的漢字' 的資料
    if entry3:
        print(entry3)


def ut03():
    import xlwings as xw

    # 測試用 Excel 活頁簿
    wb = xw.Book()

    # 初始化 JiKhooDict
    ji_khoo = JiKhooDict()

    # 新增或更新資料
    ji_khoo.add_or_update_entry("慶", "khing3", (5, 3))
    ji_khoo.add_or_update_entry("人", "jin5", (5, 6))
    ji_khoo.add_or_update_entry("慶", "khing3", (133, 11))
    ji_khoo.add_or_update_entry("慶", "khing3", (145, 7))
    ji_khoo.add_or_update_entry("人", "jin5", (97, 9))

    # 寫入 Excel
    ji_khoo.write_to_excel_sheet(wb, "漢字庫")

    # 從 Excel 建立字典
    new_ji_khoo = JiKhooDict.create_ji_khoo_dict(wb, "漢字庫")

    # 查看整個字典
    print(new_ji_khoo)

    # 獲取第三個座標
    entry = new_ji_khoo["慶"]
    third_coordinate = entry[3][2]
    print(f"座標3：({third_coordinate[0]}, {third_coordinate[1]})")


def ut04():
    import xlwings as xw

    # 測試用 Excel 活頁簿
    wb = xw.Book('output7\\a702_Test_Case.xlsx')

    # 初始化 JiKhooDict
    khuat_ji_piau = JiKhooDict.create_ji_khoo_dict(wb, "缺字表")

    # 新增或更新資料
    han_ji = "郁"
    tai_gi_im_piau = khuat_ji_piau[han_ji][1]
    hau_zing_im_piau = khuat_ji_piau[han_ji][2]
    cells_list = khuat_ji_piau[han_ji][3]
    # tai_gi_im_piau, hau_zing_im_piau, cells_list = khuat_ji_piau[han_ji]
    print(f"台語音標：{tai_gi_im_piau}")
    print(f"校正音標：{hau_zing_im_piau}")
    print(f"座標：{cells_list}")

    # 使用 get_entry 方法
    han_ji = "郁"
    entry = khuat_ji_piau.get_entry(han_ji)
    tai_gi_im_piau = entry[1]
    hau_zing_im_piau = entry[2]
    cells_list = entry[3]
    print(f"台語音標：{tai_gi_im_piau}")
    print(f"校正音標：{hau_zing_im_piau}")
    print(f"座標：{cells_list}")


def ut05():
    import xlwings as xw

    # 測試用 Excel 活頁簿
    wb = xw.Book('output7\\a702_Test_Case.xlsx')

    # 初始化 JiKhooDict
    khuat_ji_piau = JiKhooDict.create_ji_khoo_dict(wb, "缺字表")

    # 新增或更新資料
    han_ji = "郁"
    tai_gi_im_piau = khuat_ji_piau.get_value_by_key(han_ji, "台語音標")
    hau_zing_im_piau = khuat_ji_piau.get_value_by_key(han_ji, "校正音標")
    total = khuat_ji_piau.get_value_by_key(han_ji, "總數")
    cells_list = khuat_ji_piau.get_value_by_key(han_ji, "座標")
    print(f"台語音標：{tai_gi_im_piau}")
    print(f"校正音標：{hau_zing_im_piau}")
    print(f"座標：{cells_list}")

    # 更新資料
    print(f"總數：{total}")
    total -= 1
    khuat_ji_piau.update_value_by_key(han_ji, "總數", total)
    print(khuat_ji_piau.get_value_by_key(han_ji, "總數"))
    print('--------------------------------------------------------')


def ut06():
    import xlwings as xw

    # 測試用 Excel 活頁簿
    wb = xw.Book('output7\\a702_Test_Case.xlsx')

    # 初始化 JiKhooDict
    khuat_ji_piau = JiKhooDict.create_ji_khoo_dict(wb, "缺字表")

    # 將字典資料寫入「漢字注音」工作表
    khuat_ji_piau.write_to_han_ji_zu_im_sheet(wb, "漢字注音", khuat_ji_piau)

    # 保存並關閉 Excel 活頁簿
    wb.save()


def ut07():
    import xlwings as xw

    # 測試用 Excel 活頁簿
    wb = xw.Book()

    # 新增工作表
    wb.sheets.add("漢字注音")
    wb.sheets.add("缺字表")

    # 初始化 JiKhooDict
    ji_khoo = JiKhooDict()

    # 新增資料
    ji_khoo.add_entry("慶", "khing3", (5, 3))
    ji_khoo.add_entry("人", "jin5", (5, 6))
    ji_khoo.update_entry("慶", (57, 9))
    ji_khoo.update_entry("慶", (133, 11))
    ji_khoo.update_entry("人", (97, 9))

    # 模擬 khuat_ji_piau 字典
    khuat_ji_piau = {
        "慶": [3, "khing3", "N/A", [(5, 3), (57, 9), (133, 11)]],
        "人": [2, "jin5", "N/A", [(5, 6), (97, 9)]]
    }

    # 將字典資料寫入「漢字注音」工作表，並更新 khuat_ji_piau
    ji_khoo.write_to_han_ji_zu_im_sheet(wb, "漢字注音", khuat_ji_piau)

    # 保存並關閉 Excel 活頁簿
    wb.save("漢字庫.xlsx")
    wb.close()

def ut08(wb):
    # 從工作表建立 JiKhooDict
    # ji_khoo = JiKhooDict.create_from_sheet(wb, sheet_name)
    sheet_name = "人工標音字庫"
    ji_khoo = JiKhooDict.create_ji_khoo_dict_from_sheet(wb, sheet_name)

    try:
        # 新增或更新資料
        ji_khoo.add_or_update_entry("行", "kiann5", "N/A", (9, 7))  # 新增一筆資料
        ji_khoo.add_or_update_entry("行", "kiann5", "N/A", (21, 18))  # 新增一筆資料

        # 寫入 Excel 工作表
        ji_khoo.write_to_excel_sheet(wb, sheet_name)
    except ValueError as e:
        print(f"❌ {e}")
        return EXIT_CODE_FAILURE

    return EXIT_CODE_SUCCESS

def process(wb):
    # ut01()
    # ut02()
    # ut03()
    # ut04()
    # ut05()
    # ut06()
    # ut07()
    return ut08(wb)


# 單元測試
def main():
    try:
        wb = xw.apps.active.books.active
    except Exception as e:
        print("找不到作用中的 Excel 工作簿！", e)
        print("❌ 執行程式前請打開 Excel 檔案！")
        return 1

    return_code = process(wb)
    if return_code == EXIT_CODE_SUCCESS:
        print("✅ 通過單元測試！")
        return EXIT_CODE_SUCCESS
    else:
        print("❌ 單元測試失敗！")
        return EXIT_CODE_FAILURE

if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)