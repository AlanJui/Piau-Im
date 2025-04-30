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

    def write_to_excel_sheet(self, wb, sheet_name: str) -> int:
        try:
            sheet = wb.sheets[sheet_name]
        except Exception:
            sheet = wb.sheets.add(sheet_name)

        sheet.clear()
        headers = ["漢字", "台語音標", "校正音標", "座標"]
        sheet.range("A1").value = headers

        data = []
        for han_ji, entry in self.items():
            for coord in entry["coordinates"]:
                coord_str = f"({coord[0]}, {coord[1]})"
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