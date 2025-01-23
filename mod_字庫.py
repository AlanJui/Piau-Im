from mod_excel_access import ensure_sheet_exists


class JiKhooDict:
    def __init__(self):
        """
        初始化字典數據結構。
        """
        self.ji_khoo_dict = {}


    def add_entry(self, han_ji: str, tai_gi_im_piau: str, coordinates: tuple):
        """
        新建一筆【漢字】的資料。

        :param han_ji: 漢字。
        :param tai_gi_im_piau: 台語音標。
        :param coordinates: 漢字在【漢字注音】工作表中的座標 (row, col)。
        """
        if han_ji not in self.ji_khoo_dict:
            # 如果漢字不存在，初始化資料結構
            self.ji_khoo_dict[han_ji] = [1, tai_gi_im_piau, 'N/A', [coordinates]]
        else:
            raise ValueError(f"漢字 '{han_ji}' 已存在，請使用 update_entry 方法來更新資料。")


    def update_entry(self, han_ji: str, coordinates: tuple):
        """
        使用【漢字】為【總數】欄加一，並新增一個座標。

        :param han_ji: 漢字。
        :param coordinates: 新的座標 (row, col)。
        """
        if han_ji in self.ji_khoo_dict:
            # 增加總數
            self.ji_khoo_dict[han_ji][0] += 1
            # 增加新的座標
            self.ji_khoo_dict[han_ji][3].append(coordinates)
        else:
            raise ValueError(f"漢字 '{han_ji}' 不存在，請先使用 add_entry 方法新增資料。")


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


    def add_or_update_entry(self, han_ji: str, tai_gi_im_piau: str, coordinates: tuple):
        """
        新增或更新一筆【漢字】的資料。

        - 如果漢字已存在，將更新總數並新增座標。
        - 如果漢字不存在，將新建一筆資料。

        :param han_ji: 漢字。
        :param tai_gi_im_piau: 台語音標。
        :param coordinates: 漢字在【漢字注音】工作表中的座標 (row, col)。
        """
        if han_ji in self.ji_khoo_dict:
            # 如果漢字已存在，使用 update_entry 更新
            self.update_entry(han_ji, coordinates)
        else:
            # 如果漢字不存在，使用 add_entry 新增
            self.add_entry(han_ji, tai_gi_im_piau, coordinates)


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
            coords_str = "; ".join([f"({row}, {col})" for row, col in coordinates])
            data.append([han_ji, total_count, tai_gi_im_piau, kenn_ziann_im_piau, coords_str])

        sheet.range("A2").value = data
        return 0


    @classmethod
    def create_ji_khoo_dict(cls, wb, sheet_name: str):
        """
        自 Excel 工作表建立 JiKhooDict 字典。

        :param wb: Excel 活頁簿物件。
        :param sheet_name: 工作表名稱。
        :return: JiKhooDict 物件。
        """
        try:
            sheet = wb.sheets[sheet_name]
        except Exception as e:
            raise ValueError(f"無法找到工作表 '{sheet_name}'：{e}")

        # 讀取工作表內容
        data = sheet.range("A2").expand("table").value

        # 初始化 JiKhooDict
        ji_khoo = cls()

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


# 單元測試
if __name__ == "__main__":
    # ut01()
    # ut02()
    ut03()