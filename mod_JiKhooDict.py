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
