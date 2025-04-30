# JiKhooDict 單元測試程式碼：支援一字多音多座標
from mod_字庫 import JiKhooDict

def ut04():
    import xlwings as xw
    wb = xw.Book('output7\\a702_Test_Case.xlsx')

    khuat_ji_piau = JiKhooDict.create_ji_khoo_dict_from_sheet(wb, "缺字表")
    han_ji = "郁"
    entries = khuat_ji_piau.get_entry(han_ji)

    for idx, entry in enumerate(entries, start=1):
        print(f"讀音{idx}：{entry['tai_gi_im_piau']} / {entry['kenn_ziann_im_piau']}")
        for coord_idx, coord in enumerate(entry['coordinates'], start=1):
            print(f"　座標{coord_idx}：{coord}")

def ut05():
    import xlwings as xw
    wb = xw.Book('output7\\a702_Test_Case.xlsx')

    khuat_ji_piau = JiKhooDict.create_ji_khoo_dict_from_sheet(wb, "缺字表")
    han_ji = "郁"
    entries = khuat_ji_piau.get_entry(han_ji)

    for entry in entries:
        tai_gi_im_piau = entry["tai_gi_im_piau"]
        hau_zing_im_piau = entry["kenn_ziann_im_piau"]
        cells_list = entry["coordinates"]
        print(f"台語音標：{tai_gi_im_piau}")
        print(f"校正音標：{hau_zing_im_piau}")
        print(f"座標：{cells_list}")

        # 測試總數遞減更新
        print(f"總數（原）：{len(cells_list)}")
        if cells_list:
            new_coords = cells_list[:-1]
            khuat_ji_piau.update_value_by_key(han_ji, tai_gi_im_piau, "coordinates", new_coords)
            print(f"總數（更新後）：{len(new_coords)}")

def ut06():
    import xlwings as xw
    wb = xw.Book('output7\\a702_Test_Case.xlsx')

    khuat_ji_piau = JiKhooDict.create_ji_khoo_dict_from_sheet(wb, "缺字表")
    khuat_ji_piau.write_to_han_ji_zu_im_sheet(wb, "漢字注音")
    wb.save()

def ut07():
    import xlwings as xw
    wb = xw.Book()

    wb.sheets.add("漢字注音")
    wb.sheets.add("缺字表")

    ji_khoo = JiKhooDict()
    ji_khoo.add_entry("慶", "khing3", "N/A", (5, 3))
    ji_khoo.add_entry("慶", "khing3", "N/A", (57, 9))
    ji_khoo.add_entry("人", "jin5", "N/A", (5, 6))
    ji_khoo.add_entry("人", "jin5", "N/A", (97, 9))

    ji_khoo.write_to_han_ji_zu_im_sheet(wb, "漢字注音")
    wb.save("漢字庫.xlsx")
    wb.close()

def ut08(wb):
    sheet_name = "人工標音字庫"
    ji_khoo = JiKhooDict.create_ji_khoo_dict_from_sheet(wb, sheet_name)

    try:
        ji_khoo.add_or_update_entry("行", "kiann5", "N/A", (9, 7))
        ji_khoo.add_or_update_entry("行", "kiann5", "N/A", (21, 18))
        ji_khoo.write_to_excel_sheet(wb, sheet_name)
    except ValueError as e:
        print(f"❌ {e}")
        return 1

    return 0
