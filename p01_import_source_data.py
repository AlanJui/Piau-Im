# coding=utf-8

import xlwings as xw


def main_run(CONVERT_FILE_NAME):
    # %%
    # 打開活頁簿檔案
    # file_path = 'hoo-goa-chu-im.xlsx'
    file_path = CONVERT_FILE_NAME
    wb = xw.Book(file_path)

    # 指定來源工作表
    source_sheet = wb.sheets['工作表1']

    # 取得工作表內總列數
    source_rows = source_sheet.range('A' + str(wb.sheets[0].cells.last_cell.row)).end('up').row
    print(f'source_rows = {source_rows}')

    # ==========================================================
    # 備妥程式需使用之工作表
    # ==========================================================
    sheet_name_list = [
        "缺字表",
        "字庫表",
        "漢字注音表",
    ]
    # -----------------------------------------------------
    # 檢查工作表是否已存在
    for sheet_name in sheet_name_list:
        sheet = wb.sheets[sheet_name]
        try:
            sheet.select()
            sheet.clear()
            continue
        except:
            # CommandError 的 Exception 發生日，表工作表不存在
            # 新增程式需使用之工作表
            wb.sheets.add(name=sheet_name)

    tui_chiau_sheet = wb.sheets["漢字注音表"]

    # %%
    # -----------------------------------------------------
    # 將字串轉換成 List
    # Python code to convert string to list character-wise
    def convert_string_to_chars_list(string):
        list1 = []
        list1[:0] = string
        return list1

    # %%
    # ==========================================================
    # 主程式
    # ==========================================================

    i = 1  # index for target sheet
    for row in range(1, source_rows + 1):
        # Read data from source_sheet
        chit_hang_ji = str(source_sheet.range('A' + str(row)).value)
        hang_ji_str = chit_hang_ji.strip()

        # 讀到空白行
        if hang_ji_str == "None":
            hang_ji_str = "\n"
        else:
            hang_ji_str = f"{chit_hang_ji}\n"

        han_ji_range = convert_string_to_chars_list(hang_ji_str)

        # =========================================================
        # Write to target worksheet
        # =========================================================
        tui_chiau_sheet.range('A' + str(i)).options(transpose=True).value = han_ji_range

        ji_soo = len(han_ji_range)
        i += ji_soo
