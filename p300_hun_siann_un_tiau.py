import re
import xlwings as xw

from modules.han_ji import split_chu_im


def main_run(CONVERT_FILE_NAME):
    # 打開活頁簿檔案
    # file_path = 'hoo-goa-chu-im.xlsx'
    file_path = CONVERT_FILE_NAME
    wb = xw.Book(file_path)

    # 指定來源工作表
    source_sheet = wb.sheets["漢字注音表"]
    # 取得工作表內總列數
    end_row = (
        source_sheet.range("A" + str(source_sheet.cells.last_cell.row)).end("up").row
    )
    print(f"end_row = {end_row}")

    # ==========================================================
    # 備妥程式需使用之工作表
    # ==========================================================
    sheet_name_list = [
        "缺字表",
        "字庫表",
    ]
    # =========================================================
    # 檢查工作表是否已存在；若否：則建立
    # =========================================================
    for sheet_name in sheet_name_list:
        sheet = wb.sheets[sheet_name]
        try:
            sheet.select()
            sheet.clear()
            continue
        except Exception as e:
            # CommandError 的 Exception 發生日，表工作表不存在
            # 新增程式需使用之工作表
            print(e)
            wb.sheets.add(name=sheet_name)

    target_sheet = wb.sheets["漢字注音表"]

    # ==========================================================
    # 主程式
    # ==========================================================

    i = 1  # index for target sheet
    row = 1  # index for source sheet
    end_counter = end_row + 1

    while row < end_counter:
        # 自 source_sheet 取待注音漢字
        han_ji = str(source_sheet.range("A" + str(row)).value)
        han_ji.strip()

        # =========================================================
        # 如是空白或換行，處理換行
        # =========================================================
        if han_ji == "" or han_ji == "\n":
            i += 1
            row += 1
            continue

        # =========================================================
        # 如只是標點符號，不必處理為漢字注音的工作
        # =========================================================
        piau_tiam = r"[；：？！\uFF0C\uFF08-\uFF09\u2013-\u2014\u2026\\u2018-\u201D\u3000\u3001-\u303F]"  # noqa
        searchObj = re.search(piau_tiam, han_ji, re.M | re.I)
        if searchObj:
            i += 1
            row += 1
            continue

        # =========================================================
        # 在字庫中查不到注音的漢字，略過注音處理
        # =========================================================
        ping_im = str(source_sheet.range("B" + str(row)).value).strip()
        if ping_im == "":
            # 讀到空白儲存格，視為使用者：「欲終止一個段落」；故於目標工作表
            # 寫入一個「換行」字元。
            i += 1
            row += 1
            continue

        # =========================================================
        # 將台羅拼音拆分成聲母、韻母、調號
        # =========================================================
        siann = split_chu_im(ping_im)[0]
        un = split_chu_im(ping_im)[1]
        tiau = split_chu_im(ping_im)[2]

        # =========================================================
        # 將漢字之聲、韻與調，寫入【漢字注音表】
        # =========================================================
        target_sheet.range("C" + str(i)).value = siann
        target_sheet.range("D" + str(i)).value = un
        target_sheet.range("E" + str(i)).value = tiau

        # =========================================================
        # 調整讀取來源；寫入標的各工作表
        # =========================================================
        i += 1
        row += 1

    # =========================================================
    # 結束前，顯示轉換結果。
    # =========================================================
    wb.sheets["漢字注音表"].range("A1").select()

