import xlwings as xw

from modules.han_ji import split_chu_im as hun_siann_un_tiau


def main_run(CONVERT_FILE_NAME):
    # 打開活頁簿檔案
    wb = xw.Book(CONVERT_FILE_NAME)

    # ==========================================================
    # 備妥程式需使用之工作表
    # ==========================================================
    sheet_name_list = [
        "漢字注音表",
        "缺字表",
    ]
    # 確認來源工作表皆已存在
    for sheet_name in sheet_name_list:
        if sheet_name not in [sheet.name for sheet in wb.sheets]:
            # print(f"欠缺工作表：{sheet_name}！")
            # return False
            raise Exception(f"欠缺工作表：{sheet_name}！")
        else:
            print(f"【{sheet_name}】工作表存在！")

    # ==========================================================
    # 自「缺字表」取得待補入「漢字注音表」的作業總數
    # ==========================================================
    # 指定來源工作表
    source_sheet = wb.sheets["缺字表"]
    source_sheet.select()
    # 取得工作表內總列數
    end_row_no = (
        source_sheet.range("A" + str(source_sheet.cells.last_cell.row)).end("up").row
    )
    # end_row_no -= 1
    print(f"end_row_no = {end_row_no}")

    # ==========================================================
    # 備妥「漢字注音表」
    # ==========================================================
    target_sheet = wb.sheets["漢字注音表"]
    target_sheet.select()

    # ==========================================================
    # 自【缺字表】取得漢字的注音，補入【漢字注音表】
    # ==========================================================
    target_row_index = 1  # index for target sheet
    source_row_index = 1  # index for source sheet

    while source_row_index <= end_row_no:
        # =========================================================
        # 自「缺字表」取得漢字注音
        # =========================================================
        tsu_im = str(source_sheet.range("D" + str(source_row_index)).value).strip()
        if tsu_im == "":
            # 讀到空白儲存格，表使用者仍未補上注音
            source_row_index += 1
            continue
        else:
            value = source_sheet.range("C" + str(source_row_index)).value
            target_row_index = int(value)

        # =========================================================
        # 將台羅拼音拆分成聲母、韻母、調號
        # =========================================================
        siann = hun_siann_un_tiau(tsu_im)[0]
        un = hun_siann_un_tiau(tsu_im)[1]
        tiau = hun_siann_un_tiau(tsu_im)[2]

        # =========================================================
        # 將漢字之聲、韻與調，寫入【漢字注音表】
        # =========================================================
        target_sheet.range("B" + str(target_row_index)).value = tsu_im
        target_sheet.range("C" + str(target_row_index)).value = siann
        target_sheet.range("D" + str(target_row_index)).value = un
        target_sheet.range("E" + str(target_row_index)).value = tiau
        target_sheet.range("F" + str(target_row_index)).value = 0

        # =========================================================
        # 調整讀取來源；寫入標的各工作表
        # =========================================================
        source_row_index += 1
