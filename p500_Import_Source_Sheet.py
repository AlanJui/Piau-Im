import xlwings as xw


def San_Sing_Han_Ji_Tsh_Im_Piau(CONVERT_FILE_NAME):
    # 打開活頁簿檔案
    file_path = CONVERT_FILE_NAME
    wb = xw.Book(file_path)

    # 指定來源工作表
    source_sheet = wb.sheets["工作表1"]
    source_sheet.select()

    # 取得工作表內總列數
    source_row_no = int(
        source_sheet.range("A" + str(wb.sheets[0].cells.last_cell.row)).end("up").row
    )
    print(f"source_row_no = {source_row_no}")

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
        sheets =  [sheet.name for sheet in wb.sheets]  # 獲取所有工作表的名稱
        if sheet_name in sheets:
            sheet = wb.sheets[sheet_name]
            try:
                sheet.select()
                sheet.clear()
                continue
            except Exception as e:
                print(e)
        else:
            # CommandError 的 Exception 發生時，表工作表不存在
            # 新增程式需使用之工作表
            print(f"工作表 {sheet_name} 不存在，正在新增...")
            wb.sheets.add(name=sheet_name)

    # 選用「漢字注音表」
    try:
        han_ji_tsu_im_paiu = wb.sheets["漢字注音表"]
        han_ji_tsu_im_paiu.select()
    except Exception as e:
        # 处理找不到 "漢字注音表" 工作表的异常
        print(e)
        print("找不到：〖漢字注音表〗工作表。")
        return False

    # -----------------------------------------------------
    # 將「字串」轉換成「串列（Characters List）」
    # Python code to convert string to list character-wise
    def convert_string_to_chars_list(string):
        list1 = []
        list1[:0] = string
        return list1

    # ==========================================================
    # (1)
    # ==========================================================
    # 自【工作表1】的每一列，讀入一個「段落」的漢字。然後將整個段
    # 落拆成「單字」，存到【漢字注音表】；在【漢字注音表】的每個
    # 儲存格，只存放一個「單字」。
    # ==========================================================

    source_row_index = 1
    target_row_index = 1  # index for target sheet
    # for row in range(1, source_rows):
    while source_row_index <= source_row_no:
        # 自【工作表1】取得「一行漢字」
        tsit_hang_ji = str(source_sheet.range("A" + str(source_row_index)).value)
        hang_ji_str = tsit_hang_ji.strip()

        # 讀到空白行
        if hang_ji_str == "None":
            hang_ji_str = "\n"
        else:
            hang_ji_str = f"{tsit_hang_ji}\n"

        han_ji_range = convert_string_to_chars_list(hang_ji_str)

        # =========================================================
        # 讀到的整段文字，以「單字」形式寫入【漢字注音表】。
        # =========================================================
        han_ji_tsu_im_paiu.range("A" + str(target_row_index)).options(
            transpose=True
        ).value = han_ji_range

        ji_soo = len(han_ji_range)
        target_row_index += ji_soo
        source_row_index += 1
