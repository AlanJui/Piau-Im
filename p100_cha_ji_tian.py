# coding=utf-8
import re

import psycopg2
import xlwings as xw

import modules.han_ji_chu_im as ji


def main_run(CONVERT_FILE_NAME):
    # ==========================================================
    # 在「漢字注音表」B欄已有台羅拼音，需將之拆分成聲母、韻母、調號
    # 聲母、韻母、調號，分別存放在 C、D、E 欄
    # ==========================================================

    # 指定提供來源的【檔案】
    # file_path = 'hoo-goa-chu-im.xlsx'
    file_path = CONVERT_FILE_NAME
    wb = xw.Book(file_path)

    # 指定提供來源的【工作表】；及【總列數】
    source_sheet = wb.sheets["漢字注音表"]
    end_of_row_no = (
        source_sheet.range("A" + str(source_sheet.cells.last_cell.row)).end("up").row
    )
    print(f"end_row = {end_of_row_no}")

    # ==========================================================
    # 備妥程式需使用之工作表
    # ==========================================================
    sheet_name_list = [
        "缺字表",
        "字庫表",
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

    khiam_ji_sheet = wb.sheets["缺字表"]
    ji_khoo_sheet = wb.sheets["字庫表"]

    # ==========================================================
    # 在「漢字注音表」B欄已有台羅拼音，需將之拆分成聲母、韻母、調號
    # 聲母、韻母、調號，分別存放在 C、D、E 欄
    # ==========================================================
    han_ji_tsu_im_piau = wb.sheets["漢字注音表"]
    han_ji_tsu_im_piau.select()

    # =========================================================="
    # 資料庫",
    # =========================================================="
    conn = psycopg2.connect(
        database="alanjui", user="alanjui", host="127.0.0.1", port="5432"
    )
    db_cursor = conn.cursor()
    source_index = 1  # index for source sheet
    target_index = 1
    ji_khoo_index = 1
    khiam_ji_index = 1

    while source_index <= end_of_row_no:
        print(f"row = {source_index}")
        # 自 source_sheet 取出一個「欲查注音的漢字」(beh_tshue_tsu_im_e_ji)
        beh_tshue_tsu_im_e_ji = str(source_sheet.range("A" + str(source_index)).value).strip()

        # =========================================================
        # 如是空白或換行，處理換行
        # =========================================================
        if beh_tshue_tsu_im_e_ji == "\n" or beh_tshue_tsu_im_e_ji == "None":
            han_ji_tsu_im_piau.range("A" + str(target_index)).value = "\n"
            target_index += 1
            source_index += 1
            continue

        # =========================================================
        # 若取出之字為標點符號，則跳過，並繼續取下一個漢字。
        # =========================================================
        # piau_tiam = r"[，、：；。？！（）「」【】《》“]"
        # piau_tiam = r"[\uFF0C\uFF08-\uFF09\u2013-\u2014\u2026\\u2018-\u201D\u3000\u3001-\u303F\uFE50-\uFE5E]"
        piau_tiam = r"[\u2013-\u2026\u3000-\u303F\uFE50-\uFF20]"
        is_piau_tiam = re.search(piau_tiam, beh_tshue_tsu_im_e_ji, re.M | re.I)
        if is_piau_tiam:
            target_index += 1
            source_index += 1
            continue

        # =========================================================
        # 在【字庫】資料庫查找【注音碼】
        # SQL 查詢指令：自字庫查找某漢字之注音碼
        # =========================================================
        # sql = f"select id, han_ji, chu_im, freq, siann, un, tiau from han_ji where han_ji='{search_han_ji}'"
        sql = (
            "SELECT id, han_ji, chu_im, freq, siann, un, tiau "
            "FROM han_ji "
            f"WHERE han_ji='{beh_tshue_tsu_im_e_ji}' "
            "ORDER BY freq DESC;"
        )
        db_cursor.execute(sql)
        ji_e_piau_im = db_cursor.fetchall()

        # =========================================================
        # 若是查不到漢字的注音碼，在【缺字表】做記錄
        # =========================================================
        if not ji_e_piau_im:
            print(f"Can not find 【{beh_tshue_tsu_im_e_ji}】in Han-Ji-Khoo!!")
            khiam_ji_sheet.range("A" + str(khiam_ji_index)).value = beh_tshue_tsu_im_e_ji
            khiam_ji_index += 1
            target_index += 1
            source_index += 1
            continue

        # =========================================================
        # 自【字庫】查到的【漢字】，取出：聲母、韻母、調號
        # =========================================================
        han_ji_id = ji_e_piau_im[0][0]
        tsu_im = ji_e_piau_im[0][2]
        siann_bu = ji_e_piau_im[0][4]
        un_bu = ji_e_piau_im[0][5]
        tiau_ho = ji_e_piau_im[0][6]

        # =========================================================
        # 寫入：【漢字注音表】
        # =========================================================
        han_ji_tsu_im_piau.range("B" + str(target_index)).value = tsu_im
        han_ji_tsu_im_piau.range("C" + str(target_index)).value = siann_bu
        han_ji_tsu_im_piau.range("D" + str(target_index)).value = un_bu
        han_ji_tsu_im_piau.range("E" + str(target_index)).value = tiau_ho
        han_ji_tsu_im_piau.range("F" + str(target_index)).value = han_ji_id

        # =========================================================
        # 若是查到漢字有一個以上的注音碼，在【字庫表】做記錄
        # =========================================================
        piau_im_tsong_soo = len(ji_e_piau_im)
        for piau_im_index in range(piau_im_tsong_soo):
            if piau_im_index == 0:
                continue

            # 若查到注音的漢字，有兩個以上；則需記錄漢字的 ID 編碼
            han_ji_id = ji_e_piau_im[piau_im_index][0]
            tsu_im = ji_e_piau_im[piau_im_index][2]
            siann_bu = ji_e_piau_im[piau_im_index][4]
            un_bu = ji_e_piau_im[piau_im_index][5]
            tiau_ho = ji_e_piau_im[piau_im_index][6]
            # ===========================================

            # 若查到的漢字有兩個以上
            # ji_khoo_sheet  = wb.sheets["字庫表"]
            ji_khoo_sheet.range("A" + str(ji_khoo_index)).value = beh_tshue_tsu_im_e_ji
            ji_khoo_sheet.range("B" + str(ji_khoo_index)).value = tsu_im
            ji_khoo_sheet.range("C" + str(ji_khoo_index)).value = siann_bu
            ji_khoo_sheet.range("D" + str(ji_khoo_index)).value = un_bu
            ji_khoo_sheet.range("E" + str(ji_khoo_index)).value = tiau_ho
            # 記錄【字庫】資料庫的【紀錄識別碼（Record ID of Table）】
            ji_khoo_sheet.range("F" + str(ji_khoo_index)).value = han_ji_id
            # 記錄對映【漢字注音表】的【列號（Excel Row Number）】
            ji_khoo_sheet.range("G" + str(ji_khoo_index)).value = source_index

            ji_khoo_index += 1

        # =========================================================
        # 調整讀取來源；寫入標的各手標
        # =========================================================
        target_index += 1
        source_index += 1

    # ==========================================================
    # 關閉資料庫
    # ==========================================================
    conn.close()
