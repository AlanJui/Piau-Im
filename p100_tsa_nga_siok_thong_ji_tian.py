import re

import sqlite3
import xlwings as xw

# 專案全域常數
from config_dev_env import DATABASE

def main_run(CONVERT_FILE_NAME):
    # ==========================================================
    # 在「漢字注音表」B欄已有台羅拼音，需將之拆分成聲母、韻母、調號
    # 聲母、韻母、調號，分別存放在 C、D、E 欄
    # ==========================================================

    # 指定提供來源的【檔案】
    file_path = CONVERT_FILE_NAME
    wb = xw.Book(file_path)

    # 指定提供來源的【工作表】；及【總列數】
    source_sheet = wb.sheets["漢字注音表"]
    end_of_row_no = (
        source_sheet.range("A" + str(source_sheet.cells.last_cell.row)).end("up").row
    )
    end_of_row_no = int(end_of_row_no) - 1
    print(f"end_row = {end_of_row_no}")

    # ==========================================================
    # 備妥程式需使用之工作表
    # ==========================================================
    sheet_name_list = [
        "缺字表",
        "字庫表",
    ]
    # ----------------------------------------------------------
    # 檢查工作表是否已存在？
    # 若已存在，則清除工作表內容；
    # 若不存在，則新增工作表
    # ----------------------------------------------------------
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

    khiam_ji_piau = wb.sheets["缺字表"]
    ji_khoo_piau = wb.sheets["字庫表"]

    # ==========================================================
    # 在「漢字注音表」B欄已有台羅拼音，需將之拆分成聲母、韻母、調號
    # 聲母、韻母、調號，分別存放在 C、D、E 欄
    # ==========================================================
    han_ji_tsu_im_piau = wb.sheets["漢字注音表"]
    han_ji_tsu_im_piau.select()

    # =========================================================="
    # 資料庫",
    # =========================================================="
    conn = sqlite3.connect(DATABASE)
    db_cursor = conn.cursor()
    source_index = 1  # index for source sheet
    target_index = 1
    ji_khoo_index = 1
    khiam_ji_index = 1

    while source_index <= end_of_row_no:
        print(f"row = {source_index}")
        # 自 source_sheet 取出一個「欲查注音的漢字」(beh_tshue_tsu_im_e_ji)
        beh_tshue_tsu_im_e_ji = str(
            source_sheet.range("A" + str(source_index)).value
        ).strip()

        # =========================================================
        # 如是空白或換行，處理換行
        # =========================================================
        if beh_tshue_tsu_im_e_ji == " " or beh_tshue_tsu_im_e_ji == "":
            target_index += 1
            source_index += 1
            continue
        elif beh_tshue_tsu_im_e_ji == "\n":
            han_ji_tsu_im_piau.range("A" + str(target_index)).value = "\n"
            target_index += 1
            source_index += 1
            continue

        # =========================================================
        # 若取出之字為標點符號，則跳過，並繼續取下一個漢字。
        # =========================================================
        piau_tiam_1 = r"[，、：；．。？！（）「」【】《》“]"
        piau_tiam_2 = r"[\uFF0C\uFF08-\uFF09\u2013-\u2014\u2026\\u2018-\u201D\u3000\u3001-\u303F\uFE50-\uFE5E]"  # noqa: E501
        # piau_tiam = r"[\u2013-\u2026\u3000-\u303F\uFE50-\uFF20]"
        piau_tiam = f"{piau_tiam_1}|{piau_tiam_2}"
        is_piau_tiam = re.search(piau_tiam, beh_tshue_tsu_im_e_ji, re.M | re.I)
        if is_piau_tiam:
            target_index += 1
            source_index += 1
            continue

        # =========================================================
        # 在【字庫】資料庫查找【注音碼】
        # SQL 查詢指令：自字庫查找某漢字之注音碼
        # =========================================================
        # sql = select id, han_ji, tl_im, freq, siann, un, tiau
        #           from han_ji
        #           where han_ji='{search_han_ji}'
        sql = (
            "SELECT 識別號, 漢字, 切音, 字韻, 聲調, 原始拼音, 舒促聲, 聲, 韻, 調, 拼音碼, 雅俗通標音, 十五音標音, 常用度 "
            "FROM 雅俗通字庫 "
            f"WHERE 漢字='{beh_tshue_tsu_im_e_ji}' "
            "ORDER BY 常用度 DESC;"
        )
        db_cursor.execute(sql)
        ji_e_piau_im = db_cursor.fetchall()

        # =========================================================
        # 若是查不到漢字的注音碼，在【缺字表】做記錄
        # =========================================================
        if not ji_e_piau_im:
            print(f"Can not find 【{beh_tshue_tsu_im_e_ji}】in Han-Ji-Khoo!!")
            # 記錄【缺字表】的【列號】
            khiam_ji_piau.range("A" + str(khiam_ji_index)).value = khiam_ji_index
            # 記錄【缺字表】的【漢字】
            khiam_ji_piau.range("B" + str(khiam_ji_index)).value = beh_tshue_tsu_im_e_ji
            # 記錄【漢字注音表】的【列號】
            khiam_ji_piau.range("C" + str(khiam_ji_index)).value = source_index
            khiam_ji_index += 1
            target_index += 1
            source_index += 1
            continue

        # =========================================================
        # 自【字庫】查到的【漢字】，取出：聲母、韻母、調號
        # =========================================================
        piau_im_tsong_soo = len(ji_e_piau_im)
        han_ji_id = ji_e_piau_im[0][0]
        tsu_im = ji_e_piau_im[0][10]
        siann_bu = ji_e_piau_im[0][7]
        un_bu = ji_e_piau_im[0][8]
        tiau_ho = ji_e_piau_im[0][9]
        freq = ji_e_piau_im[0][13]

        # =========================================================
        # 寫入：【漢字注音表】
        # =========================================================
        han_ji_tsu_im_piau.range("B" + str(target_index)).value = tsu_im
        han_ji_tsu_im_piau.range("C" + str(target_index)).value = siann_bu
        han_ji_tsu_im_piau.range("D" + str(target_index)).value = un_bu
        han_ji_tsu_im_piau.range("E" + str(target_index)).value = tiau_ho
        han_ji_tsu_im_piau.range("F" + str(target_index)).value = piau_im_tsong_soo
        han_ji_tsu_im_piau.range("G" + str(target_index)).value = freq

        # =========================================================
        # 若是查到漢字有一個以上的注音碼，在【字庫表】做記錄
        # ji_khoo_sheet  = wb.sheets["字庫表"]
        # =========================================================
        if piau_im_tsong_soo > 1:
            for piau_im_index in range(piau_im_tsong_soo):
                han_ji_id = ji_e_piau_im[piau_im_index][0]
                tsu_im = ji_e_piau_im[piau_im_index][2]
                freq = ji_e_piau_im[piau_im_index][3]
                siann_bu = ji_e_piau_im[piau_im_index][4]
                un_bu = ji_e_piau_im[piau_im_index][5]
                tiau_ho = ji_e_piau_im[piau_im_index][6]

                # 記錄對映至【漢字注音表】的【列號】
                ji_khoo_piau.range("A" + str(ji_khoo_index)).value = source_index

                # 記錄【字庫】資料庫的【紀錄識別碼（Record ID of Table）】
                ji_khoo_piau.range("B" + str(ji_khoo_index)).value = han_ji_id

                ji_khoo_piau.range(
                    "C" + str(ji_khoo_index)
                ).value = beh_tshue_tsu_im_e_ji
                ji_khoo_piau.range("D" + str(ji_khoo_index)).value = tsu_im
                ji_khoo_piau.range("E" + str(ji_khoo_index)).value = siann_bu
                ji_khoo_piau.range("F" + str(ji_khoo_index)).value = un_bu
                ji_khoo_piau.range("G" + str(ji_khoo_index)).value = tiau_ho
                ji_khoo_piau.range("H" + str(ji_khoo_index)).value = freq

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
