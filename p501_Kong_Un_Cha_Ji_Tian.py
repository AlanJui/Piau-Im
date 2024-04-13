import re

import xlwings as xw

from mod_廣韻 import han_ji_cha_piau_im, piau_tiau_ho

# import sqlite3



# 專案全域常數
# from config_dev_env import DATABASE
# DATABASE = "Kong_Un_V2.db"


def Kong_Un_Piau_Im(CONVERT_FILE_NAME, db_cursor):
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
    han_ji_cu_im_piau = wb.sheets["漢字注音表"]
    han_ji_cu_im_piau.select()

    # =========================================================="
    # 資料庫",
    # =========================================================="
    # conn = sqlite3.connect(DATABASE)
    # db_cursor = conn.cursor()
    source_index = 1  # index for source sheet
    target_index = 1
    ji_khoo_index = 1
    khiam_ji_index = 1

    while source_index <= end_of_row_no:
        print(f"row = {source_index}")
        # 自 source_sheet 取出一個「欲查注音的漢字」(beh_piau_im_e_han_ji)
        beh_piau_im_e_han_ji = str(
            source_sheet.range("A" + str(source_index)).value
        ).strip()

        # =========================================================
        # 如是空白或換行，處理換行
        # =========================================================
        if beh_piau_im_e_han_ji == " " or beh_piau_im_e_han_ji == "":
            target_index += 1
            source_index += 1
            continue
        elif beh_piau_im_e_han_ji == "\n":
            han_ji_cu_im_piau.range("A" + str(target_index)).value = "\n"
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
        is_piau_tiam = re.search(piau_tiam, beh_piau_im_e_han_ji, re.M | re.I)
        if is_piau_tiam:
            target_index += 1
            source_index += 1
            continue

        # =========================================================
        # 在【字庫】資料庫查找【注音碼】
        # SQL 查詢指令：自字庫查找某漢字之注音碼
        # =========================================================
        kong_un_piau_im = han_ji_cha_piau_im(db_cursor, beh_piau_im_e_han_ji)

        # =========================================================
        # 若是查不到漢字的注音碼，在【缺字表】做記錄
        # =========================================================
        if not kong_un_piau_im:
            print(f"廣韻字典查不到此漢字：【{beh_piau_im_e_han_ji}】!!")
            # 記錄【缺字表】的【列號】
            khiam_ji_piau.range("A" + str(khiam_ji_index)).value = khiam_ji_index
            # 記錄【缺字表】的【漢字】
            khiam_ji_piau.range("B" + str(khiam_ji_index)).value = beh_piau_im_e_han_ji
            # 記錄【漢字注音表】的【列號】
            khiam_ji_piau.range("C" + str(khiam_ji_index)).value = source_index
            khiam_ji_index += 1
            target_index += 1
            source_index += 1
            continue

        # =========================================================
        # 自【字庫】查到的【漢字】，取出：聲母、韻母、調號
        # =========================================================
        piau_im_tsong_soo = len(kong_un_piau_im)
        piau_im = kong_un_piau_im[0]
        han_ji_id = piau_im['漢字識別號']
        # sing_bu = piau_im['上字標音'] if piau_im['上字標音'] != "Ø" else "q"
        sing_bu = piau_im['上字標音']
        un_bu = piau_im['下字標音']
        tiau_ho = piau_tiau_ho(piau_im)
        cu_im = f"{sing_bu}{un_bu}{tiau_ho}"

        # =========================================================
        # 寫入：【漢字注音表】
        # =========================================================
        han_ji_cu_im_piau.range("B" + str(target_index)).value = cu_im
        han_ji_cu_im_piau.range("C" + str(target_index)).value = sing_bu
        han_ji_cu_im_piau.range("D" + str(target_index)).value = un_bu
        han_ji_cu_im_piau.range("E" + str(target_index)).value = tiau_ho
        han_ji_cu_im_piau.range("F" + str(target_index)).value = piau_im_tsong_soo

        # =========================================================
        # 若是查到漢字有一個以上的注音碼，在【字庫表】做記錄
        # ji_khoo_sheet  = wb.sheets["字庫表"]
        # =========================================================
        if piau_im_tsong_soo > 1:
            for index in range(piau_im_tsong_soo):
                piau_im = kong_un_piau_im[index]
                han_ji_id = piau_im['漢字識別號']
                # sing_bu = piau_im['上字標音'] if piau_im['上字標音'] != "Ø" else "q"
                sing_bu = piau_im['上字標音']
                un_bu = piau_im['下字標音']
                tiau_ho = piau_tiau_ho(piau_im)
                cu_im = f"{sing_bu}{un_bu}{tiau_ho}"

                # 記錄對映至【漢字注音表】的【列號】
                ji_khoo_piau.range("A" + str(ji_khoo_index)).value = source_index

                # 記錄【字庫】資料庫的【紀錄識別碼（Record ID of Table）】
                ji_khoo_piau.range("B" + str(ji_khoo_index)).value = han_ji_id

                ji_khoo_piau.range("C" + str(ji_khoo_index)).value = (
                    beh_piau_im_e_han_ji
                )
                ji_khoo_piau.range("D" + str(ji_khoo_index)).value = cu_im
                ji_khoo_piau.range("E" + str(ji_khoo_index)).value = sing_bu
                ji_khoo_piau.range("F" + str(ji_khoo_index)).value = un_bu
                ji_khoo_piau.range("G" + str(ji_khoo_index)).value = tiau_ho

                ji_khoo_index += 1

        # =========================================================
        # 調整讀取來源；寫入標的各手標
        # =========================================================
        target_index += 1
        source_index += 1


if __name__ == "__main__":
    CONVERT_FILE_NAME = "output\\Piau-Tsu-Im.xlsx"
    Kong_Un_Piau_Im(CONVERT_FILE_NAME)