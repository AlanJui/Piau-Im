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
    end_row = source_sheet.range('A' + str(source_sheet.cells.last_cell.row)).end('up').row
    print(f'end_row = {end_row}')

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
    chu_im_sheet = wb.sheets["漢字注音表"]

    row = 1     # index for source sheet
    chu_im_index = 1
    ji_khoo_index = 1
    khiam_ji_index = 1
    end_counter = end_row + 1

    while row < end_counter:
        print(f'row = {row}')
        # 自 source_sheet 取待注音漢字
        han_ji = str(source_sheet.range('A' + str(row)).value)
        han_ji.strip()

        # =========================================================
        # 如是空白或換行，處理換行
        # =========================================================
        if han_ji == '\n' or han_ji == 'None':
            chu_im_sheet.range('A' + str(chu_im_index)).value = '\n'
            chu_im_index += 1
            row += 1
            continue

        # =========================================================
        # 自【來源工作表】，讀入【整段】的漢字，轉存到目的工作表：【漢字注音表】
        # 在【漢字注音表】的每個儲存格，只存放一個漢字
        # =========================================================
        han_ji_list = ji.convert_string_to_list(han_ji)

        chu_im_sheet.range('A' + str(chu_im_index)) \
                    .options(transpose=True).value = han_ji_list

        # =========================================================
        # 將整段讀入的漢字，逐一加注音
        # =========================================================
        i = chu_im_index
        # 取用每一個漢字，自【字庫】資料庫查找其【注音碼】
        for search_han_ji in han_ji_list:
            # 若取出之字為【換行】控制字元，則繼續取下一個漢字
            if search_han_ji == "\n":
                i += 1
                continue

            # 若取出之字為標點符號，則跳過，並繼續取下一個漢字。
            # piau_tiam = r"[，、：；。？！（）「」【】《》“]"
            # piau_tiam = r"[\uFF0C\uFF08-\uFF09\u2013-\u2014\u2026\\u2018-\u201D\u3000\u3001-\u303F\uFE50-\uFE5E]"
            piau_tiam = r"[\u2013-\u2026\u3000-\u303F\uFE50-\uFF20]"
            searchObj = re.search(piau_tiam, search_han_ji, re.M | re.I)
            if searchObj:
                i += 1
                continue

            # SQL 查詢指令：自字庫查找某漢字之注音碼
            # sql = f"select id, han_ji, chu_im, freq, siann, un, tiau from han_ji where han_ji='{search_han_ji}'"
            sql = "SELECT id, han_ji, chu_im, freq, siann, un, tiau "\
                "FROM han_ji "\
                f"WHERE han_ji='{search_han_ji}' "\
                "ORDER BY freq DESC;"
            cur.execute(sql)
            query_rows = cur.fetchall()

            # 漢字能否查到注音碼，將有不同的處理作業
            if not query_rows:
                # 問題發生：找不到漢字的注音碼
                print(f"Can not find 【{search_han_ji}】in Han-Ji-Khoo!!")
                khiam_ji_sheet.range('A' + str(khiam_ji_index)).value = search_han_ji
                khiam_ji_index += 1
                i += 1
                continue

            ji_soo = len(query_rows)
            位於字庫表的列號清單 = []
            for ji_found in range(ji_soo):
                # 若查到注音的漢字，有兩個以上；則需記錄漢字的 ID 編碼
                han_ji_id = query_rows[ji_found][0]
                chu_im = query_rows[ji_found][2]
                # ===========================================
                # 自【字庫】查到的【漢字】，取出：聲母、韻母、調號
                # ===========================================
                siann_bu = query_rows[ji_found][4]
                un_bu = query_rows[ji_found][5]
                tiau_ho = query_rows[ji_found][6]

                # =========================================================
                # 將已注音之漢字加入【漢字注音表】
                # =========================================================
                if ji_found == 0:
                    # 處理查到的第一個漢字
                    chu_im_sheet.range('B' + str(i)).value = chu_im
                    chu_im_sheet.range('C' + str(i)).value = siann_bu
                    chu_im_sheet.range('D' + str(i)).value = un_bu
                    chu_im_sheet.range('E' + str(i)).value = tiau_ho
                else:
                    # 若查到的漢字有兩個以上
                    # ji_khoo_sheet  = wb.sheets["字庫表"]
                    idx = ji_khoo_index
                    ji_khoo_sheet.range('A' + str(idx)).value = search_han_ji

                    ji_khoo_sheet.range('B' + str(idx)).value = chu_im
                    ji_khoo_sheet.range('C' + str(idx)).value = siann_bu
                    ji_khoo_sheet.range('D' + str(idx)).value = un_bu
                    ji_khoo_sheet.range('E' + str(idx)).value = tiau_ho

                    # 記錄對映【漢字注音表】的【列號（Excel Row Number）】
                    ji_khoo_sheet.range('F' + str(idx)).value = i
                    # 記錄【字庫】資料庫的【紀錄識別碼（Record ID of Table）】
                    ji_khoo_sheet.range('G' + str(idx)).value = han_ji_id

                    位於字庫表的列號清單 += [idx]
                    ji_khoo_index += 1

                # 記錄漢字在字庫中所擁有的不同注音
                if ji_soo > 1:
                    # 同字異音紀錄 = f"本字共【{ji_soo}】個注音碼：參考【字庫表】第 {位於字庫表的列號清單} 列。"
                    chu_im_sheet.range('F' + str(i)).value = ji_soo
                    chu_im_sheet.range('G' + str(i)).value = 位於字庫表的列號清單

            # =========================================================
            # 計數【整段讀入，逐一處理的漢字】已處理多少個
            # =========================================================
            i += 1

        # =========================================================
        # 調整讀取來源；寫入標的各手標
        # =========================================================
        chu_im_index += len(han_ji_list)
        chu_im_sheet.range('A' + str(chu_im_index)).value = '\n'
        chu_im_index += 1
        row += 1

    # ==========================================================
    # 關閉資料庫
    # ==========================================================
    conn.close()
