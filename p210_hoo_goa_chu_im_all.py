# %%
# coding=utf-8
import re

import modules.han_ji_chu_im as ji
import xlwings as xw


def main_run(CONVERT_FILE_NAME):
    # ==========================================================
    # Excel 檔案
    # ==========================================================
    # file_path = 'hoo-goa-chu-im.xlsx'
    file_path = CONVERT_FILE_NAME
    wb = xw.Book(file_path)

    source_sheet = wb.sheets['漢字注音表']
    end_row = source_sheet.range('A' + str(source_sheet.cells.last_cell.row)).end('up').row
    print(f'end_row = {end_row}')

    # %%
    # =========================================================
    # 檢查工作表是否已存在；若否：則建立
    # =========================================================

    def prepare_sheets(sheet_name_list):
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

    # %%
    # =========================================================
    # 依據指定的【注音／拼音方法】在工作表輸出含 Ruby Tags 的 HTML 網頁
    # =========================================================

    def main_process(target_sheet, zhu_im_method, div_class, rt_tag):
        # =========================================================
        # 輸出 <div> tag
        # =========================================================
        i = 1       # index for target sheet
        html_str = f"<div class='{div_class}'><p>"
        target_sheet.range('A' + str(i)).value = html_str
        i += 1

        row = 1     # index for source sheet
        end_counter = end_row + 1

        while row < end_counter:
            # 自 source_sheet 取待注音漢字
            han_ji = str(source_sheet.range('A' + str(row)).value)
            han_ji.strip()

            # =========================================================
            # 如是空白或換行，處理換行
            # =========================================================
            if han_ji == '' or han_ji == '\n':
                html_str = "</p><p>"
                target_sheet.range('A' + str(i)).value = html_str
                i += 1
                row += 1
                continue

            # =========================================================
            # 如只是標點符號，不必處理為漢字注音的工作
            # =========================================================
            piau_tiam = r"[；：？！\uFF0C\uFF08-\uFF09\u2013-\u2014\u2026\\u2018-\u201D\u3000\u3001-\u303F]"
            searchObj = re.search(piau_tiam, han_ji, re.M | re.I)
            if searchObj:
                # 將取到的「標點符號」，寫入目標工作表
                target_sheet.range('A' + str(i)).value = han_ji
                i += 1
                row += 1
                continue

            # =========================================================
            # 在字庫中查不到注音的漢字，略過注音處理
            # =========================================================
            chu_im_code = str(source_sheet.range('B' + str(row)).value).strip()
            if chu_im_code == 'None':
                # 讀到空白儲存格，視為使用者：「欲終止一個段落」；故於目標工作表寫入一個「換行」字元。
                chu_im = ''
                ruby_tag = f'  <ruby><rb>{han_ji}</rb><rp>(</rp><{rt_tag}>{chu_im}</{rt_tag}><rp>)</rp></ruby>'
                # if zhu_im_method == 'sip_ngoo_im':
                #     ruby_tag = f'  <ruby><rb>{han_ji}</rb><rp>(</rp><rtc>{chu_im}</rtc><rp>)</rp></ruby>'
                # else:
                #     ruby_tag  = f'  <ruby><rb>{han_ji}</rb><rp>(</rp><rt>{chu_im}</rt><rp>)</rp></ruby>'
                target_sheet.range('A' + str(i)).value = ruby_tag
                i += 1
                row += 1
                continue

            # =========================================================
            # 備妥注音時需參考用的資料
            # =========================================================
            # 取得聲母之聲母碼
            siann_bu = source_sheet.range('C' + str(row)).value
            if siann_bu.strip() != "":
                siann_index = ji.get_siann_idx(siann_bu)
                if siann_index == -1:
                    # 記錄沒找到之聲母
                    print(f"漢字：【{han_ji}】，找不到【聲母】：{siann_bu}！")

            # 取得韻母之韻母碼
            un_bu = source_sheet.range('D' + str(row)).value
            un_index = ji.get_un_idx(un_bu)
            if un_index == -1:
                # 記錄沒找到之韻母
                print(f"漢字：【{han_ji}】，找不到【韻母】：{un_bu}！")

            # 取得調號
            tiau_ho = int(source_sheet.range('E' + str(row)).value)

            # =========================================================
            # 使用注音碼，取得各式之〖 注音／拼音〗
            # =========================================================
            if zhu_im_method == 'sip_ngoo_im':
                # 輸出十五音
                chu_im = ji.get_sip_ngoo_im_chu_im(siann_index, un_index, tiau_ho)
            elif zhu_im_method == 'TPS':
                # 方音符號注音
                chu_im = ji.get_TPS_chu_im(siann_index, un_index, tiau_ho)
            elif zhu_im_method == 'POJ':
                # 輸出白話字拼音
                chu_im = ji.get_POJ_chu_im(siann_index, un_index, tiau_ho)
            elif zhu_im_method == 'TL':
                # 輸出羅馬拼音
                chu_im = ji.get_TL_chu_im(siann_index, un_index, tiau_ho)
            elif zhu_im_method == 'BP':
                chu_im = ji.get_BP_chu_im(siann_index, un_index, tiau_ho)
                # 輸出閩拼拼音
                # BP_chu_im1 = ji.get_BP_chu_im(siann_index, un_index, tiau_ho)

            # =========================================================
            # 將已注音之漢字加入【漢字注音表】
            # =========================================================
            # if zhu_im_method == 'sip_ngoo_im':
            #     ruby_tag = f'  <ruby><rb>{han_ji}</rb><rp>(</rp><rtc>{chu_im}</rtc><rp>)</rp></ruby>'
            # else:
            #     ruby_tag  = f'  <ruby><rb>{han_ji}</rb><rp>(</rp><rt>{chu_im}</rt><rp>)</rp></ruby>'
            ruby_tag = f'  <ruby><rb>{han_ji}</rb><rp>(</rp><{rt_tag}>{chu_im}</{rt_tag}><rp>)</rp></ruby>'
            target_sheet.range('A' + str(i)).value = ruby_tag

            # =========================================================
            # 調整讀取來源；寫入標的各工作表
            # =========================================================
            i += 1
            row += 1

        # =========================================================
        # 輸出 </div>
        # =========================================================
        html_str = "</p></div>"
        target_sheet.range('A' + str(i)).value = html_str

    # %%
    # ==========================================================
    # 設定輸出使用的注音方法
    # ==========================================================
    zhu_im_config = {
        'sip_ngoo_im': [
            'fifteen_yin',  # <div class="">
            'rt',         # Ruby Tag: <rt> / <rtc>
            '十五音注音',   # 輸出工作表名稱
        ],
        'TPS': [
            'zhu_yin',    # <div class="">
            'rtc',        # Ruby Tag: <rt> / <rtc>
            '方音符號注音',  # 輸出工作表名稱
        ],
        'POJ': [
            'pin_yin',    # <div class="">
            'rt',         # Ruby Tag: <rt> / <rtc>
            '白話字拼音',   # 輸出工作表名稱
        ],
        'TL': [
            'pin_yin',    # <div class="">
            'rt',         # Ruby Tag: <rt> / <rtc>
            '台羅拼音',    # 輸出工作表名稱
        ],
        'BP': [
            'pin_yin',    # <div class="">
            'rt',         # Ruby Tag: <rt> / <rtc>
            '閩拼拼音',    # 輸出工作表名稱
        ],
    }

    # %%
    # ==========================================================
    # 備妥程式需使用之工作表
    # ==========================================================
    for zhu_im_method in zhu_im_config:
        # CONFIG_IDX = 3
        # zhu_im_method = list(zhu_im_config)[CONFIG_IDX]
        div_class = zhu_im_config[zhu_im_method][0]
        rt_tag = zhu_im_config[zhu_im_method][1]
        zhu_im_sheet_name = zhu_im_config[zhu_im_method][2]

        # -----------------------------------------------------
        # 檢查工作表是否已存在
        # sheet_name_list.append(zhu_im_sheet_name)
        prepare_sheets([zhu_im_sheet_name])
        chu_im_sheet = wb.sheets[zhu_im_sheet_name]
        # -----------------------------------------------------
        main_process(chu_im_sheet, zhu_im_method, div_class, rt_tag)
