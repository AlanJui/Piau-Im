import sys

import xlwings as xw


def export_han_ji_to_txt(wb, sheet_name='漢字注音', output_file='tmp.txt'):
    """
    將 Excel 工作表中指定區域的漢字取出，儲存為一個純文字檔。
    """
    # 選擇工作表
    sheet = wb.sheets[sheet_name]
    sheet.activate()

    # 初始化儲存字串
    han_ji_text = ""

    # 取得總列數與每列總字數
    TOTAL_LINES = int(wb.names['每頁總列數'].refers_to_range.value)
    CHARS_PER_ROW = int(wb.names['每列總字數'].refers_to_range.value)

    # 設定起始及結束的欄位（【D欄=4】到【R欄=18】）
    row = 5
    start_col = 4
    end_col = start_col + CHARS_PER_ROW

    # 從第 5 列、9 列、13 列等列取出漢字，並組合成純文字
    line_text = ""
    end_of_file = False
    line = 1
    while line < TOTAL_LINES:
        # 設定【作用儲存格】為列首
        sheet.range((row, 1)).select()
        # 每列逐欄取出漢字
        for col in range(start_col, end_col):
            cell_value = sheet.range((row, col)).value

            if cell_value == 'φ':
                end_of_file = True
                break
            elif cell_value == '\n':
                line_text += '\n'
                break
            else:
                line_text += cell_value

        # 若該列為空白列，則識作【檔案結尾（EOF）】
        if end_of_file:
            # han_ji_text += '\nEOF\n'
            print(f"第 {row} 列為檔案結尾處，結束處理作業。")
            break

        # 輸出當前行處理的內容
        # print(f"第 {row} 列的輸出內容：")
        # print(line_text)

        # 每處理 15 個字元後，換到下一行
        row += 4
        line += 1

    # 將所有漢字寫入文字檔
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(line_text)

    print(f"已成功將漢字輸出至檔案：{output_file}")


def dump_txt_file(file_path):
    """
    在螢幕 Dump 純文字檔內容。
    """
    print("\n【文字檔內容】：")
    print("========================================\n")
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
            print(content)
    except FileNotFoundError:
        print(f"無法找到檔案：{file_path}")


# 主函數示例
if __name__ == "__main__":
    # 開啟 Excel 工作簿
    # app = xw.App(visible=False)
    # wb = app.books.open('your_excel_file.xlsx')  # 修改為實際 Excel 檔案名稱
    wb = None
    # 使用已打開且處於作用中的 Excel 工作簿
    try:
        # 嘗試獲取當前作用中的 Excel 工作簿
        wb = xw.apps.active.books.active
    except Exception as e:
        print(f"發生錯誤: {e}")
        print("無法找到作用中的 Excel 工作簿")
        sys.exit(2)

    if not wb:
        print("無法執行，可能原因：(1) 未指定輸入檔案；(2) 未找到作用中的 Excel 工作簿")
        sys.exit(2)

    # 設定純文字檔案名稱
    output_file = 'tmp.txt'

    # 呼叫函數將漢字導出為純文字檔
    export_han_ji_to_txt(wb, output_file=output_file)

    # 螢幕 Dump 檔案內容
    dump_txt_file(output_file)

    # 關閉工作簿和應用程式
    # wb.close()
