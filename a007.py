import xlwings as xw

# 開啟 Excel 應用程式和工作簿
# app = xw.App(visible=False)  # 設置為不可見模式
# wb = app.books.open('Template.xlsx')  # 打開 Excel 文件
wb = xw.apps.active.books.active
sheet = wb.sheets['漢字注音']  # 選擇【漢字注音】工作表

# 從 env 工作表中獲取每頁總列數和每列總字數
env_sheet = wb.sheets['env']
total_lines = int(env_sheet.range('每頁總列數').value)
chars_per_row = int(env_sheet.range('每列總字數').value)

# 設定起始及結束的【列】位址
ROWS_PER_LINE = 4
start_row = 5
end_row = start_row + (total_lines * ROWS_PER_LINE)

# 設定起始及結束的【欄】位址
start_col = 4  # D 欄
end_col = start_col + chars_per_row - 1  # 因為欄位是從 1 開始計數

# 定義儲存格格式
def set_range_format(range_obj, font_name, font_size, font_color, fill_color=None):
    range_obj.api.Font.Name = font_name
    range_obj.api.Font.Size = font_size
    range_obj.api.Font.Color = font_color
    if fill_color:
        range_obj.api.Interior.Color = fill_color
    else:
        range_obj.api.Interior.Pattern = xw.constants.Pattern.xlPatternNone  # 無填滿

# 清除內容並設置格式
for row in range(start_row, end_row + 1, ROWS_PER_LINE):
    line = ((row - start_row) // ROWS_PER_LINE) + 1
    print(f'重置第 {line} 行，共 {ROWS_PER_LINE} 列之儲存格內容！')
    # 人工標音
    range_人工標音 = sheet.range((row - 2, start_col), (row - 2, end_col))
    range_人工標音.value = None
    set_range_format(range_人工標音, font_name='Arial', font_size=24, font_color=0xFF0000)  # 紅色

    # 台語音標
    range_台語音標 = sheet.range((row - 1, start_col), (row - 1, end_col))
    range_台語音標.value = None
    set_range_format(range_台語音標, font_name='Sitka Text Semibold', font_size=24, font_color=0xFF9933)  # 橙色

    # 漢字
    range_漢字 = sheet.range((row, start_col), (row, end_col))
    range_漢字.value = None
    set_range_format(range_漢字, font_name='吳守禮細明台語注音', font_size=48, font_color=0x000000)  # 黑色

    # 漢字標音
    range_漢字標音 = sheet.range((row + 1, start_col), (row + 1, end_col))
    range_漢字標音.value = None
    set_range_format(range_漢字標音, font_name='芫荽 0.94', font_size=26, font_color=0x009900)  # 綠色

# 保存文件
# wb.save('漢字注音_reset.xlsx')
# wb.close()
# app.quit()