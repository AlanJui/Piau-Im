import xlwings as xw
import settings

# ===========================================================================
# (1) 取得需要注音的「檔案名稱」及其「目錄路徑」。
# ===========================================================================
# 取得 Input 檔案名稱
file_path = settings.get_input_file_path()
if not file_path:
    print("未設定 .env 檔案")
    CONVERT_FILE_NAME = "hoo-goa-chu-im.xlsx"
else:
    CONVERT_FILE_NAME = file_path
print(f"CONVERT_FILE_NAME = {CONVERT_FILE_NAME}")

file_path = CONVERT_FILE_NAME
wb = xw.Book(file_path)

# ===========================================================================
# (2)
# ===========================================================================
# 取得「漢字注音表」的總列數
source_sheet = wb.sheets["漢字注音表"]
source_sheet.select()
# sheet_name = source_sheet.name
sheet_name = wb.sheets.active.name
print(f"sheet_name = {sheet_name}")

# end_of_row_no = (
#     source_sheet.range("A" + str(source_sheet.cells.last_cell.row)).end("up").row
# )
last_row = source_sheet.cells.last_cell.row
end_of_row_no = source_sheet.range("A" + str(last_row)).end("up").row
print(f"end_of_row_no = {end_of_row_no}")

# ===========================================================================
# (3)
# ===========================================================================
# 通过命名单元格的名称获取值
# file_name = wb.sheets["env"].names["FILE_NAME"].refers_to_range.value
# file_name = xw.Range("FILE_NAME").value
# file_name = wb.sheets["env"].range("FILE_NAME").value
# print(f"file_name = {file_name}")

# ws = wb.sheets["env"]
# ws.select()
# wb.sheets["env"].activate()
# title = xw.Range("TITLE").value
# image_url = xw.Range("IMAGE_URL").value
title = wb.sheets["env"].range("TITLE").value
image_url = wb.sheets["env"].range("IMAGE_URL").value
sheet_name = source_sheet.select().name

# 定义文章标题和注音/拼音方法
methods = [
    "十五音標音",
    "方音符號注音",
    "白話字拼音",
    "台羅拼音",
    "閩拼拼音",
]

# ruff: noqa: E501
div_tag = (
    "《%s》【%s】\n"
    '<div class="separator" style="clear: both">\n'
    '  <a href="圖片" style="display: block; padding: 1em 0; text-align: center">\n'
    '    <img alt="" border="0" width="400" data-original-height="630" data-original-width="1200"\n'
    '      src="%s" />\n'
    "  </a>\n"
    "</div>\n"
    "\n"
)

output = div_tag % (title, sheet_name, image_url)
print(output)

# 製作每種注音/拼音方法的 HTML Tags
# for method in methods:
#     output = div_tag % (title, method, image_url)
#     print(output)
