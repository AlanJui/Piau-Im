import os
import sys
from pathlib import Path

import xlwings as xw

# 指定虛擬環境的 Python 路徑
venv_python = os.path.join(".venv", "Scripts", "python.exe") if sys.platform == "win32" else os.path.join(".venv", "bin", "python")

# (0) 取得專案根目錄。
# 使用已打開且處於作用中的 Excel 工作簿
try:
    wb = xw.apps.active.books.active
except Exception as e:
    print(f"發生錯誤: {e}")
    print("無法找到作用中的 Excel 工作簿")
    sys.exit(2)

# 獲取活頁簿的完整檔案路徑
file_path = wb.fullname
print(f"完整檔案路徑: {file_path}")

# 獲取活頁簿的檔案名稱（不包括路徑）
file_name = wb.name
print(f"檔案名稱: {file_name}")

# 獲取當前檔案的路徑
current_file_path = Path(file_path).resolve()

# 專案根目錄
working_dir_path = current_file_path.parent
print(f"專案根目錄為: {working_dir_path}")

# (1) 存成作業暫存檔
new_file_name = "working"

# 設定檔案輸出路徑，存於專案根目錄下的 output2 資料夾
new_file_path = os.path.join(
    working_dir_path, 
    f"【河洛話注音】{new_file_name}.xlsx")

# 儲存新建立的工作簿
wb.save(new_file_path)
print(f"作業中暫存檔名: {wb.name}")

# (2) 將儲存格內的舊資料清除
sheet = wb.sheets['漢字注音']   # 選擇工作表
sheet.activate()               # 將「漢字注音」工作表設為作用中工作表
sheet.range('A1').select()     # 將 A1 儲存格設為作用儲存格

total_rows = wb.names['每頁總列數'].refers_to_range.value
cells_per_row = 4
end_of_rows = int((total_rows * cells_per_row ) + 2) 
cells_range = f'D3:R{end_of_rows}'

sheet.range(cells_range).clear_contents()     # 清除 C3:R{end_of_row} 範圍的內容

# 獲取 V3 儲存格的合併範圍
merged_range = sheet.range('V3').merge_area
# 清空合併儲存格的內容
merged_range.clear_contents()

# 顯示「已輸入之拼音字母及注音符號」 
named_range = wb.names['顯示注音輸入']  # 選擇名為 "顯示注音輸入" 的命名範圍# 選擇名為 "顯示注音輸入" 的命名範圍
named_range.refers_to_range.value = True

# 設定 V3 儲存格為作用儲存格
sheet.range('V3').select()