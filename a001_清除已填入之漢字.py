import os
import sys
from pathlib import Path

import xlwings as xw

# =========================================================================
# 取得專案根目錄。
# =========================================================================
# 獲取當前檔案的路徑
current_file_path = Path(__file__).resolve()

# 專案根目錄
project_root = current_file_path.parent

print(f"專案根目錄為: {project_root}")

# =========================================================================
# 若無指定輸入檔案，則獲取當前作用中的 Excel 檔案並另存新檔。
# =========================================================================
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

# 選擇指定的工作表
sheet = wb.sheets['漢字注音']   # 選擇工作表
sheet.activate()  # 將「漢字注音」工作表設為作用中工作表
sheet.range('A1').select()     # 將 A1 儲存格設為作用儲存格

# 每頁最多處理的列數
TOTAL_ROWS = int(wb.names['每頁總列數'].refers_to_range.value)  # 從名稱【每頁總列數】取得值
# 每列最多處理的字數
CHARS_PER_ROW = int(wb.names['每列總字數'].refers_to_range.value)  # 從名稱【每列總字數】取得值

#--------------------------------------------------------------------------
# 將儲存格內的舊資料清除
#--------------------------------------------------------------------------
end_of_rows = int((TOTAL_ROWS * CHARS_PER_ROW ) + 2)
cells_range = f'D3:R{end_of_rows}'

sheet.range(cells_range).clear_contents()     # 清除 C3:R{end_of_row} 範圍的內容

