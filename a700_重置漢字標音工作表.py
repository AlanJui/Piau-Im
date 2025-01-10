import os
import sys
from pathlib import Path

import xlwings as xw

from p709_reset_han_ji_cells import reset_han_ji_cells

# =========================================================================
# (1) 取得專案根目錄。
# =========================================================================
# 獲取當前檔案的路徑
current_file_path = Path(__file__).resolve()

# 專案根目錄
project_root = current_file_path.parent

print(f"專案根目錄為: {project_root}")

# =========================================================================
# (2) 若無指定輸入檔案，則獲取當前作用中的 Excel 檔案並另存新檔。
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

#--------------------------------------------------------------------------
# 將儲存格內的舊資料清除
#--------------------------------------------------------------------------
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

#--------------------------------------------------------------------------
# 將待注音的【漢字儲存格】，文字顏色重設為黑色（自動 RGB: 0, 0, 0）；填漢顏色重設為無填滿
#--------------------------------------------------------------------------
reset_han_ji_cells(wb)

# 儲存新建立的工作簿
wb.save()

print(f"【漢字注音】工作表已清空、重置！")
#--------------------------------------------------------------------------
# 另存新檔
#--------------------------------------------------------------------------
# 將待注音的【漢字儲存格】，文字顏色重設為黑色（自動 RGB: 0, 0, 0）；填漢顏色重設為無填滿
# # 將檔案存放路徑設為【專案根目錄】之下
# try:
#     file_name = str(wb.names['TITLE'].refers_to_range.value).strip()
# except KeyError:
#     print("未找到命名範圍 'TITLE'，使用預設名稱")
#     # file_name = "Tai_Gi_Zu_Im_Bun.xlsx"   # 提供一個預設檔案名稱
#     setting_sheet = wb.sheets["env"]
#     file_name = str(
#         setting_sheet.range("C4").value
#     ).strip()

# # 設定檔案輸出路徑，存於專案根目錄下的 output2 資料夾
# output_path = wb.names['OUTPUT_PATH'].refers_to_range.value
# new_file_path = os.path.join(
#     ".\\{0}".format(output_path),
#     f"【河洛話注音】{file_name}.xlsx")

# # 儲存新建立的工作簿
# wb.save(new_file_path)

# print(f"待注音漢字已備妥： {new_file_path}")
