import os
import sys
from pathlib import Path

import xlwings as xw

from mod_file_access import get_han_ji_khoo, get_sound_type, save_as_new_file
from p701_Clear_Cells import clear_han_ji_kap_piau_im
from p702_Ca_Han_Ji_Thak_Im import ca_han_ji_thak_im
from p709_reset_han_ji_cells import reset_han_ji_cells
from p710_thiam_han_ji import fill_hanji_in_cells

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

# 將儲存格已填入之漢字及標音清除
clear_han_ji_kap_piau_im(wb)

# 將待注音的【漢字儲存格】，文字顏色重設為黑色（自動 RGB: 0, 0, 0）；填漢顏色重設為無填滿
reset_han_ji_cells(wb)

# 將待注音的漢字填入
fill_hanji_in_cells(wb)

# A731: 自動為漢字查找讀音，並抄寫到漢字的上方(拼音)及下方(注音)。
type = get_sound_type(wb)
han_ji_khoo = get_han_ji_khoo(wb)
if han_ji_khoo == "河洛話" and type == "白話音":
    ca_han_ji_thak_im(wb, sheet_name='漢字注音', cell='V3', hue_im=type, han_ji_khoo="河洛話", db_name='Ho_Lok_Ue.db', module_name='mod_河洛話', function_name='han_ji_ca_piau_im')
elif han_ji_khoo == "河洛話" and type == "文讀音":
    ca_han_ji_thak_im(wb, sheet_name='漢字注音', cell='V3', hue_im=type, han_ji_khoo="河洛話", db_name='Ho_Lok_Ue.db', module_name='mod_河洛話', function_name='han_ji_ca_piau_im')
elif han_ji_khoo == "廣韻":
    ca_han_ji_thak_im(wb, sheet_name='漢字注音', cell='V3', hue_im="文讀音", han_ji_khoo="廣韻", db_name='Kong_Un.db', module_name='mod_廣韻', function_name='han_ji_ca_piau_im')
else:
    print("無法執行漢字標音作業，請確認【env】工作表【語音類型】及【漢字庫】欄位的設定是否正確！")
    # sys.exit(2)

# 將檔案存放路徑設為【專案根目錄】之下
save_as_new_file(wb=wb)