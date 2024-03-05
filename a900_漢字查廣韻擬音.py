#==============================================================================
# 透過 Excel ，使用廣韻的反切方法，查找漢字的羅馬拼音。
# 
# 操作方式：
#  - 輸入欲查詢羅馬拼音之漢字；
#  - 輸入廣韻的查找索引資料。
# 
#   ipython a900_廣韻反切查羅馬拼音.ipynb 攝 "書涉 (《廣韻·入聲·葉·攝》)"
#==============================================================================

import getopt
import sys
import xlwings as xw

from a910_於字典網站查詢漢字之廣韻切語發音 import fetch_kong_un_info


# 接收使用者輸入的 "反切" 查詢參數
# 根據傳入的 siann_lui 參數取出 "聲" 字左邊的一個中文字
# "調類" siann_lui 可能值：上平聲、下平聲、上聲、去聲、入聲
def tshu_tiau(tiau_lui):
    # 永遠取出 "聲" 字左邊的一個中文字
    return tiau_lui[tiau_lui.find("聲")-1]

def fetch_arg():
    # 取得命令行參數
    cmd_arg = sys.argv[1:]  # 取得所有除腳本名稱之外的命令行參數

    # 檢查 cmd_arg 是否有內容
    if not cmd_arg:  # 如果沒有傳入任何參數
        print("沒有傳入任何參數，使用預設參數。")
    else:
        # 無論是否使用預設參數，都遍歷 cmd_arg 中的每個元素
        for i, arg in enumerate(cmd_arg, start=1):
            print(f"參數 {i}: {arg}")

    # 根據獲取的 cmd_arg 分別賦值
    han_ji = cmd_arg[0] if len(cmd_arg) > 0 else "詼"
    # kong_un_huan_tshiat = cmd_arg[1] if len(cmd_arg) > 1 else "苦回(《廣韻·上平聲·灰·恢》)"
    kong_un_huan_tshiat = cmd_arg[1] if len(cmd_arg) > 1 else ""

    print(f"han_ji = {han_ji}")
    print(f"kong_un_huan_tshiat = {kong_un_huan_tshiat}")

    return {
        "han_ji": han_ji,
        "kong_un_huan_tshiat": kong_un_huan_tshiat,
    }

# 根據傳入的廣韻查詢索引，取出 "反切語" 與 "四聲調類" 
def fetch_tshiat_gu_tiau_lui(kong_un_huan_tshiat):
    # 分離 "苦回" 與 "廣韻·上平聲·灰·恢"
    tshiat_gu, kong_un_with_brackets = kong_un_huan_tshiat.split('(')
    tshiat_gu = tshiat_gu.strip()  # 清除前後的空白

    siong_ji = tshiat_gu[0]  # 取反切之上字：即反切的首字
    ha_ji = tshiat_gu[1] if len(tshiat_gu) > 1 else ""  # 取反切之下字：即反切的第二個字符，如果有的話

    # 移除結尾的 "》)"
    kong_un_khi_bue = kong_un_with_brackets[:-2]  
    # 移除 "《" 並重新分離 "廣韻·上平聲·灰·恢"
    kong_un_cleaned = kong_un_khi_bue[1:]  # 移除開頭的 "《"

    # 將 "廣韻·上平聲·灰·恢" 依 "·" 切分成有 4 個元素的字串陣列
    kong_un = kong_un_cleaned.split('·')

    # 分離 "廣韻·上平聲·灰·恢" 中的 "上平聲"
    tiau_lui = kong_un[1]

    # 取四聲調之調類
    siann_tiau = tshu_tiau(tiau_lui)  

    return {
        "siong_ji": siong_ji,
        "ha_ji": ha_ji,
        "siann_tiau": siann_tiau,
    }


# 程式作業流程：
# 1. 欲查詢之漢字，如：詼；
# 2. 反切雙字，如：苦回；
# 3. 廣韻查詢索引：廣韻·上平聲·灰·恢；
# 4. 自反切雙字分離出：反切上字、反切下字，如：反切上字=苦、反切下字=回。
if __name__ == "__main__":
    # 取得使用者輸入的參數
    args = fetch_arg()
    han_ji = args["han_ji"]
    kong_un_huan_tshiat = args["kong_un_huan_tshiat"]

    # 取得反切語與四聲調類
    if not kong_un_huan_tshiat and han_ji:
        # 僅只輸入漢字，未輸入廣韻查詢索引
        print("請輸入廣韻查詢索引。")
        sys.exit(1)
    else:
        tshiat_gu_tiau_lui = fetch_tshiat_gu_tiau_lui(kong_un_huan_tshiat)
        siong_ji = tshiat_gu_tiau_lui["siong_ji"]
        ha_ji = tshiat_gu_tiau_lui["ha_ji"]
        siann_tiau = tshiat_gu_tiau_lui["siann_tiau"]

    ## 開啟可執行反切查羅馬拼音之活頁簿檔案
    # 1. 開啟 Excel 活頁簿檔案： .\tools\廣韻反切查音工具.xlsx ；
    # 2. 擇用 "反切" 工作表。
    # 
    # [程式規格]：
    #  - 使用 xlwings 套件，操作 Excel 檔案；
    #  - 以上兩步的作業程序，都用 try: exception: 形式執行，遇有意外事件發生時，於畫面顯示問題狀況，然後終止程式的繼續執行。

    try:
        # 指定 Excel 檔案路徑
        file_path = r'.\\tools\\廣韻反切查音工具.xlsx'
        
        # 使用 xlwings 開啟 Excel 檔案
        wb = xw.Book(file_path)
        
        # 選擇名為 "反切" 的工作表
        sheet = wb.sheets['反切']
        
        # 將變數值填入指定的儲存格
        sheet.range('C2').value = han_ji
        sheet.range('D2').value = kong_un_huan_tshiat
                
        # 從 D8 儲存格取出值，存放於變數 tai_lo_phing_im
        tai_lo_phing_im = sheet.range('D8').value
        
        #=======================================================
        # 顯示查詢結果
        #=======================================================
        print("\n===================================================")
        print(f"查詢漢字：{han_ji}\t廣韻反切為: {kong_un_huan_tshiat}")
        print(f"反切上字：{siong_ji}\t得聲母台羅拼音為: {sheet.range('D5').value}\t分清濁為：{sheet.range('E5').value}")
        print(f"反切下字：{ha_ji}\t得韻母台羅拼音為: {sheet.range('D6').value}\t辨四聲為：{sheet.range('E6').value}聲")
        print(f"依分清濁與辨四聲，得聲調為：{sheet.range('E7').value}，即：台羅四聲八調之第 {int(sheet.range('D7').value)} 調")
        print(f"漢字：{han_ji}\t台羅拼音為: {tai_lo_phing_im}")

    except Exception as e:
        # 如果遇到任何錯誤，顯示錯誤信息並終止程式
        print(f"發生錯誤：{e}")



