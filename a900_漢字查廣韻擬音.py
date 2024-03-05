#==============================================================================
# 透過 Excel ，使用廣韻的反切方法，查找漢字的羅馬拼音。
# 
# 操作方式：
#  - 輸入欲查詢羅馬拼音之漢字；
#  - 輸入廣韻的查找索引資料。
# 
#   ipython a900_廣韻反切查羅馬拼音.ipynb 攝 "書涉 (《廣韻·入聲·葉·攝》)"
#==============================================================================

import sys
import xlwings as xw

from a910_於字典網站查詢漢字之廣韻切語發音 import fetch_kong_un_info


def fetch_arg():
    # 取得命令行參數
    cmd_arg = sys.argv[1:]  # 取得所有除腳本名稱之外的命令行參數

    # 檢查 cmd_arg 是否有內容
    if not cmd_arg:  # 如果沒有傳入任何參數
        print("沒有傳入任何參數，使用預設參數。")
    # else:
    #     # 無論是否使用預設參數，都遍歷 cmd_arg 中的每個元素
    #     for i, arg in enumerate(cmd_arg, start=1):
    #         print(f"參數 {i}: {arg}")
    # else:
    #     # 如果沒有傳入任何參數，則顯示提示信息，然後終止程式
    #     print("請輸入欲查詢《廣韻》切語之漢字。")
    #     sys.exit(1)

    # 若使用者未輸入欲查詢之漢字，則賦予預設值
    han_ji = cmd_arg[0] if len(cmd_arg) > 0 else "詼"
    print(f"han_ji = {han_ji}")

    return han_ji

# 接收使用者輸入的 "反切" 查詢參數
# 根據傳入的 siann_lui 參數取出 "聲" 字左邊的一個中文字
# "調類" siann_lui 可能值：上平聲、下平聲、上聲、去聲、入聲
def tshu_tiau(tiau_lui):
    # 永遠取出 "聲" 字左邊的一個中文字
    return tiau_lui[tiau_lui.find("聲")-1]


# 程式作業流程：
# 1. 欲查詢之漢字，如：詼；
# 2. 反切雙字，如：苦回；
# 3. 廣韻查詢索引：廣韻·上平聲·灰·恢；
# 4. 自反切雙字分離出：反切上字、反切下字，如：反切上字=苦、反切下字=回。
if __name__ == "__main__":
    # 取得使用者輸入的參數
    han_ji = fetch_arg()

    # 取得反切語與四聲調類
    tshiat_gu_list = fetch_kong_un_info(han_ji)

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
        for item in tshiat_gu_list:
            # 將漢字之切語填入指定的儲存格    
            sheet.range('D2').value = item["tshiat_gu"]
            # 將切語上字填入指定的儲存格
            sheet.range('C5').value = siong_ji = item["tshiat_gu"][0]
            # 將切語下字填入指定的儲存格
            sheet.range('C6').value = ha_ji = item["tshiat_gu"][1]
            # 將切語下字所屬之四聲調類填入指定的儲存格    
            sheet.range('E6').value = tshu_tiau(item["tiau"]) 
                
            # 從 D8 儲存格取出值，存放於變數 tai_lo_phing_im
            tai_lo_phing_im = sheet.range('D8').value
            
            #=======================================================
            # 顯示查詢結果
            #=======================================================
            print("\n===================================================")
            print(f"查詢漢字：{han_ji}\t廣韻切語為: {item['tshiat_gu']}\t台羅拼音為: {tai_lo_phing_im}")
            print(f"反切上字：{siong_ji}\t得聲母台羅拼音為: {sheet.range('D5').value}\t分清濁為：{sheet.range('E5').value}")
            print(f"反切下字：{ha_ji}\t得韻母台羅拼音為: {sheet.range('D6').value}\t辨四聲為：{sheet.range('E6').value}聲")
            if not sheet.range('D7').value == "找不到":
                print(f"依分清濁與辨四聲，得聲調為：{sheet.range('E7').value}，即：台羅四聲八調之第 {int(sheet.range('D7').value)} 調")

    except Exception as e:
        # 如果遇到任何錯誤，顯示錯誤信息並終止程式
        print(f"發生錯誤：{e}")



