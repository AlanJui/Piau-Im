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

from mod_Query_for_Tshiat_Gu import query_tshiat_gu_siong_ji, query_tshiat_gu_ha_ji
from mod_huan_tshiat import tshu_tiau, query_tiau_ho
from mod_於字典網站查詢漢字之廣韻切語發音 import fetch_kong_un_info


def fetch_arg():
    # 取得命令行參數
    cmd_arg = sys.argv[1:]  # 取得所有除腳本名稱之外的命令行參數

    # 檢查 cmd_arg 是否有內容
    if not cmd_arg:  # 如果沒有傳入任何參數
        print("沒有傳入任何參數，使用預設參數。")

    # 若使用者未輸入欲查詢之漢字，則賦予預設值
    han_ji = cmd_arg[0] if len(cmd_arg) > 0 else "詼"
    print(f"han_ji = {han_ji}")

    return han_ji

# 程式作業流程：
# 1. 欲查詢之漢字，如：詼；
# 2. 反切雙字，如：苦回；
# 3. 廣韻查詢索引：廣韻·上平聲·灰·恢；
# 4. 自反切雙字分離出：反切上字、反切下字，如：反切上字=苦、反切下字=回。
if __name__ == "__main__":
    # 取得使用者輸入的參數
    han_ji = fetch_arg()

    # 自廣韻線上字典，取得反切語與四聲調類
    # 查詢回傳的結果是一個字典陣列，每個字典包含了反切語、四聲調類、韻系、反切下字等資訊
    # tshiat_gu_item = {
    #     "tshiat_gu": tshiat_gu,  # 切語：苦回
    #     "tiau": tiau,            # 調：平/上/去/入的四聲調類
    #     "un_he": un_he,          # 韻系：灰
    #     "tshia_gu_ha_ji": tshia_gu_ha_ji,  # 切語下字：恢
    # }
    tshiat_gu_list = fetch_kong_un_info(han_ji)

    for item in tshiat_gu_list:
        # 取得切語
        tshiat_gu = item["tshiat_gu"]

        # 查詢切語上字
        siong_ji = item["tshiat_gu"][0]
        tshiat_gu_siong_ji = query_tshiat_gu_siong_ji(siong_ji)
        tai_lo_siann_bu = tshiat_gu_siong_ji[0]["tai_lo"]
        tshing_tok = tshiat_gu_siong_ji[0]["tshing_tok"]

        # 查詢切語下字
        ha_ji = item["tshiat_gu"][1]
        tshiat_gu_ha_ji = query_tshiat_gu_ha_ji(ha_ji)
        tai_lo_un_bu = tshiat_gu_ha_ji[0]["tai_lo"]

        # 取得四聲調類
        tiau_lui = item["tiau"]
        su_sing = tshu_tiau(tiau_lui)
        # 查詢反切語之四聲八調之調號
        tiau_ho = query_tiau_ho(tshing_tok, su_sing)

        # 組合台羅拼音
        tai_lo_phing_im = f"{tai_lo_siann_bu}{tai_lo_un_bu}{tiau_ho}"
        
        #=======================================================
        # 顯示查詢結果
        #=======================================================
        print("\n===================================================")
        print(f"查詢漢字：{han_ji}\t廣韻切語為: {tshiat_gu}\t台羅拼音為: {tai_lo_phing_im}")
        print(f"反切上字：{siong_ji}\t得聲母台羅拼音為: {siong_ji}\t分清濁為：{tshing_tok}")
        print(f"反切下字：{ha_ji}\t得韻母台羅拼音為: {ha_ji}\t辨四聲為：{su_sing}聲")
        print(f"依分清濁與辨四聲，得聲調為：{tshing_tok[-1]}{su_sing}，即：台羅四聲八調之第 {tiau_ho} 調")
        # if not sheet.range('D7').value == "找不到":
        #     print(f"依分清濁與辨四聲，得聲調為：{sheet.range('E7').value}，即：台羅四聲八調之第 {int(sheet.range('D7').value)} 調")



