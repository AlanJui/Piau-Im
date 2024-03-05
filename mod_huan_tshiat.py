import xlwings as xw

def siong_ji_tsa_siann_bu(siong_ji):
    try:
        # 使用 xlwings 打開 Excel 檔案
        file_path = r'.\\tools\\反切上字與聲母對映表.xlsx'
        wb = xw.Book(file_path)
        sheet = wb.sheets['反切上字表']  # 假設工作表的名稱為 "反切上字"
        
        # 查找 siong_ji 在 "反切上字" 欄 (G2:G39) 的位置
        for row in range(2, 40):  # 假設資料從第2行開始，到第39行
            cell_value = sheet.range(f'G{row}').value
            if cell_value:
                words = cell_value.split()  # 按空格拆分儲存格的值
                if siong_ji in words:
                    # 如果找到了 siong_ji，則記錄下其相對映之 "列號"，並從相關欄位獲取資料
                    data = {
                        "id": row - 1,  # 第一行為標題行，故實際列號需要減1
                        "lui": sheet.range(f'B{row}').value,
                        "siann_bu": sheet.range(f'C{row}').value,
                        "tshing_lo": sheet.range(f'F{row}').value,
                        "tai_lo": sheet.range(f'D{row}').value,
                        "IPA": sheet.range(f'E{row}').value,
                    }
                    wb.close()  # 處理完成後關閉工作簿
                    return data
        wb.close()  # 如果沒找到，也關閉工作簿
        return None  # 如果在範圍內沒有找到 siong_ji，返回 None
    except Exception as e:
        print(f"發生錯誤：{e}")
        return None
    

def e_ji_tsa_un_bu(e_ji):
    try:
        # 使用 xlwings 打開 Excel 檔案
        file_path = r'.\\tools\\反切下字與韻母對映表.xlsx'
        wb = xw.Book(file_path)
        sheet = wb.sheets['反切下字表']  
        
        # 查找 e_ji 在指定範圍 (J2:J187) 的位置
        for row in range(2, 188):  # 資料從第2列開始，到第187列
            cell_value = sheet.range(f'J{row}').value
            if cell_value:
                words = cell_value.split()  # 按空格拆分儲存格的值
                if e_ji in words:
                    # 如果找到了 e_ji，則記錄下其相對映之 "列號"，並從相關欄位獲取資料
                    data = {
                        "id": row - 1,  # 第一行為標題行，故實際列號需要減1
                        "liap": sheet.range(f'B{row}').value, # 攝
                        "un_he": sheet.range(f'C{row}').value,  # 韻系
                        "si_siann": sheet.range(f'D{row}').value,  # 四聲 (平上去入)
                        "ting_de": sheet.range(f'E{row}').value,  # 等第
                        "khai_hap": sheet.range(f'F{row}').value,  # 開合
                        "un_bu": sheet.range(f'G{row}').value,  # 韻母
                        "tai_lo": sheet.range(f'H{row}').value,  # 臺羅拼音
                        "IPA": sheet.range(f'I{row}').value,  # 國際音標
                    }
                    wb.close()  # 處理完成後關閉工作簿
                    return data
        wb.close()  # 如果沒找到，也關閉工作簿
        return None  # 如果在範圍內沒有找到 e_ji，返回 None
    except Exception as e:
        print(f"發生錯誤：{e}")
        return None

# 接收使用者輸入的 "反切" 查詢參數
# 根據傳入的 siann_lui 參數取出 "聲" 字左邊的一個中文字
# "調類" siann_lui 可能值：上平聲、下平聲、上聲、去聲、入聲
def tshu_tiau(tiau_lui):
    # 永遠取出 "聲" 字左邊的一個中文字
    return tiau_lui[tiau_lui.find("聲")-1]

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

if __name__ == "__main__":
    # 測試 siong_ji_tsa_siann_bu 函數
    siong_ji = "普"
    result_siong_ji = siong_ji_tsa_siann_bu(siong_ji)
    assert result_siong_ji["siann_bu"] == "滂"
    assert result_siong_ji["tai_lo"] == "ph"
    assert result_siong_ji["tshing_lo"] == "次清"
    print(f"\n查反切上字：{siong_ji}")
    print(f"siong_ji_tsa_siann_bu 測試結果：{result_siong_ji}")

    # 測試 e_ji_tsa_un_bu 函數
    e_ji = "荅"
    result_e_ji = e_ji_tsa_un_bu(e_ji)
    assert result_e_ji["un_bu"] == "合"
    assert result_e_ji["tai_lo"] == "ap"
    print(f"\n查反切下字：{e_ji}")
    print(f"e_ji_tsa_un_bu 測試結果：{result_e_ji}")
