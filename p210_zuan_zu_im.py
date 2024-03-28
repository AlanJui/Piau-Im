#==========================================================
# 模組功能：提供 "注音轉換功能" ，令聲母、韻母及聲調，能於不同的注音方法
# 之間相互轉換。
# 
# 支援的注音方法有：國際音標、白話字、台羅、閩拼、方音符號、雅俗通十五音。
#==========================================================

import sqlite3


def init_siann_bu_dict():
    # 建立連接
    conn = sqlite3.connect('Nga_Siok_Thong.db')

    # 建立游標物件
    cursor = conn.cursor()

    # 執行 SQL 查詢
    cursor.execute("SELECT * FROM 聲母對照表")

    # 獲取所有資料
    rows = cursor.fetchall()

    # 初始化字典
    siann_bu_dict = {}

    # 從查詢結果中提取資料並將其整理成一個字典
    for row in rows:
        siann_bu_dict[row[1]] = {
            'ipa': row[2],
            'poj': row[3],
            'tl': row[4],
            'bp': row[5],
            'tps': row[6],
            'sni': row[7],
        }

    # 關閉連接
    conn.close()

    return siann_bu_dict


def init_un_bu_dict():
    # 建立連接
    conn = sqlite3.connect('Nga_Siok_Thong.db')

    # 建立游標物件
    cursor = conn.cursor()

    # 執行 SQL 查詢
    cursor.execute("SELECT * FROM 韻母對照表")

    # 獲取所有資料
    rows = cursor.fetchall()

    # 初始化字典
    un_bu_dict = {}

    # 從查詢結果中提取資料並將其整理成一個字典
    for row in rows:
        un_bu_dict[row[1]] = {
            'ipa': row[2],
            'poj': row[3],
            'tl': row[4],
            'bp': row[5],
            'tps': row[6],
            'sni': row[7],
            'sni_su_ho': int(row[8]),
            'sni_su_ciok_sing': row[9],
        }

    # 關閉連接
    conn.close()

    return un_bu_dict


if __name__ == "__main__":
    siann_bu_dict = init_siann_bu_dict()    
    siann_code = 'c'

    siann_bu_tl = siann_bu_dict[siann_code]['tl']
    assert siann_bu_tl == 'tsh', "轉換錯誤！"

    siann_bu_ipa = siann_bu_dict[siann_code]['ipa']
    assert siann_bu_ipa == 'ʦʰ', "轉換錯誤！"

    siann_bu_poj = siann_bu_dict[siann_code]['poj']
    assert siann_bu_poj == 'chh', "轉換錯誤！"

    siann_bu_bp = siann_bu_dict[siann_code]['bp']
    assert siann_bu_bp == 'c', "轉換錯誤！"

    siann_bu_tps = siann_bu_dict[siann_code]['tps']
    assert siann_bu_tps == 'ㄘ', "轉換錯誤！"

    siann_bu_sni = siann_bu_dict[siann_code]['sni']
    assert siann_bu_sni == '出', "轉換錯誤！"

    #--------------------------------------------------
    un_bu_dict = init_un_bu_dict()    
    un_code = 'ee'

    un_bu_tl = un_bu_dict[un_code]['tl']
    assert un_bu_tl == 'ee', "轉換錯誤！"

    un_bu_ipa = un_bu_dict[un_code]['ipa']
    assert un_bu_ipa == 'ɛ', "轉換錯誤！"

    un_bu_poj = un_bu_dict[un_code]['poj']
    assert un_bu_poj == 'e', "轉換錯誤！"

    un_bu_bp = un_bu_dict[un_code]['bp']
    assert un_bu_bp == 'e', "轉換錯誤！"

    un_bu_tps = un_bu_dict[un_code]['tps']
    assert un_bu_tps == 'ㄝ', "轉換錯誤！"

    un_bu_sni = un_bu_dict[un_code]['sni']
    assert un_bu_sni == '嘉', "轉換錯誤！"