import sqlite3


def connect_to_db_by_context_manager_decorator(db_path):
    def connect_to_db(func):
        def wrapper(*args, **kwargs):
            # 創建數據庫連接
            conn = sqlite3.connect(db_path)

            # 創建一個游標
            cursor = conn.cursor()

            # 執行函數
            result = func(cursor, *args, **kwargs)

            # 關閉數據庫連接
            conn.close()

            return result

        return wrapper

    return connect_to_db


def connect_to_db(db_path):
    # 創建數據庫連接
    conn = sqlite3.connect(db_path)

    # 創建一個游標
    cursor = conn.cursor()

    return conn, cursor


def connect_to_db2(db_path):
    # 創建數據庫連接
    conn = sqlite3.connect(db_path)

    # 創建一個游標
    cursor = conn.cursor()  # noqa: F841

    return conn


def close_db_connection(conn):
    # 關閉數據庫連接
    conn.close()


#===============================================================================
# 查音雅俗通十五音聲母對照表
#===============================================================================
def cha_siann_bu_tui_ciau_piau(cursor):
    # SQL 查詢語句
    query = """
    SELECT 識別號,
        聲母碼,
        聲母國際音標,
        白話字聲母,
        閩拼聲母,
        台羅聲母,
        方音聲母,
        十五音聲母
    FROM 聲母對照表;
    """

    # 執行 SQL 查詢
    cursor.execute(query)

    # 獲取查詢結果
    results = cursor.fetchall()

    # 將結果轉換為字典列表
    fields = [
        '識別號', '聲母碼', '聲母國際音標', '白話字聲母', '閩拼聲母', '台羅聲母', 
        '方音聲母', '十五音聲母', 
    ]
    dict_results = [dict(zip(fields, result)) for result in results]

    # 回傳字典列表
    return dict_results

    
"""
查詢雅俗通十五音韻母對照表
"""
def cha_un_bu_tui_ciau_piau(cursor):
    # SQL 查詢語句
    query = """
    SELECT 識別號,
        韻母碼,
        韻母國際音標,
        白話字韻母,
        閩拼韻母,
        台羅韻母,
        方音韻母,
        十五音韻母,
        舒促聲,
        十五音序
    FROM 韻母對照表;
    """

    # 執行 SQL 查詢
    cursor.execute(query)

    # 獲取查詢結果
    results = cursor.fetchall()

    # 將結果轉換為字典列表
    fields = [
        '識別號', '韻母碼', '韻母國際音標', '白話字韻母', '閩拼韻母', '台羅韻母', 
        '方音韻母', '十五音韻母', '舒促聲', '十五音序'
    ]
    dict_results = [dict(zip(fields, result)) for result in results]

    # 回傳字典列表
    return dict_results


"""
使用 `小韻檢視` 查詢某小韻之切語及標音
"""
def cha_siau_un_piau_im(cursor, han_ji):
    """
    根據漢字查詢其讀音資訊。
    
    :param cursor: 數據庫游標
    :param han_zi: 欲查詢的漢字
    :return: 包含讀音資訊的字典列表
    """
    query = """
    SELECT 小韻字, 小韻切語, 小韻標音
    FROM 小韻檢視
    WHERE 小韻字 = ?;
    """
    cursor.execute(query, (han_ji,))
    results = cursor.fetchall()
    
    # 將結果轉換為字典列表
    fields = ['小韻字', '小韻切語', '小韻標音']
    return [dict(zip(fields, result)) for result in results]

"""
查詢某漢字的 `小韻` 資料
"""
def han_ji_cha_siau_un(cursor, han_ji):
    # SQL 查詢語句
    query = """
    SELECT *
    FROM 小韻檢視
    WHERE 小韻字 = ?;
    """

    # 執行 SQL 查詢
    cursor.execute(query, (han_ji,))

    # 獲取查詢結果
    results = cursor.fetchall()

    # 將結果轉換為字典列表
    fields = [
        '小韻識別號', '上字表識別號', '下字表識別號', 
        '小韻字', '小韻切語', '小韻標音', '小韻目次', '小韻字序號', '小韻字集', '字數', 
        '廣韻聲母', '七聲類', '發音部位', '清濁', '發送收', 
        '聲母碼', '聲母國際音標', '白話字聲母', '閩拼聲母', '台羅聲母', '方音聲母','十五音聲母', 
        '廣韻韻母', '目次', '攝', '韻系', '韻目', '調', '呼', '等', '韻母', '等呼', '下字標音', 
        '韻母碼', '韻母國際音標', '白話字韻母', '閩拼韻母', '台羅韻母', '方音韻母', '十五音韻母',
    ]

    dict_results = [dict(zip(fields, result)) for result in results]

    # 回傳字典列表
    return dict_results

"""
用 `漢字` 查詢《廣韻》的標音
"""
def han_ji_cha_piau_im(cursor, han_ji):
    """
    根據漢字查詢其讀音資訊。
    
    :param cursor: 數據庫游標
    :param han_ji: 欲查詢的漢字
    :return: 包含讀音資訊的字典列表
    """
    query = """
    SELECT *
    FROM 漢字檢視
    WHERE 字 = ?;
    """
    cursor.execute(query, (han_ji,))
    results = cursor.fetchall()
    
    # 將結果轉換為字典列表
    fields = [
        '漢字識別號', '字', '同音字序', '切語', '漢字標音', '字義',
        '小韻識別號', '上字識別號', '下字識別號', '小韻字', '小韻切語', '小韻標音', '小韻目次', '小韻字序號', 
        '廣韻聲母', '七聲類', '發音部位', '清濁', '發送收', 
        '聲母碼', '聲母國際音標', '白話字聲母', '閩拼聲母', '台羅聲母', '方音聲母', '十五音聲母',
        '廣韻韻母', '目次', '攝', '韻系', '韻目', '調', '呼', '等', '等呼', 
        '韻母碼', '韻母國際音標', '白話字韻母', '閩拼韻母', '台羅韻母', '方音韻母', '十五音韻母',
    ]   
    return [dict(zip(fields, result)) for result in results]


def query_table_by_field(cursor, table_name, fields, query_field, keyword):
    # 執行 SQL 查詢
    cursor.execute(f"SELECT * FROM {table_name} WHERE {query_field} LIKE ?", ('%' + keyword + '%',))

    # 獲取查詢結果
    results = cursor.fetchall()

    # 將結果轉換為字典列表
    dict_results = [dict(zip(fields, result)) for result in results]

    # 回傳字典列表
    return dict_results
    

def cha_ciat_gu_siong_ji(cursor, siong_ji):
    table_name = "切語上字表"
    fields = [
       '識別號', '廣韻聲母識別號', '七聲類', '發音部位', '聲母', '清濁', '發送收', 
       '聲母拼音碼', '切語上字集', '備註',
    ]
    query_field = "切語上字集"
    return query_table_by_field(cursor, table_name, fields, query_field, siong_ji)
    

def cha_ciat_gu_ha_ji(cursor, ha_ji):
    table_name = "切語下字表"
    fields = [
        '識別號', '廣韻韻母識別號', '韻系列號', '韻系行號',
        '韻目索引', '目次識別號', '目次',
        '攝', '韻系', '韻目', '調', '呼', '等', '韻母',
        '切語下字集', '等呼', '韻母拼音碼', '備註',
    ]
    query_field = "切語下字集"
    return query_table_by_field(cursor, table_name, fields, query_field, ha_ji)


def query_table_by_id(cursor, table_name, fields, id):
    # 執行 SQL 查詢
    cursor.execute(f"SELECT * FROM {table_name} WHERE 識別號 = ?", (id,))

    # 獲取查詢結果
    results = cursor.fetchall()

    # 將結果轉換為字典列表
    dict_results = [dict(zip(fields, result)) for result in results]

    # 回傳字典列表
    return dict_results


# =========================================================
# 判斷調號
# =========================================================
def piau_tiau_ho(ji_tian_piau_im):
    tiau_ho_list = {
        '清平': 1,
        '清上': 2,
        '清去': 3,
        '清入': 4,
        '濁平': 5,
        '濁上': 2,
        '濁去': 7,
        '濁入': 8,
    }
    siong_ji_cing_tok = ji_tian_piau_im['清濁']
    cing_tok = siong_ji_cing_tok[-1]
    sing_tiau = ji_tian_piau_im['調']
    su_sing_pat_tiau = tiau_ho_list[ f"{cing_tok}{sing_tiau}" ]
    return su_sing_pat_tiau


def init_sing_bu_dict(cursor):
    # 執行 SQL 查詢
    cursor.execute("SELECT * FROM 聲母對照表")

    # 獲取所有資料
    rows = cursor.fetchall()

    # 初始化字典
    sing_bu_dict = {}

    # 從查詢結果中提取資料並將其整理成一個字典
    for row in rows:
        sing_bu_dict[row[1]] = {
            'code': row[1],
            'ipa': row[2],
            'poj': row[3],
            'bp': row[4],
            'tl': row[5],
            'tps': row[6],
            'sni': row[7],
        }

    return sing_bu_dict


def init_un_bu_dict(cursor):
    # 執行 SQL 查詢
    cursor.execute("SELECT * FROM 韻母對照表")

    # 獲取所有資料
    rows = cursor.fetchall()

    # 初始化字典
    un_bu_dict = {}

    # 從查詢結果中提取資料並將其整理成一個字典
    for row in rows:
        un_bu_dict[row[1]] = {
            'code': row[1],
            'ipa': row[2],
            'poj': row[3],
            'bp': row[4],
            'tl': row[5],
            'tps': row[6],
            'sni': row[7],
            'sni_su_ciok_sing': row[8],
            'sni_su_ho': int(row[9]),
        }

    return un_bu_dict


if __name__ == "__main__":
    sing_bu_dict = init_sing_bu_dict()    
    sing_code = 'c'

    sing_bu_tl = sing_bu_dict[sing_code]['tl']
    assert sing_bu_tl == 'tsh', "轉換錯誤！"

    sing_bu_ipa = sing_bu_dict[sing_code]['ipa']
    assert sing_bu_ipa == 'ʦʰ', "轉換錯誤！"

    sing_bu_poj = sing_bu_dict[sing_code]['poj']
    assert sing_bu_poj == 'chh', "轉換錯誤！"

    sing_bu_bp = sing_bu_dict[sing_code]['bp']
    assert sing_bu_bp == 'c', "轉換錯誤！"

    sing_bu_tps = sing_bu_dict[sing_code]['tps']
    assert sing_bu_tps == 'ㄘ', "轉換錯誤！"

    sing_bu_sni = sing_bu_dict[sing_code]['sni']
    assert sing_bu_sni == '出', "轉換錯誤！"

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