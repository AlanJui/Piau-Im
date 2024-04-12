import sqlite3



def connect_to_db(db_name):
    # 創建數據庫連接
    conn = sqlite3.connect(db_name)

    # 創建一個游標
    cursor = conn.cursor()

    return conn, cursor


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
        '小韻識別號', '小韻字', '小韻切語', '小韻標音', '目次', '小韻字序號', '小韻字集', '字數', 
        '聲母發音部位', '清濁', '發送收', '廣韻聲母', '雅俗通聲母', '上字標音', 
        '聲母國際音標', '白話字聲母', '閩拼聲母', '台羅聲母', '方音聲母',
        '攝', '韻目', '調', '呼', '等', '韻母', '等呼', '廣韻韻母', '雅俗通韻母', '下字標音', 
        '韻母國際音標', '白話字韻母', '閩拼韻母', '台羅韻母', '方音韻母',
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
    FROM 漢字廣韻標音檢視
    WHERE 字 = ?;
    """
    cursor.execute(query, (han_ji,))
    results = cursor.fetchall()
    
    # 將結果轉換為字典列表
    fields = [
        '漢字識別號', '字', '同音字序', '切語', '漢字標音', '字義',
        '小韻字', '目次編碼', '小韻字序號', '小韻字集', '字數', 
        '發音部位', '清濁', '發送收', '切語上字集', '廣韻聲母', '雅俗通聲母',
        '上字標音', '聲母國際音標', '白話字聲母', '閩拼聲母', '台羅聲母', '方音聲母',
        '韻系列號', '韻系行號', '韻目索引', '目次', '攝', '韻系', '韻目', '調', '呼', '等', '韻母', '切語下字集', '等呼',
        '廣韻韻母', '雅俗通韻母', '下字標音', '韻母國際音標', '白話字韻母', '閩拼韻母', '台羅韻母', '方音韻母',
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
       '識別號', '聲母識別號', '發音部位', '聲母', '清濁', '發送收', 
       '聲母拼音碼', '切語上字集', '備註',
    ]
    query_field = "切語上字集"
    return query_table_by_field(cursor, table_name, fields, query_field, siong_ji)
    

def cha_ciat_gu_ha_ji(cursor, ha_ji):
    table_name = "切語下字表"
    fields = [
        '識別號', '韻母識別號', '韻系列號', '韻系行號',
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
