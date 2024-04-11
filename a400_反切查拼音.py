import os
import sqlite3
import sys

global conn, cursor


def connect_to_db(db_name):
    # 創建數據庫連接
    conn = sqlite3.connect(db_name)

    # 創建一個游標
    cursor = conn.cursor()

    return conn, cursor


def close_db_connection(conn):
    # 關閉數據庫連接
    conn.close()


"""
查音雅俗通十五音聲母對照表
"""
def query_sip_ngoo_im_siann_bu_tui_ciau_piau(cursor):
    # SQL 查詢語句
    query = """
    SELECT 識別號,
        廣韻聲母,
        雅俗通聲母,
        聲母拼音碼,
        國際音標
    FROM 聲母對照表;
    """

    # 執行 SQL 查詢
    cursor.execute(query)

    # 獲取查詢結果
    results = cursor.fetchall()

    # 將結果轉換為字典列表
    fields = ['識別號', '廣韻聲母', '雅俗通聲母', '聲母拼音碼', '國際音標', ]
    dict_results = [dict(zip(fields, result)) for result in results]

    # 回傳字典列表
    return dict_results

    
"""
查詢雅俗通十五音韻母對照表
"""
def query_sip_ngoo_im_un_bu_tui_ciau_piau(cursor):
    # SQL 查詢語句
    query = """
    SELECT 識別號,
        廣韻韻母,
        雅俗通韻母,
        舒促聲,
        拚音碼,
        國際音標,
        十五音序號,
        林進三拚音碼
    FROM 韻母對照表;
    """

    # 執行 SQL 查詢
    cursor.execute(query)

    # 獲取查詢結果
    results = cursor.fetchall()

    # 將結果轉換為字典列表
    fields = ['識別號', '廣韻韻母', '雅俗通韻母', '舒促聲', '韻母拼音碼', '國際音標', ]
    dict_results = [dict(zip(fields, result)) for result in results]

    # 回傳字典列表
    return dict_results


"""
查詢漢字在 `小韻` 中的標音
"""
def query_han_ji_piau_im(cursor, han_ji):
    """
    根據漢字查詢其讀音資訊。
    
    :param cursor: 數據庫游標
    :param han_zi: 欲查詢的漢字
    :return: 包含讀音資訊的字典列表
    """
    query = """
    SELECT 小韻字, 切語, 拼音
    FROM 小韻查詢
    WHERE 小韻字 = ?;
    """
    cursor.execute(query, (han_ji,))
    results = cursor.fetchall()
    
    # 將結果轉換為字典列表
    fields = ['小韻字', '切語', '標音']
    return [dict(zip(fields, result)) for result in results]

"""
查詢某漢字的 `小韻` 資料
"""
def han_ji_cha_siau_un(cursor, han_ji):
    # SQL 查詢語句
    query = """
    SELECT *
    FROM 小韻查詢
    WHERE 小韻字 = ?;
    """

    # 執行 SQL 查詢
    cursor.execute(query, (han_ji,))

    # 獲取查詢結果
    results = cursor.fetchall()

    # 將結果轉換為字典列表
    fields = ['小韻字', '切語', '標音', '目次編碼', '小韻字序號', '小韻字集', '字數', 
        '發音部位', '聲母', '清濁', '發送收', '聲母拼音碼', '切語上字集',
        '韻系列號', '韻系行號', '韻目索引', '目次', '攝', '韻系', 
        '韻目', '調', '呼', '等', '韻母', '切語下字集', '等呼', '韻母拼音碼']

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
    FROM 漢字查廣韻標音
    WHERE 字 = ?;
    """
    cursor.execute(query, (han_ji,))
    results = cursor.fetchall()
    
    # 將結果轉換為字典列表
    fields = [
        '識別號', '字', '同音字序', '切語', '漢字標音', '字義',
        '小韻字', '目次編碼', '小韻字序號', '小韻字集', '字數',
        '發音部位', '聲母', '清濁', '發送收', '上字標音', '切語上字集',
        '韻系列號', '韻系行號', '韻目索引', '目次', '攝', '韻系', 
        '韻目', '調', '呼', '等', '韻母', '切語下字集', '等呼', '下字標音', 
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


if __name__ == "__main__":
    # 確認使用者有輸入反切之切語參數
    if len(sys.argv) != 2:
        print("請輸入欲查詢讀音之【切語】(反切上字及下字)!")
        os._exit(-1)

    ciat_gu = sys.argv[1]

    # 檢查反切拼音是否有兩個字
    if len(ciat_gu) != 2:
        print("反切用的切語，必須有兩個漢字！")
        os._exit(-1)

    # 連上 DB
    conn, cursor = connect_to_db('.\\Kong_Un_V2.db')

    # 根據反切上字和反切下字來查詢台羅拼音
    siong_ji = ciat_gu[0]
    ha_ji = ciat_gu[1]

    # 顯示結果
    os.system('cls')
    print('\n=================================================')
    print(f'欲查詢拼音之切語為：【{ciat_gu}】')

    # 查詢反切上字
    print('\n-------------------------------------------------')
    print('【切語上字】：\n')
    siong_ji_im = han_ji_cha_piau_im(cursor, siong_ji)
    siong_ji_piau = cha_ciat_gu_siong_ji(cursor, siong_ji)
    if not siong_ji_piau:
        print(f'查不到【反切上字】：{siong_ji}')
    else:
        print(f"切語上字 = {siong_ji} (標音：{siong_ji_im[0]['漢字標音']})")
        print(f"聲母：{siong_ji_piau[0]['聲母']} [{siong_ji_piau[0]['聲母拼音碼']}] ")
        print(f"(發音部位：{siong_ji_piau[0]['發音部位']}，清濁：{siong_ji_piau[0]['清濁']}，發送收：{siong_ji_piau[0]['發送收']})")

    # 查詢反切下字
    print('\n-------------------------------------------------')
    print('【切語下字】：\n')
    ha_ji_im = han_ji_cha_piau_im(cursor, ha_ji)
    ha_ji_piau = cha_ciat_gu_ha_ji(cursor, ha_ji)
    if not ha_ji_piau:
        print(f'查不到【反切下字】：{ha_ji}')
    else:
        print(f"切語下字 = {ha_ji} (標音：{ha_ji_im[0]['漢字標音']})")
        print(f"韻母：{ha_ji_piau[0]['韻母']} [{ha_ji_piau[0]['韻母拼音碼']}]")
        print(f"攝：{ha_ji_piau[0]['攝']}，調：{ha_ji_piau[0]['調']}聲，目次：{ha_ji_piau[0]['目次']}")
        print(f"{ha_ji_piau[0]['韻系']}韻系，{ha_ji_piau[0]['韻目']}韻，{ha_ji_piau[0]['呼']}口呼，{ha_ji_piau[0]['等']}等 ({ha_ji_piau[0]['等呼']})")

    # 組合拼音
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
    siann = siong_ji_piau[0]['聲母拼音碼']
    cing_tok_str = siong_ji_piau[0]['清濁']
    cing_tok = cing_tok_str[-1]
    un = ha_ji_piau[0]['韻母拼音碼']
    tiau_ho = tiau_ho_list[ f"{cing_tok}{ha_ji_piau[0]['調']}" ]

    print('\n-------------------------------------------------')
    print(f'【切語拼音】：{ciat_gu} [{siann}{un}{tiau_ho}]\n')

    # 關閉 DB
    close_db_connection(conn)