import sqlite3

global conn, cursor

def connect_to_kong_un_db():
    # 創建數據庫連接
    conn = sqlite3.connect('.\\Kong_Un.db')

    # 創建一個游標
    cursor = conn.cursor()

    return conn, cursor


def close_kong_un_db(conn):
    # 關閉數據庫連接
    conn.close()

def query_sip_ngoo_im_siann_bu_tui_ziau_piau():
    # SQL 查詢語句
    query = """
    SELECT 識別號,
        廣韻聲母,
        雅俗通聲母,
        聲母拼音碼,
        國際音標
    FROM 十五音聲母對照表;
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
    
def query_sip_ngoo_im_un_bu_tui_ziau_piau():
    # SQL 查詢語句
    query = """
    SELECT 識別號,
        廣韻韻母,
        韻母拼音碼,
        雅俗通韻母,
        舒促聲,
        國際音標
    FROM 十五音韻母對照表;
    """

    # 執行 SQL 查詢
    cursor.execute(query)

    # 獲取查詢結果
    results = cursor.fetchall()

    # 將結果轉換為字典列表
    fields = ['識別號', '廣韻韻母', '韻母拼音碼', '雅俗通韻母', '舒促聲', '國際音標', ]
    dict_results = [dict(zip(fields, result)) for result in results]

    # 回傳字典列表
    return dict_results


def query_table_by_id(table_name, fields, id):
    # 執行 SQL 查詢
    cursor.execute(f"SELECT * FROM {table_name} WHERE 識別號 = ?", (id,))

    # 獲取查詢結果
    results = cursor.fetchall()

    # 將結果轉換為字典列表
    dict_results = [dict(zip(fields, result)) for result in results]

    # 回傳字典列表
    return dict_results
    

def query_table_by_field(table_name, fields, query_field, keyword):
    """
    Query the specified table in the database for records that match the given keyword.

    Args:
        table_name (str): The name of the table to query.
        fields (list): A list of field names to include in the query results.
        query_field (str): The field to use for the query.
        keyword (str): The keyword to search for in the specified field.

    Returns:
        list: A list of dictionaries representing the query results. Each dictionary contains field-value pairs.
    """

    # 執行 SQL 查詢
    cursor.execute(f"SELECT * FROM {table_name} WHERE {query_field} LIKE ?", ('%' + keyword + '%',))

    # 獲取查詢結果
    results = cursor.fetchall()

    # 將結果轉換為字典列表
    dict_results = [dict(zip(fields, result)) for result in results]

    # 回傳字典列表
    return dict_results

def query_ji_piau(han_ji):
    table_name = "字表"
    fields = [
        '識別號', '字', '小韻切語', '拼音', '字義', '上字表識別號', '聲母', '清濁', 
        '聲母拼音碼', '小韻識別號', '小韻字序', '韻母', '韻母拼音碼', '調', '四聲八調', 
        '拼音調號', '備註'
    ]
    query_field = "字"
    result = query_table_by_field(table_name, fields, query_field, han_ji)

    if not result:
        raise Exception("查詢沒有返回結果！")
    
    return result
    

def query_siau_un(ciat_gu):
    table_name = "小韻表"
    fields = [
        '識別號', '小韻字', '拼音', '小韻字集', '字數', '目次編碼', '切語', '小韻字序號', 
        '上字表識別號', '聲母', '聲母拼音碼', '清濁', 
        '韻', '等', '呼', '調', '舒促聲', '韻碼', '韻母', '韻母拼音碼', '四聲八調', '拼音調號',
        '備註', '原有備註', '異體字', '其它備註' 
    ]
    query_field = "切語"
    return query_table_by_field(table_name, fields, query_field, ciat_gu)


def query_ciat_gu_siong_ji(han_ji):
    table_name = "切語上字表"
    fields = [
       '識別號', '發音部位', '聲母', '清濁', '發送收', '聲母拼音碼', '國際音標', '切語上字', '備註',
    ]
    query_field = "切語上字"
    return query_table_by_field(table_name, fields, query_field, han_ji)
    

def query_ciat_gu_ha_ji(han_ji):
    table_name = "切語下字表"
    fields = [
        '識別號', '韻碼', '韻母', '韻母拼音碼', '國際音標', '韻目', '舒促聲', 
        '攝', '調', '韻', '等', '呼', '切語下字',
    ]
    query_field = "切語下字"
    return query_table_by_field(table_name, fields, query_field, han_ji)


def query_un_bu(han_ji):
    query = """
    SELECT 字表.識別號, 字, 拼音, 字表.韻母, 字表.韻母拼音碼, 十五音韻母對照表.國際音標, 
           韻碼表.攝, 韻碼表.目次編碼, 韻碼表.調, 韻碼表.韻, 韻碼表.等, 韻碼表.等呼, 韻碼表.呼
    FROM 字表
    LEFT JOIN 十五音韻母對照表 ON 字表.韻母 = 十五音韻母對照表.廣韻韻母
    LEFT JOIN 韻碼表 ON 字表.韻母 = 韻碼表.韻母
    WHERE 字 = ?;
    """

    # 執行 SQL 查詢
    cursor.execute(query, (han_ji,))

    # 獲取查詢結果
    results = cursor.fetchall()

    # 將結果轉換為字典列表
    fields = [ '識別號', '字', '拼音', '韻母', '韻母拼音碼', '韻母國際音標', '攝', '目次', '調', '韻', '等', '等呼', '呼' ]

    dict_results = [dict(zip(fields, result)) for result in results]

    # 回傳字典列表
    return dict_results


if __name__ == "__main__":
    # 連上 DB
    conn, cursor = connect_to_kong_un_db()
    
    # 測試 "在字表查漢字"
    keyword = "東"
    results = query_ji_piau(keyword)
    # print(results)
    print('\n-------------------------------------------------')
    print(f'字= {results[0]["字"]}')
    print(f'拼音= {results[0]["拼音"]}')
    print(f'小韻切語= {results[0]["小韻切語"]}')

    # 測試 "使用在小韻表查詢"
    keyword = "德紅"
    results = query_siau_un(keyword)
    # print(results)
    print('\n-------------------------------------------------')
    print(f'切語= {results[0]["切語"]}')
    print(f'小韻字= {results[0]["小韻字"]}')
    print(f'拼音= {results[0]["拼音"]}')
    print('\n')
    print(f'聲母= {results[0]["聲母"]}')
    print(f'清濁= {results[0]["清濁"]}')
    print(f'韻母= {results[0]["韻母"]}')
    print(f'四聲八調= {results[0]["四聲八調"]}')
    print(f'聲母拼音碼= {results[0]["聲母拼音碼"]}')
    print(f'韻母拼音碼= {results[0]["韻母拼音碼"]}')
    print(f'拼音調號= {results[0]["拼音調號"]}')
    print('\n')
    print(f'調= {results[0]["調"]}')
    print(f'韻= {results[0]["韻"]}')
    print(f'等= {results[0]["等"]}')
    print(f'呼= {results[0]["呼"]}')

    # 測試 "查詢切語上字" 
    print('\n-------------------------------------------------')
    siong_ji = "德"
    siann_bu = query_ji_piau(siong_ji)
    if not siann_bu:
        print(f'查不到【反切上字】：{siong_ji}')
    else:
        siong_ji_piau = query_table_by_id(
            '切語上字表', 
            ['識別號', '發音部位', '聲母', '清濁', '發送收', '聲母拼音碼', '國際音標', '切語上字', '備註'], 
            siann_bu[0]['上字表識別號']
        )
        print(f"切語上字 = {siong_ji} (拼音：{siann_bu[0]['拼音']})")
        print(f"聲母：{siong_ji_piau[0]['聲母']} [{siong_ji_piau[0]['聲母拼音碼']}] IPA: /{siong_ji_piau[0]['國際音標']}/")
        print(f"(發音部位：{siong_ji_piau[0]['發音部位']}  ，清濁：{siong_ji_piau[0]['清濁']})")

    # 測試 "查詢切語下字" 
    print('\n-------------------------------------------------')
    ha_ji = "紅"
    un_bu = query_ji_piau(ha_ji)
    if not un_bu:
        print(f'查不到【反切下字】：{ha_ji}')
    else:
        un = query_un_bu(ha_ji)
        print(f"切語下字 = {ha_ji} (拼音：{un_bu[0]['拼音']})")
        print(f"韻母：{un[0]['韻母']} [{un[0]['韻母拼音碼']}] IPA: /{un[0]['韻母國際音標']}/")
        print(f"攝：{un[0]['攝']}，調：{un[0]['調']}聲，目次：{un[0]['目次']}")
        print(f"{un[0]['韻']}韻，{un[0]['等']}等（{un[0]['等呼']}），{un[0]['呼']}口呼")

    # 測試 "查詢十五音聲母對照表"
    # print('\n-------------------------------------------------')
    # print('測試 "查詢十五音聲母對照表"\n') 
    # results = query_sip_ngoo_im_siann_bu_tui_ziau_piau()
    # print(results)

    # 測試 "查詢十五音韻母對照表"
    # print('\n-------------------------------------------------')
    # print('測試 "查詢十五音韻母對照表"\n') 
    # results = query_sip_ngoo_im_un_bu_tui_ziau_piau()
    # print(results)

    # 關閉 DB
    close_kong_un_db(conn)