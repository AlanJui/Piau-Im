import sqlite3

def query_data_from_kong_un_db(table_name, fields, query_field, keyword):
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
    # 創建數據庫連接
    conn = sqlite3.connect('.\\Kong_Un.db')

    # 創建一個游標
    cursor = conn.cursor()

    # 執行 SQL 查詢
    cursor.execute(f"SELECT * FROM {table_name} WHERE {query_field} LIKE ?", ('%' + keyword + '%',))

    # 獲取查詢結果
    results = cursor.fetchall()

    # 將結果轉換為字典列表
    dict_results = [dict(zip(fields, result)) for result in results]

    # 關閉數據庫連接
    conn.close()

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
    result = query_data_from_kong_un_db(table_name, fields, query_field, han_ji)

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
    return query_data_from_kong_un_db(table_name, fields, query_field, ciat_gu)


def query_ciat_gu_siong_ji(han_ji):
    table_name = "切語上字表"
    fields = [
       '識別號', '發音部位', '聲母', '清濁', '發送收', '聲母拼音碼', '國際音標', '切語上字', '備註',
    ]
    query_field = "切語上字"
    return query_data_from_kong_un_db(table_name, fields, query_field, han_ji)
    

def query_ciat_gu_ha_ji(han_ji):
    table_name = "切語下字表"
    fields = [
        '識別號', '韻碼', '韻母', '韻母拼音碼', '國際音標', '韻目', '舒促聲', 
        '攝', '調', '韻', '等', '呼', '切語下字',
    ]
    query_field = "切語下字"
    return query_data_from_kong_un_db(table_name, fields, query_field, han_ji)


if __name__ == "__main__":
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
    keyword = "徒"
    results = query_ciat_gu_siong_ji(keyword)
    print(results)

    # 測試 "查詢切語下字" 
    keyword = "紅"
    results = query_ciat_gu_ha_ji(keyword)
    print(results)