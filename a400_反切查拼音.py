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
    conn, cursor = connect_to_db('.\\Kong_Un.db')
    
    # 根據反切上字和反切下字來查詢台羅拼音
    siong_ji = ciat_gu[0]
    ha_ji = ciat_gu[1]

    # 顯示結果
    os.system('cls')
    print('\n=================================================')
    print(f'欲查詢拼音之切語為：【{ciat_gu}】')

    # 查詢反切上字
    print('\n-------------------------------------------------')
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

    # 查詢反切下字
    print('\n-------------------------------------------------')
    un_bu = query_ji_piau(ha_ji)
    if not un_bu:
        print(f'查不到【反切下字】：{ha_ji}')
    else:
        un = query_un_bu(ha_ji)
        print(f"切語下字 = {ha_ji} (拼音：{un_bu[0]['拼音']})")
        print(f"韻母：{un[0]['韻母']} [{un[0]['韻母拼音碼']}] IPA: /{un[0]['韻母國際音標']}/")
        print(f"攝：{un[0]['攝']}，調：{un[0]['調']}聲，目次：{un[0]['目次']}")
        print(f"{un[0]['韻']}韻，{un[0]['等']}等（{un[0]['等呼']}），{un[0]['呼']}口呼")

    # 關閉 DB
    close_db_connection(conn)

