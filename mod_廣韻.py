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


"""
用 `漢字` 查詢《廣韻》的標音
"""
def han_ji_ca_piau_im(cursor, han_ji):
    """
    根據漢字查詢其讀音資訊。 若資料紀錄在`台羅聲調`欄位儲存值為空值(NULL)
    ，則將其視為 0，因此可排在查詢結果的最後。
    
    :param cursor: 數據庫游標
    :param han_ji: 欲查詢的漢字
    :return: 包含讀音資訊的字典列表

    SELECT *
    FROM 廣韻漢字庫
    WHERE 漢字 = ?
    ORDER BY COALESCE(台羅聲調, 0) DESC;
    """

    query = """
    SELECT *
    FROM 廣韻漢字庫
    WHERE 漢字 = ?
    ORDER BY COALESCE(台羅聲調, 0) DESC;
    """
    cursor.execute(query, (han_ji,))
    results = cursor.fetchall()
    
    # 將結果轉換為字典列表
    fields = [
        '字號', '漢字', '標音', '上字', '下字', '上字號', '聲母', '聲母標音', '七聲類',
        '清濁', '發送收', '下字號', '韻母', '韻母標音', '韻目列號', '攝', '調', '目次',
        '韻目', '等呼', '等', '呼', '廣韻調名', '台羅聲調', '字義識別號'
    ]   
    return [dict(zip(fields, result)) for result in results]


def ca_siann_bu_piau_im(cursor, siann_bu):
    """
    根據聲母標音查詢其國際音標聲母和方音聲母。
    
    :param cursor: 數據庫游標
    :param siann_bu: 欲查詢的聲母標音（台語音標）
    :return: 包含國際音標聲母和方音聲母的字典列表

    SELECT 國際音標聲母, 方音聲母
    FROM 聲母對照表
    WHERE 台語音標聲母 = ?;
    """

    query = """
    SELECT 國際音標聲母, 方音聲母
    FROM 聲母對照表
    WHERE 台語音標聲母 = ?;
    """
    cursor.execute(query, (siann_bu,))
    results = cursor.fetchall()
    
    fields = ['國際音標聲母', '方音聲母']
    return [dict(zip(fields, result)) for result in results]


def ca_un_bu_piau_im(cursor, un_bu):
    """
    根據韻母標音查詢其國際音標韻母和方音韻母。
    
    :param cursor: 數據庫游標
    :param un_bu: 欲查詢的韻母標音（台語音標）
    :return: 包含國際音標韻母和方音韻母的字典列表

    SELECT 國際音標韻母, 方音韻母
    FROM 韻母對照表
    WHERE 台語音標韻母 = ?;
    """

    query = """
    SELECT 國際音標韻母, 方音韻母
    FROM 韻母對照表
    WHERE 台語音標韻母 = ?;
    """
    cursor.execute(query, (un_bu,))
    results = cursor.fetchall()
    
    fields = ['國際音標韻母', '方音韻母']
    return [dict(zip(fields, result)) for result in results]


def huan_ciat_ca_piau_im(cursor, siong_ji, ha_ji):
    """
    根據反切上字和下字查詢符合條件的所有漢字及其讀音資訊。
    
    :param cursor: 數據庫游標
    :param siong_ji: 反切上字
    :param ha_ji: 反切下字
    :return: 包含讀音資訊的字典列表

    SELECT *
    FROM 廣韻漢字庫
    WHERE 上字 = ? AND 下字 = ?;
    """

    query = """
    SELECT *
    FROM 廣韻漢字庫
    WHERE 上字 = ? AND 下字 = ?;
    """
    cursor.execute(query, (siong_ji, ha_ji))
    results = cursor.fetchall()
    
    # 將結果轉換為字典列表
    fields = [
        '字號', '漢字', '標音', '上字', '下字', '上字號', '聲母', '聲母標音', '七聲類',
        '清濁', '發送收', '下字號', '韻母', '韻母標音', '韻目列號', '攝', '調', '目次',
        '韻目', '等呼', '等', '呼', '廣韻調名', '台羅聲調', '字義識別號'
    ]
    return [dict(zip(fields, result)) for result in results]
