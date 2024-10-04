import sqlite3


def connect_to_db(db_path):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    return conn, cursor

def close_db_connection(conn):
    conn.close()

def TL_Tng_Zu_Im(siann_bu, un_bu, siann_tiau, cursor):
    """
    根據傳入的台羅拼音聲母、韻母、聲調，轉換成對應的方音符號
    :param siann_bu: 聲母 (台羅拼音)
    :param un_bu: 韻母 (台羅拼音)
    :param siann_tiau: 聲調 (台羅拼音中的數字)
    :param cursor: 數據庫游標
    :return: 包含方音符號的字典
    """

    # 查詢聲母表，將台羅拼音的聲母轉換成方音符號
    cursor.execute("SELECT 方音符號 FROM 聲母表 WHERE 台羅拚音 = ?", (siann_bu,))
    siann_bu_result = cursor.fetchone()
    if siann_bu_result:
        zu_im_siann_bu = siann_bu_result[0]  # 取得方音符號
    else:
        zu_im_siann_bu = ''  # 無聲母的情況

    # 查詢韻母表，將台羅拼音的韻母轉換成方音符號
    cursor.execute("SELECT 方音符號 FROM 韻母表 WHERE 台羅拚音 = ?", (un_bu,))
    un_bu_result = cursor.fetchone()
    if un_bu_result:
        zu_im_un_bu = un_bu_result[0]  # 取得方音符號
    else:
        zu_im_un_bu = ''

    # 查詢聲調表，將台羅拼音的聲調轉換成方音符號
    cursor.execute("SELECT 方音符號 FROM 聲調表 WHERE 台羅八聲調 = ?", (siann_tiau,))
    siann_tiau_result = cursor.fetchone()
    if siann_tiau_result:
        zu_im_siann_tiau = siann_tiau_result[0]  # 取得方音符號
    else:
        zu_im_siann_tiau = ''

    return {
        '注音符號': f"{zu_im_siann_bu}{zu_im_un_bu}{zu_im_siann_tiau}",
        '聲母': zu_im_siann_bu,
        '韻母': zu_im_un_bu,
        '聲調': zu_im_siann_tiau
    }

def dict_to_str(zu_im_hu_ho):
    return f"{zu_im_hu_ho['聲母']}{zu_im_hu_ho['韻母']}{zu_im_hu_ho['聲調']}"



# 測試範例
if __name__ == "__main__":
    # 連接到資料庫
    db_path = "Tai_Loo_Han_Ji_Khoo.db"  # 請替換成正確的資料庫路徑
    conn, cursor = connect_to_db(db_path)

    # 測試 1
    zu_im_hu_ho = TL_Tng_Zu_Im(siann_bu='p', un_bu='ut', siann_tiau=4, cursor=cursor)
    assert zu_im_hu_ho['聲母'] == 'ㄅ', "聲母不正確"
    assert zu_im_hu_ho['韻母'] == 'ㄨㆵ', "韻母不正確"
    assert zu_im_hu_ho['聲調'] == '', "聲調不正確"
    print("測試 1 成功")
    print(str(zu_im_hu_ho))
    zu_im_hu_ho_str = dict_to_str(zu_im_hu_ho)
    print(zu_im_hu_ho_str)

    # 測試 2
    zu_im_hu_ho = TL_Tng_Zu_Im(siann_bu='', un_bu='m', siann_tiau=7, cursor=cursor)
    assert zu_im_hu_ho['聲母'] == '', "聲母不正確"
    assert zu_im_hu_ho['韻母'] == 'ㆬ', "韻母不正確"
    assert zu_im_hu_ho['聲調'] == '˫', "聲調不正確"
    print("測試 2 成功")
    print(str(zu_im_hu_ho))
    zu_im_hu_ho_str = dict_to_str(zu_im_hu_ho)
    print(zu_im_hu_ho_str)

    # 關閉資料庫連接
    close_db_connection(conn)
