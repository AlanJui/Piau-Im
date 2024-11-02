import re

"""
用 `漢字` 查詢《彙集雅俗通十五音》的標音
"""
def han_ji_ca_piau_im(cursor, han_ji):
    """
    根據漢字查詢其讀音資訊。 若資料紀錄在`常用度`欄位儲存值為空值(NULL)
    ，則將其視為 0，因此可排在查詢結果的最後。

    :param cursor: 數據庫游標
    :param han_ji: 欲查詢的漢字
    :return: 包含讀音資訊的字典列表

    SELECT *
    FROM 漢字表
    WHERE 漢字 = ?
    ORDER BY COALESCE(常用度, 0) DESC;
    """

    query = """
    SELECT 識別號, 漢字, 漢字標音, 常用度, 切音, 字韻, 聲調, 舒促聲,
        聲, 韻, 調, 雅俗通標音, 十五音標音
    FROM 漢字表
    WHERE 漢字 = ?
    ORDER BY COALESCE(常用度, 0) DESC;
    """
    cursor.execute(query, (han_ji,))
    results = cursor.fetchall()

    fields = [
        '識別號', '漢字', '漢字標音', '常用度', '切音', '字韻', '聲調', '舒促聲',
        '聲', '韻', '調', '雅俗通標音', '十五音標音'
    ]
    return [dict(zip(fields, result)) for result in results]


"""
反切查詢：根據《彙集雅俗通十五音》的切語查詢漢字、讀音
"""
def huan_ciat_ca_piau_im(cursor, 字韻, 聲調, 切音):
    """
    根據切語查詢漢字、讀音。

    :param cursor: 數據庫游標
    :param 字韻: 韻母
    :param 切音: 聲母
    :param 聲調: 聲調
    :return: 包含讀音資訊的字典列表

    SELECT *
    FROM 漢字表
    WHERE 字韻 = ? AND 切音 = ? AND 聲調 = ?
    ORDER BY COALESCE(常用度, 0) DESC;
    """

    query = """
    SELECT 識別號, 漢字, 漢字標音, 常用度, 切音, 字韻, 聲調, 舒促聲,
        聲, 韻, 調, 雅俗通標音, 十五音標音
    FROM 漢字表
    WHERE 字韻 = ? AND 切音 = ? AND 聲調 = ?
    ORDER BY COALESCE(常用度, 0) DESC;
    """
    cursor.execute(query, (字韻, 切音, 聲調))
    results = cursor.fetchall()

    fields = [
        '識別號', '漢字', '漢字標音', '常用度', '切音', '字韻', '聲調', '舒促聲',
        '聲', '韻', '調', '雅俗通標音', '十五音標音'
    ]
    return [dict(zip(fields, result)) for result in results]


def tiau_ho_tng_siann_tiau(調號):
    """
    將調號轉換成對應的聲調名稱。
    """

    聲調 = {
        '一' : '上平',
        '二' : '上上',
        '三' : '上去',
        '四' : '上入',
        '五' : '下平',
        '六' : '下上',
        '七' : '下去',
        '八' : '下入',
    }
    return 聲調.get(調號, None)