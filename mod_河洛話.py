# 新檔名：mod_河洛話.py
# 舊檔名：mod_台羅音標漢字庫.py
import sqlite3

from mod_標音 import split_tai_gi_im_piau


# ==========================================================
# 用 `漢字` 查詢《台語音標》的讀音資訊
# 在【台羅音標漢字庫】資料表結構中，以【常用度】欄位之值，區分【文讀音】與【白話音】。
# 通用音：常用度 < 1.00；表文、白通用的讀音，最常用的讀音其值為 1.00，次常用的讀音值為 0.90，其餘則使用值為 0.89 ~ 0.81。
# 文讀音：常用度 > 0.60；最常用的讀音其值為 0.80，次常用的讀音其值為 0.70；其餘則使用數值 0.69 ~ 0.61。
# 白話音：常用度 > 0.40；最常用的讀音其值為 0.60，次常用的讀音其值為 0.50；其餘則使用數值 0.59 ~ 0.41。
# 其　它：常用度 > 0.00；使用數值 0.40 ~ 0.01；使用時機為：（1）方言地方腔；(2) 罕見發音；(3) 尚未查證屬文讀音或白話音 。
# ==========================================================
def han_ji_ca_piau_im(cursor, han_ji, ue_im_lui_piat="文讀音"):
    """
    根據漢字查詢其台羅音標及相關讀音資訊，並將台羅音標轉換為台語音標。
    若資料紀錄在常用度欄位儲存值為空值(NULL)，則將其視為 0，因此可排在查詢結果的最後。

    :param cursor: 數據庫游標
    :param han_ji: 欲查詢的漢字
    :param hue_im: 查詢的讀音類型，可以是 "文讀音"、"白話音" 或 "其它"
    :return: 包含讀音資訊的字典列表，包含台語音標、聲母、韻母、聲調。
    """

    # 將文白通用音視為第一優選
    common_reading_condition = "常用度 >= 0.81 AND 常用度 <= 1.0"

    # 根據不同讀音類型，添加額外的查詢條件
    if ue_im_lui_piat == "文讀音":
        reading_condition = f"({common_reading_condition}) OR (常用度 >= 0.61 AND 常用度 < 0.81)"
    elif ue_im_lui_piat == "白話音":
        reading_condition = f"({common_reading_condition}) OR (常用度 > 0.40 AND 常用度 < 0.61)"
    elif ue_im_lui_piat == "其它":
        reading_condition = "常用度 > 0.00 AND 常用度 <= 0.40"
    else:
        reading_condition = "1=1"  # 查詢所有

    query = f"""
    SELECT
        識別號,
        漢字,
        台羅音標,
        常用度,
        摘要說明
    FROM
        漢字庫
    WHERE
        漢字 = ? AND ({reading_condition})
    ORDER BY
        COALESCE(常用度, 0) DESC;
    """

    cursor.execute(query, (han_ji,))
    results = cursor.fetchall()

    # 如果沒有找到符合條件的讀音，則查詢所有讀音，並選擇常用度最高者
    if not results:
        query = """
        SELECT
            識別號,
            漢字,
            台羅音標,
            常用度,
            摘要說明
        FROM
            漢字庫
        WHERE
            漢字 = ?
        ORDER BY
            COALESCE(常用度, 0) DESC
        LIMIT 1;
        """
        cursor.execute(query, (han_ji,))
        results = cursor.fetchall()

    # 將結果轉換為字典列表
    fields = ['識別號', '漢字', '台語音標', '常用度', '摘要說明']

    data = []
    for result in results:
        row_dict = dict(zip(fields, result))
        # 取得台羅音標
        tai_loo_im = row_dict['台語音標']

        # 將台羅音標轉換為台語音標
        split_result = split_tai_gi_im_piau(tai_loo_im)
        row_dict['聲母'] = split_result[0]
        row_dict['韻母'] = split_result[1]
        row_dict['聲調'] = split_result[2]

        # 更新 row_dict 中的台語音標
        row_dict['台語音標'] = f'{row_dict["聲母"]}{row_dict["韻母"]}{row_dict["聲調"]}'

        # 將結果加入列表
        data.append(row_dict)

    return data



# 使用範例
if __name__ == "__main__":
    def connect_to_db(db_path):
        # 創建數據庫連接
        conn = sqlite3.connect(db_path)

        # 創建一個游標
        cursor = conn.cursor()

        return conn, cursor

    def close_db_connection(conn):
        # 關閉數據庫連接
        conn.close()

    # 測試 m, ng 當作韻母的情況
    test_cases = ["m7", "ng7", "tsha1", "thau3", "khong2"]

    for im_piau in test_cases:
        print(f"{im_piau}: {split_tai_gi_im_piau(im_piau)}")

    # 連接到資料庫
    db_path = "Tai_Loo_Han_Ji_Khoo.db"   # 替換成你的資料庫路徑
    conn, cursor = connect_to_db(db_path)

    # 驗證資料庫所需使用的表格已存在
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
    tables = cursor.fetchall()
    print(tables)

    # 測試查詢漢字 "不"
    result = han_ji_ca_piau_im(cursor, '不')
    for row in result:
        print(row)

    # 關閉資料庫連接
    close_db_connection(conn)

