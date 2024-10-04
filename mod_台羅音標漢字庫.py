import re
import sqlite3


def connect_to_db(db_path):
    # 創建數據庫連接
    conn = sqlite3.connect(db_path)

    # 創建一個游標
    cursor = conn.cursor()

    return conn, cursor

def close_db_connection(conn):
    # 關閉數據庫連接
    conn.close()

# ==========================================================
# 用 `漢字` 查詢《台語音標》的讀音資訊
# ==========================================================
def han_ji_ca_piau_im(cursor, han_ji):
    """
    根據漢字查詢其台羅音標及相關讀音資訊，並將台羅音標轉換為台語音標。
    若資料紀錄在`常用度`欄位儲存值為空值(NULL)，則將其視為 0，因此可排在查詢結果的最後。
    
    :param cursor: 數據庫游標
    :param han_ji: 欲查詢的漢字
    :return: 包含讀音資訊的字典列表，包含台語音標、聲母、韻母、聲調。
    """

    query = """
    SELECT 
        識別號,
        漢字,
        台羅音標,
        常用度,
        摘要說明
    FROM 
        台羅音標漢字庫
    WHERE 
        漢字 = ?
    ORDER BY 
        COALESCE(常用度, 0) DESC;
    """

    cursor.execute(query, (han_ji,))
    results = cursor.fetchall()

    # 定義【台羅音標】到【台語音標】的轉換規則
    tai_luo_to_tai_gi_mapping = {
        'tsh': 'c',
        'ts': 'z'
    }

    # 將結果轉換為字典列表
    fields = ['識別號', '漢字', '台語音標', '常用度', '摘要說明']
    
    data = []
    for result in results:
        row_dict = dict(zip(fields, result))
        # 取得台羅音標
        tai_loo_im = row_dict['台語音標']

        # 將台羅音標轉換為台語音標
        tai_gi_im = tai_loo_im
        for tai_luo, tai_gi in tai_luo_to_tai_gi_mapping.items():
            tai_gi_im = tai_gi_im.replace(tai_luo, tai_gi)

        # 更新 row_dict 中的台語音標
        row_dict['台語音標'] = tai_gi_im

        # 分析台語音標，拆分出聲母、韻母、聲調
        split_result = split_zu_im(tai_gi_im)
        row_dict['聲母'] = split_result[0]
        row_dict['韻母'] = split_result[1]
        row_dict['聲調'] = split_result[2]

        # 將結果加入列表
        data.append(row_dict)
    
    return data


# ==========================================================
# 自「台羅音標」，分析出：聲母、韻母、調號
# ==========================================================
def split_zu_im(zu_im):
    # 定義聲母的正規表示式，包括常見的聲母，但不包括 m 和 ng
    # siann_bu_pattern = re.compile(r"(b|tsh|ts|g|h|j|kh|k|l|m(?!\d)|ng(?!\d)|n|ph|p|s|th|t|Ø)")
    siann_bu_pattern = re.compile(r"(b|c|z|g|h|j|kh|k|l|m(?!\d)|ng(?!\d)|n|ph|p|s|th|t|Ø)")
    
    # 韻母為 m 或 ng 這種情況的正規表示式 (m\d 或 ng\d)
    un_bu_as_m_or_ng_pattern = re.compile(r"(m|ng)\d")

    result = []

    # 首先檢查是否是 m 或 ng 當作韻母的特殊情況
    if un_bu_as_m_or_ng_pattern.match(zu_im):
        siann_bu = ""  # 沒有聲母
        un_bu = zu_im[:-1]  # 韻母是 m 或 ng
        tiau = zu_im[-1]  # 聲調是最後一個字符
    else:
        # 使用正規表示式來匹配聲母
        siann_bu_match = siann_bu_pattern.match(zu_im)
        if siann_bu_match:
            siann_bu = siann_bu_match.group()  # 找到聲母
            un_bu = zu_im[len(siann_bu):-1]  # 韻母部分
        else:
            siann_bu = ""  # 沒有匹配到聲母，聲母為空字串
            un_bu = zu_im[:-1]  # 韻母是剩下的部分，去掉最後的聲調

        tiau = zu_im[-1]  # 最後一個字符是聲調

    result += [siann_bu]
    result += [un_bu]
    result += [tiau]
    return result


# 使用範例
if __name__ == "__main__":
    # 測試 m, ng 當作韻母的情況
    test_cases = ["m7", "ng7", "tsha1", "thau3", "khong2"]

    for zu_im in test_cases:
        print(f"{zu_im}: {split_zu_im(zu_im)}")
    
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
