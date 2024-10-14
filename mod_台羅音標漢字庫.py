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
# 查詢語音類型，若未設定則預設為文讀音
# ==========================================================
def get_sound_type(wb):
    try:
        reading_type = wb.names['語音類型'].refers_to_range.value
    except KeyError:
        reading_type = "文讀音"
    return reading_type


# ==========================================================
# 用 `漢字` 查詢《台語音標》的讀音資訊
# 在【台羅音標漢字庫】資料表結構中，以【常用度】欄位之值，區分【文讀音】與【白話音】。
# 文讀音：常用度 > 0.60；最常用的讀音其值為 0.80，次常用的讀音其值為 0.70；其餘則使用數值 0.69 ~ 0.61。
# 白話音：常用度 > 0.40；最常用的讀音其值為 0.60，次常用的讀音其值為 0.50；其餘則使用數值 0.59 ~ 0.41。
# 其　它：常用度 > 0.00；使用數值 0.40 ~ 0.01；使用時機為：（1）方言地方腔；(2) 罕見發音；(3) 尚未查證屬文讀音或白話音 。
# ==========================================================
def han_ji_ca_piau_im(cursor, han_ji, reading_type="文讀音"):
    """
    根據漢字查詢其台羅音標及相關讀音資訊，並將台羅音標轉換為台語音標。
    若資料紀錄在`常用度`欄位儲存值為空值(NULL)，則將其視為 0，因此可排在查詢結果的最後。
    
    :param cursor: 數據庫游標
    :param han_ji: 欲查詢的漢字
    :param reading_type: 查詢的讀音類型，可以是 "文讀音"、"白話音" 或 "其它"
    :return: 包含讀音資訊的字典列表，包含台語音標、聲母、韻母、聲調。
    """

    if reading_type == "文讀音":
        reading_condition = "常用度 >= 0.61 AND 常用度 <= 1.0"
    elif reading_type == "白話音":
        reading_condition = "常用度 <= 0.60 AND 常用度 > 0.40"
    elif reading_type == "其它":
        reading_condition = "常用度 <= 0.40 AND 常用度 >= 0.01"
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
        台羅音標漢字庫
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
            台羅音標漢字庫
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
        tai_gi_im = tai_loo_im
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
# 自「台語音標+」，分析出：聲母、韻母、聲調
# ----------------------------------------------------------
# 【台羅音標】到【台語音標】的轉換規則
# tai_loo_to_tai_gi_mapping = {
#     'tsh': 'c',
#     'ts': 'z'
# }
# for tai_loo, tai_gi in tai_loo_to_tai_gi_mapping.items():
#     tai_gi_im = tai_gi_im.replace(tai_loo, tai_gi)
# ==========================================================
def split_zu_im(zu_im):
    # 聲母相容性轉換處理（將 tsh 轉換為 c；將 ts 轉換為 z）
    # zu_im = zu_im.replace("tsh", "c")   # 將 tsh 轉換為 c
    # zu_im = zu_im.replace("ts", "z")    # 將 ts  轉換為 z
    if zu_im.startswith("tsh") or zu_im.startswith("ch"):
        zu_im = zu_im.replace("tsh", "c", 1).replace("ch", "c", 1)  # 將 tsh, ch 轉換為 c
    elif zu_im.startswith("ts") or zu_im.startswith("c"):
        zu_im = zu_im.replace("ts", "z", 1).replace("c", "z", 1)  # 將 ts, c 轉換為 z

    # 定義聲母的正規表示式，包括常見的聲母，但不包括 m 和 ng
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

