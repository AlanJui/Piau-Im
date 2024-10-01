# 雅俗通十五音漢字典查詢模組
import re
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


# ==========================================================
# 用 `漢字` 查詢《雅俗通十五音》的標音
# ==========================================================
def han_ji_cha_piau_im(cursor, han_ji):
    """
    根據漢字查詢其讀音資訊。若資料紀錄在`常用度`欄位儲存值為空值(NULL)，
    則將其視為 0，因此可排在查詢結果的最後。
    
    :param cursor: 數據庫游標
    :param han_ji: 欲查詢的漢字
    :return: 包含讀音資訊的字典列表
    """

    query = """
    SELECT 
        HJT.[識別號] AS [識別號],
        HJT.[聲母] AS [十五音聲母],
        HJT.[韻母] AS [十五音韻母],
        HJT.[聲調] AS [十五音聲調],
        HJT.[常用度] AS [常用度],
        SBP.[台語音標] AS [聲母台語音標],
        UBP.[台語音標] AS [韻母台語音標],
        SBP.[方音符號] AS [聲母方音符號],
        UBP.[方音符號] AS [韻母方音符號],
        STP.[台羅八聲調] AS [八聲調]
    FROM 
        Han_Ji_Tian HJT
    LEFT JOIN 
        Siann_Bu_Piau SBP ON HJT.[聲母識別號] = SBP.[識別號]
    LEFT JOIN 
        Un_Bu_Piau UBP ON HJT.[韻母識別號] = UBP.[識別號]
    LEFT JOIN 
        Siann_Tiau_Piau STP ON HJT.[聲調識別號] = STP.[識別號]
    WHERE 
        HJT.[漢字] = ?
    ORDER BY 
        COALESCE(HJT.[常用度], 0) DESC;
    """

    cursor.execute(query, (han_ji,))
    results = cursor.fetchall()
    
    # 將結果轉換為字典列表
    fields = [
        '識別號', '十五音聲母', '十五音韻母', '十五音聲調', '常用度', 
        '聲母台語音標', '韻母台語音標', '聲母方音符號', '韻母方音符號', '八聲調'
    ]
    
    return [dict(zip(fields, result)) for result in results]



# ==========================================================
# 自漢字的「注音碼」，分析出：聲母、韻母、調號
# ==========================================================
def split_cu_im(cu_im):
    sing_bu_pattern = re.compile(r"(b|ch|c|g|h|j|kh|k|l|m|ng|n|ph|p|s|th|t|Ø)")
    result = []

    sing_bu = sing_bu_pattern.match(cu_im).group()
    un_bu = cu_im[len(sing_bu) : len(cu_im) - 1]
    tiau = cu_im[len(cu_im) - 1]

    result += [sing_bu]
    result += [un_bu]
    result += [tiau]
    return result


if __name__ == "__main__":
    # 在所有測試開始前，連接資料庫
    conn = sqlite3.connect('Nga_Siok_Thong_Sip_Ngoo_Im.db')  # 替換為實際資料庫路徑
    cursor = conn.cursor()

    #--------------------------------------------------
    # 測試 `han_ji_cha_piau_im` 函數
    #--------------------------------------------------
    han_ji = '不'
    result = han_ji_cha_piau_im(cursor, han_ji)
    print(result)
    assert result[0]['十五音聲母'] == '邊', "轉換錯誤！"
    assert result[0]['十五音韻母'] == '君', "轉換錯誤！"
    assert result[0]['十五音聲調'] == '上入', "轉換錯誤！"
    assert result[0]['八聲調'] == 4, "轉換錯誤！"
    assert result[0]['聲母台語音標'] == 'p', "轉換錯誤！"
    assert result[0]['韻母台語音標'] == 'ut', "轉換錯誤！"
    assert result[0]['聲母方音符號'] == 'ㄅ', "轉換錯誤！"
    assert result[0]['韻母方音符號'] == 'ㄨㆵ', "轉換錯誤！"

    # 在所有測試結束後，關閉資料庫連接
    conn.close()