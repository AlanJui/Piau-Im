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
    ORDER BY COALESCE(常用度, 0) DESC;
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


def TL_Tng_Zu_Im(siann_bu, un_bu, siann_tiau, cursor):
    """
    根據傳入的台語音標聲母、韻母、聲調，轉換成對應的方音符號
    :param siann_bu: 聲母 (台語音標)
    :param un_bu: 韻母 (台語音標)
    :param siann_tiau: 聲調 (台語音標中的數字)
    :param cursor: 數據庫游標
    :return: 包含方音符號的字典
    """

    # 查詢聲母表，將台語音標的聲母轉換成方音符號
    cursor.execute("SELECT 方音聲母 FROM 聲母對照表 WHERE 台語音標聲母 = ?", (siann_bu,))
    siann_bu_result = cursor.fetchone()
    if siann_bu_result:
        zu_im_siann_bu = siann_bu_result[0]  # 取得方音符號
    else:
        zu_im_siann_bu = ''  # 無聲母的情況

    # 查詢韻母表，將台語音標的韻母轉換成方音符號
    # cursor.execute("SELECT 方音符號 FROM 韻母表 WHERE 台語音標 = ?", (un_bu,))
    cursor.execute("SELECT 方音韻母 FROM 韻母對照表 WHERE 台語音標韻母 = ?", (un_bu,))
    un_bu_result = cursor.fetchone()
    if un_bu_result:
        zu_im_un_bu = un_bu_result[0]  # 取得方音符號
    else:
        zu_im_un_bu = ''

    # 查詢聲調表，將台語音標的聲調轉換成方音符號
    cursor.execute("SELECT 方音符號 FROM 聲調表 WHERE 台羅八聲調 = ?", (siann_tiau,))
    siann_tiau_result = cursor.fetchone()
    if siann_tiau_result:
        zu_im_siann_tiau = siann_tiau_result[0]  # 取得方音符號
    else:
        zu_im_siann_tiau = ''

    #=======================================================================
    # 【聲母】校調
    #
    # 齒間音【聲母】：ㄗ、ㄘ、ㄙ、ㆡ，若其後所接【韻母】之第一個符號亦為：ㄧ、ㆪ時，須變改
    # 為：ㄐ、ㄑ、ㄒ、ㆢ。
    #-----------------------------------------------------------------------
    # 參考 RIME 輸入法如下規則：
    # - xform/ㄗ(ㄧ|ㆪ)/ㄐ$1/
    # - xform/ㄘ(ㄧ|ㆪ)/ㄑ$1/
    # - xform/ㄙ(ㄧ|ㆪ)/ㄒ$1/
    # - xform/ㆡ(ㄧ|ㆪ)/ㆢ$1/
    #=======================================================================
    
    # 比對聲母是否為 ㄗ、ㄘ、ㄙ、ㆡ，且韻母的第一個符號是 ㄧ 或 ㆪ
    if siann_bu == 'z' and (un_bu[0] == 'i' or un_bu == 'inn'):
        zu_im_siann_bu = 'ㄐ'
    elif siann_bu == 'c' and (un_bu[0] == 'i' or un_bu == 'inn'):
        zu_im_siann_bu = 'ㄑ'
    elif siann_bu == 's' and (un_bu[0] == 'i' or un_bu == 'inn'):
        zu_im_siann_bu = 'ㄒ'
    elif siann_bu == 'j' and (un_bu[0] == 'i' or un_bu == 'inn'):
        zu_im_siann_bu = 'ㆢ'

    return {
        '注音符號': f"{zu_im_siann_bu}{zu_im_un_bu}{zu_im_siann_tiau}",
        '聲母': zu_im_siann_bu,
        '韻母': zu_im_un_bu,
        '聲調': zu_im_siann_tiau
    }



def Kong_Un_Tng_Tai_Loo(廣韻調名):
    """
    將【廣韻調名】轉換成【台羅聲調】號
    清平(1)、清上(2)、清去(3)、清入(4)
    濁平(5)、濁上(6)、濁去(7)、濁入(8)
    
    :param 廣韻調名: 廣韻的調名
    :return: 對應的台羅聲調號
    """
    調名對照 = {
        "清平": 1,
        "清上": 2,
        "清去": 3,
        "清入": 4,
        "濁平": 5,
        "濁上": 6,
        "濁去": 7,
        "濁入": 8
    }
    return 調名對照.get(廣韻調名, None)


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
    if zu_im.startswith("tsh"):
        zu_im = zu_im.replace("tsh", "c", 1)  # 將 tsh, ch 轉換為 c
    elif zu_im.startswith("ts"):
        zu_im = zu_im.replace("ts", "z", 1)  # 將 ts, c 轉換為 z

    # 定義聲母的正規表示式，包括常見的聲母，並加入 Ø 表示無聲母
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
        # 使用正規表示式來匹配聲母，包括 Ø 符號
        siann_bu_match = siann_bu_pattern.match(zu_im)
        if siann_bu_match:
            siann_bu = siann_bu_match.group()  # 找到聲母或無聲母（Ø）
            # 如果聲母是 Ø，將其轉換為空字串，並確保不影響 un_bu
            if siann_bu == "Ø":
                siann_bu = ""
                un_bu = zu_im[1:-1]  # 跳過 Ø 從第二個字符開始取韻母
            else:
                un_bu = zu_im[len(siann_bu):-1]  # 正常處理韻母部分
        else:
            siann_bu = ""  # 沒有匹配到聲母，聲母為空字串
            un_bu = zu_im[:-1]  # 韻母是剩下的部分，去掉最後的聲調

        tiau = zu_im[-1]  # 最後一個字符是聲調

    result += [siann_bu]
    result += [un_bu]
    result += [tiau]
    return result

# def split_zu_im(zu_im):
#     # 聲母相容性轉換處理（將 tsh 轉換為 c；將 ts 轉換為 z）
#     # zu_im = zu_im.replace("tsh", "c")   # 將 tsh 轉換為 c
#     # zu_im = zu_im.replace("ts", "z")    # 將 ts  轉換為 z
#     if zu_im.startswith("tsh") or zu_im.startswith("ch"):
#         zu_im = zu_im.replace("tsh", "c", 1).replace("ch", "c", 1)  # 將 tsh, ch 轉換為 c
#     elif zu_im.startswith("ts") or zu_im.startswith("c"):
#         zu_im = zu_im.replace("ts", "z", 1).replace("c", "z", 1)  # 將 ts, c 轉換為 z

#     # 定義聲母的正規表示式，包括常見的聲母，但不包括 m 和 ng
#     siann_bu_pattern = re.compile(r"(b|c|z|g|h|j|kh|k|l|m(?!\d)|ng(?!\d)|n|ph|p|s|th|t|Ø)")
    
#     # 韻母為 m 或 ng 這種情況的正規表示式 (m\d 或 ng\d)
#     un_bu_as_m_or_ng_pattern = re.compile(r"(m|ng)\d")

#     result = []

#     # 首先檢查是否是 m 或 ng 當作韻母的特殊情況
#     if un_bu_as_m_or_ng_pattern.match(zu_im):
#         siann_bu = ""  # 沒有聲母
#         un_bu = zu_im[:-1]  # 韻母是 m 或 ng
#         tiau = zu_im[-1]  # 聲調是最後一個字符
#     else:
#         # 使用正規表示式來匹配聲母
#         siann_bu_match = siann_bu_pattern.match(zu_im)
#         if siann_bu_match:
#             siann_bu = siann_bu_match.group()  # 找到聲母
#             un_bu = zu_im[len(siann_bu):-1]  # 韻母部分
#         else:
#             siann_bu = ""  # 沒有匹配到聲母，聲母為空字串
#             un_bu = zu_im[:-1]  # 韻母是剩下的部分，去掉最後的聲調

#         tiau = zu_im[-1]  # 最後一個字符是聲調

#     result += [siann_bu]
#     result += [un_bu]
#     result += [tiau]
#     return result

def dict_to_str(zu_im_hu_ho):
    return f"{zu_im_hu_ho['聲母']}{zu_im_hu_ho['韻母']}{zu_im_hu_ho['聲調']}"

