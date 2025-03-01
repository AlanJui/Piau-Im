import re

"""
用 `漢字` 查詢《廣韻》的標音
"""
def han_ji_ca_piau_im(cursor, han_ji, hue_im="文讀音"):
    """
    根據漢字查詢其讀音資訊。 若資料紀錄在`常用度`欄位儲存值為空值(NULL)
    ，則將其視為 0，因此可排在查詢結果的最後。

    :param cursor: 數據庫游標
    :param han_ji: 欲查詢的漢字
    :return: 包含讀音資訊的字典列表

    SELECT *
    FROM 漢字檢視
    WHERE 漢字 = ?
    ORDER BY COALESCE(常用度, 0) DESC;
    """

    query = """
    SELECT 漢字號, 漢字, 漢字標音, 常用度, 上字, 下字, 字義解釋,
           七聲類, 發音部位, 聲母, 聲母標音, 清濁, 發送收,
           韻母, 調, 韻母標音, 攝, 韻系列號, 韻系, 韻目, 目次, 等呼, 等, 呼
    FROM 漢字檢視
    WHERE 漢字 = ?
    ORDER BY COALESCE(常用度, 0) DESC;
    """
    cursor.execute(query, (han_ji,))
    results = cursor.fetchall()

    # 將結果轉換為字典列表
    fields = [
        '字號', '漢字', '標音', '常用度', '上字', '下字', '字義解釋',
        '七聲類', '發音部位', '聲母', '聲母標音', '清濁', '發送收',
        '韻母', '調', '韻母標音', '攝', '韻系列號', '韻系', '韻目', '目次', '等呼', '等', '呼'
    ]
    return [dict(zip(fields, result)) for result in results]


def ca_siann_bu_piau_im(cursor, siann_bu):
    """
    根據聲母標音查詢其國際音標聲母和方音聲母。

    :param cursor: 數據庫游標
    :param siann_bu: 欲查詢的聲母標音（台語音標）
    :return: 包含國際音標聲母和方音聲母的字典列表

    SELECT 國際音標, 方音聲母
    FROM 聲母對照表
    WHERE 台語音標 = ?;
    """

    query = """
    SELECT 國際音標, 方音符號
    FROM 聲母對照表
    WHERE 台語音標 = ?;
    """
    cursor.execute(query, (siann_bu,))
    results = cursor.fetchall()

    fields = ['國際音標', '方音符號']
    return [dict(zip(fields, result)) for result in results]


def ca_un_bu_piau_im(cursor, un_bu):
    """
    根據韻母標音查詢其國際音標韻母和方音韻母。

    :param cursor: 數據庫游標
    :param un_bu: 欲查詢的韻母標音（台語音標）
    :return: 包含國際音標和方音韻母的字典列表

    SELECT 國際音標, 方音韻母
    FROM 韻母對照表
    WHERE 台語音標 = ?;
    """

    query = """
    SELECT 國際音標, 方音符號
    FROM 韻母對照表
    WHERE 台語音標 = ?;
    """
    cursor.execute(query, (un_bu,))
    results = cursor.fetchall()

    fields = ['國際音標', '方音符號']
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


# ==========================================================
# 台羅音標轉換為【十五音】切韻標音
# ==========================================================
def TL_Tng_Sip_Ngoo_Im(siann_bu, un_bu, siann_tiau, cursor):
    """
    根據傳入的台羅音標聲母、韻母、聲調，轉換成對應的十五音切韻標音
    :param siann_bu: 聲母 (台羅音標)
    :param un_bu: 韻母 (台羅音標)
    :param siann_tiau: 聲調 (台羅音標中的數字)
    :param cursor: 資料庫游標
    :return: 包含十五音標音的字典
    """

    # 查詢聲母對照表，將台羅音標的聲母轉換成十五音
    cursor.execute("SELECT 十五音 FROM 聲母對照表 WHERE 台語音標 = ?", (siann_bu,))
    siann_bu_result = cursor.fetchone()
    if siann_bu_result:
        sni_siann_bu = siann_bu_result[0]  # 取得十五音
    else:
        sni_siann_bu = ''  # 無聲母的情況

    # 查詢韻母對照表，將台羅音標的韻母轉換成十五音
    cursor.execute("SELECT 十五音 FROM 韻母對照表 WHERE 台語音標 = ?", (un_bu,))
    un_bu_result = cursor.fetchone()
    if un_bu_result:
        sni_un_bu = un_bu_result[0]  # 取得十五音
    else:
        sni_un_bu = ''

    # 查詢聲調對照表，將台羅音標的聲調轉換成十五音的調符
    cursor.execute("SELECT 十五音聲調 FROM 聲調對照表 WHERE 台羅調號 = ?", (siann_tiau,))
    siann_tiau_result = cursor.fetchone()
    if siann_tiau_result:
        sni_siann_tiau = siann_tiau_result[0]  # 取得十五音調符
    else:
        sni_siann_tiau = ''

    #=======================================================================
    # 【聲母】校調（如有需要，可根據十五音的規則進行聲母的調整）
    #-----------------------------------------------------------------------
    # 此處暫無特別的聲母校調規則，如有需要請添加
    #=======================================================================

    return {
        '標音': f"{sni_un_bu}{sni_siann_tiau}{sni_siann_bu}",
        '聲母': sni_siann_bu,
        '韻母': sni_un_bu,
        '聲調': sni_siann_tiau
    }


# 台羅音標轉換為方音符號
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
    cursor.execute("SELECT 方音符號 FROM 聲母對照表 WHERE 台語音標 = ?", (siann_bu,))
    siann_bu_result = cursor.fetchone()
    if siann_bu_result:
        zu_im_siann_bu = siann_bu_result[0]  # 取得方音符號
    else:
        zu_im_siann_bu = ''  # 無聲母的情況

    # 查詢韻母表，將台語音標的韻母轉換成方音符號
    # cursor.execute("SELECT 方音符號 FROM 韻母表 WHERE 台語音標 = ?", (un_bu,))
    cursor.execute("SELECT 方音符號 FROM 韻母對照表 WHERE 台語音標 = ?", (un_bu,))
    un_bu_result = cursor.fetchone()
    if un_bu_result:
        zu_im_un_bu = un_bu_result[0]  # 取得方音符號
    else:
        zu_im_un_bu = ''

    # 查詢聲調表，將台語音標的聲調轉換成方音符號
    cursor.execute("SELECT 方音符號調符 FROM 聲調對照表 WHERE 台羅調號 = ?", (siann_tiau,))
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


def Kong_Un_Siann_Tiau_Tng_Tai_Loo(廣韻調名):
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
        "濁上": 2,
        "濁去": 7,
        "濁入": 8
    }
    return 調名對照.get(廣韻調名, None)


# def Cu_Hong_Im_Hu_Ho(tai_lo_tiau_ho):
    """
    取方音符號：將【台羅調號】轉換成【方音符號調號】
    :param tai_lo_tiau_ho: 台羅調號
    :return: 對應的方音符號調號
    """
    方音符號調號 = {
        1: '',
        2: 'ˋ',
        3: '˪',
        4: '',
        5: 'ˊ',
        6: 'ˋ',
        7: '˫',
        8: '˙'
    }
    return 方音符號調號.get(tai_lo_tiau_ho, None)

