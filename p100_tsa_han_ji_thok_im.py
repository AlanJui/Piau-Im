import sqlite3


# 宣告全域變數
from config_dev_env import DATABASE

db_connection = None
db_cursor = None

# =========================================================="
# 連接資料庫
# =========================================================="
def setup_database():
    global db_connection, db_cursor
    db_connection = sqlite3.connect(DATABASE)
    db_cursor = db_connection.cursor()

# ==========================================================
# 關閉資料庫
# ==========================================================
def close_database():
    global db_connection
    db_connection.close()

# =========================================================
# 查漢字讀音
# =========================================================
def tsa_han_ji_thok_im(han_ji):
    global db_cursor
    sql = (
        "SELECT H.Han_Ji, H.TL_Phing_Im, H.freq, L.Siann, L.Un, L.Tiau "
        "FROM Han_Ji_Phing_Im_Ji_Tian AS H "
        "JOIN Lui_Tsip_Nga_Siok_Thong AS L ON H.NST_ID = L.ID "
        f"WHERE H.Han_Ji='{han_ji}';"
    )
    db_cursor.execute(sql)
    results = db_cursor.fetchall()

    # 將查詢結果轉換為字典
    result_dict = {}
    for result in results:
        if result[0] not in result_dict:
            result_dict[result[0]] = []
        result_dict[result[0]].append({"TL_Phing_Im": result[1], "freq": result[2], "Siann": result[3], "Un": result[4], "Tiau": result[5]})

    return result_dict

# =========================================================
# 查詢十五音字典
# =========================================================
def Tsa_Sip_Ngoo_Im(han_ji):
    global db_cursor
    sql = (
        "SELECT id, han_ji, tl_im, freq, siann, un, tiau, siann_bu, un_bu "
        f"FROM Sip_Ngoo_Im_Han_Ji_Tian "
        f"WHERE han_ji='{han_ji}' "
        "ORDER BY freq DESC;"
    )
    db_cursor.execute(sql)
    return db_cursor.fetchall()

# =========================================================
# 查詢彙集雅俗通十五音字典
# =========================================================
def Tsa_Nga_Siok_Thong(han_ji):
    global db_cursor
    sql = (
        "SELECT * "
        f"FROM Lui_Tsip_Nga_Siok_Thong "
        f"WHERE Ji='{han_ji}';"
    )
    db_cursor.execute(sql)
    return db_cursor.fetchall()



if __name__ == "__main__":
    # 在程式開始時設定資料庫
    setup_database()

    # 查詢【字庫】資料庫
    han_ji = "鍵"
    result = tsa_han_ji_thok_im(han_ji)
    print(result)
    
    result2 = Tsa_Sip_Ngoo_Im(han_ji)
    print(result2)

    result3 = Tsa_Nga_Siok_Thong(han_ji)
    print(result3)

    # 在程式結束時關閉資料庫
    close_database()