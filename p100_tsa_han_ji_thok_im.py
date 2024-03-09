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
# 在【十五音字庫】資料表查找【注音碼】
# =========================================================
def tsa_han_ji_thok_im(han_ji):
    global db_cursor
    sql = (
        "SELECT id, han_ji, tl_im, freq, siann, un, tiau, siann_bu, un_bu "
        f"FROM Sip_Ngoo_Im_Han_Ji_Tian "
        f"WHERE han_ji='{han_ji}' "
        "ORDER BY freq DESC;"
    )
    db_cursor.execute(sql)
    return db_cursor.fetchall()


if __name__ == "__main__":
    # 在程式開始時設定資料庫
    setup_database()

    # 查詢【字庫】資料庫
    result = tsa_han_ji_thok_im("鍵")
    print(result)

    # 在程式結束時關閉資料庫
    close_database()