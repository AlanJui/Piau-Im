import sqlite3

# 連接到 SQLite 數據庫
conn = sqlite3.connect('.\\Kong_Un.db')

# 創建一個游標對象
cur = conn.cursor()

# 執行 SQL 查詢以獲取 "Lui_Tsip_Nga_Siok_Thong" 表的所有記錄
cur.execute("SELECT * FROM Lui_Tsip_Nga_Siok_Thong")

# 獲取所有記錄
records = cur.fetchall()

# 對於每一條記錄
for record in records:
    # 更新 "Han-Ji-Phing-Im-Ji-Tian" 表
    cur.execute("""
        UPDATE Han-Ji-Phing-Im-Ji-Tian
        SET NST_ID = ?, Siann = ?, Un = ?, Tiau = ?
        WHERE Han_Ji = ?
    """, (record[0], record[1], record[2], record[3], record[4]))

# 提交事務
conn.commit()

# 關閉連接
conn.close()