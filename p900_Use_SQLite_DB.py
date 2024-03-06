import sqlite3

# 創建數據庫連接
conn = sqlite3.connect('.\\Kong_Un.db')

# 創建一個游標
cursor = conn.cursor()

# 執行 SQL 查詢
#cursor.execute("SELECT * FROM your_table WHERE your_condition")
# cursor.execute("SELECT * FROM Tshiat_Gu_Siong_Ji")
tshiat_gu_siong_ji = "徒"
cursor.execute("SELECT * FROM Tshiat_Gu_Siong_Ji WHERE 切語上字 LIKE ?", ('%' + tshiat_gu_siong_ji + '%',))

# 獲取查詢結果
results = cursor.fetchall()

# 關閉數據庫連接
conn.close()

