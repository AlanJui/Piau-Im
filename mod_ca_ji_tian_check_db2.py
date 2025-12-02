import sqlite3

conn = sqlite3.connect('Ho_Lok_Ue.db')
cursor = conn.cursor()

# 檢查「漢字庫」表的結構
print("表「漢字庫」的結構：")
cursor.execute("PRAGMA table_info(漢字庫)")
columns = cursor.fetchall()
for col in columns:
    print(f"  {col[1]:20} {col[2]:10}")

# 查詢一筆資料看看
print("\n資料範例（前 5 筆）：")
cursor.execute("SELECT * FROM 漢字庫 LIMIT 5")
rows = cursor.fetchall()
for row in rows:
    print(f"  {row}")

conn.close()
