import sqlite3

# 連接資料庫
conn = sqlite3.connect('Ho_Lok_Ue.db')
cursor = conn.cursor()

# 查詢所有表名
cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
tables = cursor.fetchall()

print("資料庫中的表：")
for table in tables:
    print(f"  - {table[0]}")

# 如果有表，顯示第一個表的結構
if tables:
    first_table = tables[0][0]
    print(f"\n表 '{first_table}' 的結構：")
    cursor.execute(f"PRAGMA table_info({first_table})")
    columns = cursor.fetchall()
    for col in columns:
        print(f"  {col[1]} ({col[2]})")

conn.close()
