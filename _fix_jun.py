"""檢視受損紀錄及其可能衝突的既有紀錄。"""
import sqlite3

conn = sqlite3.connect("Ho_Lok_Ue.db")
cur = conn.cursor()
for han_ji in ["媆", "閏", "嫩", "潤", "蝡", "韌"]:
    print(f"--- {han_ji} ---")
    for row in cur.execute(
        "SELECT 識別號, 台羅音標, 常用度, 摘要說明, 更新時間, 最近揀用時間 FROM 漢字庫 WHERE 漢字=?", (han_ji,)
    ):
        print("   ", row)
conn.close()
