import sqlite3

from mod_標音 import convert_tl_to_tlpa

conn = sqlite3.connect("Ho_Lok_Ue.db")
cur = conn.cursor()
print("=== 原始（漢字, 台羅音標）重複組 ===")
dups = list(cur.execute("""
    SELECT 漢字, 台羅音標, COUNT(*) FROM 漢字庫
    GROUP BY 漢字, 台羅音標 HAVING COUNT(*) > 1
"""))
for row in dups:
    print("   ", row)
print("共", len(dups), "組")

print("=== 轉 TLPA 後才重複的組 ===")
rows = cur.execute("SELECT 漢字, 台羅音標 FROM 漢字庫").fetchall()
seen = {}
collide = []
for han_ji, tl in rows:
    key = (han_ji, convert_tl_to_tlpa(tl))
    seen.setdefault(key, []).append(tl)
for key, tls in seen.items():
    if len(tls) > 1 and len(set(tls)) > 1:
        collide.append((key, tls))
for item in collide[:20]:
    print("   ", item)
print("共", len(collide), "組")
conn.close()
