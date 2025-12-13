import sqlite3
import os
from dotenv import load_dotenv

# 載入環境變數
load_dotenv()
DB_HO_LOK_UE = os.getenv('DB_HO_LOK_UE', 'Ho_Lok_Ue.db')

print("=" * 70)
print("【漢字庫】資料表重整作業")
print("=" * 70)

# 連接資料庫
conn = sqlite3.connect(DB_HO_LOK_UE)
cursor = conn.cursor()

try:
    # 1. 檢查是否有重複資料
    print("\n步驟 1：檢查重複資料...")
    cursor.execute('''
        SELECT 漢字, 台羅音標, COUNT(*) as cnt
        FROM 漢字庫
        GROUP BY 漢字, 台羅音標
        HAVING COUNT(*) > 1
    ''')
    duplicates = cursor.fetchall()
    
    if not duplicates:
        print("✅ 資料表無重複資料，無需重整")
        exit(0)
    
    print(f"⚠️ 發現 {len(duplicates)} 組重複資料")
    
    # 顯示部分重複資料
    print("\n重複資料範例（前10組）：")
    for i, (han_ji, tai_lo, cnt) in enumerate(duplicates[:10], 1):
        print(f"  {i}. 漢字: {han_ji}, 台羅音標: {tai_lo}, 重複次數: {cnt}")
    
    # 2. 詢問用戶是否繼續
    print("\n" + "=" * 70)
    response = input("是否繼續清理重複資料？(y/n): ").strip().lower()
    if response != 'y':
        print("❌ 作業已取消")
        exit(0)
    
    # 3. 備份資料表
    print("\n步驟 2：備份資料表...")
    cursor.execute("DROP TABLE IF EXISTS 漢字庫_backup")
    cursor.execute('''
        CREATE TABLE 漢字庫_backup AS 
        SELECT * FROM 漢字庫
    ''')
    conn.commit()
    
    cursor.execute("SELECT COUNT(*) FROM 漢字庫_backup")
    backup_count = cursor.fetchone()[0]
    print(f"✅ 已備份 {backup_count} 筆資料到 漢字庫_backup")
    
    # 4. 清理重複資料（保留最新的一筆）
    print("\n步驟 3：清理重複資料...")
    print("策略：保留每組重複資料中【更新時間】最新、【識別號】最大的一筆")
    
    # 刪除重複資料，保留識別號最大的（通常是最新的）
    cursor.execute('''
        DELETE FROM 漢字庫
        WHERE 識別號 NOT IN (
            SELECT MAX(識別號)
            FROM 漢字庫
            GROUP BY 漢字, 台羅音標
        )
    ''')
    deleted_count = cursor.rowcount
    conn.commit()
    
    print(f"✅ 已刪除 {deleted_count} 筆重複資料")
    
    # 5. 驗證是否還有重複
    print("\n步驟 4：驗證清理結果...")
    cursor.execute('''
        SELECT 漢字, 台羅音標, COUNT(*) as cnt
        FROM 漢字庫
        GROUP BY 漢字, 台羅音標
        HAVING COUNT(*) > 1
    ''')
    remaining_duplicates = cursor.fetchall()
    
    if remaining_duplicates:
        print(f"⚠️ 仍有 {len(remaining_duplicates)} 組重複資料")
        for han_ji, tai_lo, cnt in remaining_duplicates[:5]:
            print(f"  - 漢字: {han_ji}, 台羅音標: {tai_lo}, 重複次數: {cnt}")
    else:
        print("✅ 已無重複資料")
    
    # 6. 嘗試建立或重建 UNIQUE INDEX
    print("\n步驟 5：建立/重建 UNIQUE INDEX...")
    
    # 先刪除舊的索引（如果存在）
    cursor.execute("DROP INDEX IF EXISTS idx_漢字_台羅音標")
    
    # 建立新的 UNIQUE INDEX
    cursor.execute('''
        CREATE UNIQUE INDEX idx_漢字_台羅音標 
        ON 漢字庫 (漢字, 台羅音標)
    ''')
    conn.commit()
    
    print("✅ UNIQUE INDEX 已成功建立")
    
    # 7. 顯示最終統計
    print("\n" + "=" * 70)
    print("清理完成統計：")
    print("=" * 70)
    
    cursor.execute("SELECT COUNT(*) FROM 漢字庫_backup")
    backup_count = cursor.fetchone()[0]
    
    cursor.execute("SELECT COUNT(*) FROM 漢字庫")
    current_count = cursor.fetchone()[0]
    
    print(f"原始資料筆數：{backup_count}")
    print(f"清理後筆數：  {current_count}")
    print(f"刪除筆數：    {backup_count - current_count}")
    print(f"保留率：      {current_count / backup_count * 100:.2f}%")
    
    print("\n✅ 資料表重整完成！")
    print("📝 備份資料表：漢字庫_backup（可用於還原）")
    
    # 8. 提供還原指令
    print("\n" + "=" * 70)
    print("如需還原資料，請執行以下 SQL：")
    print("=" * 70)
    print("DROP TABLE 漢字庫;")
    print("ALTER TABLE 漢字庫_backup RENAME TO 漢字庫;")
    
except sqlite3.IntegrityError as e:
    print(f"\n❌ 建立 UNIQUE INDEX 失敗：{e}")
    print("可能仍有重複資料未清理完成")
    conn.rollback()
    
except Exception as e:
    print(f"\n❌ 執行失敗：{e}")
    import traceback
    traceback.print_exc()
    conn.rollback()
    
finally:
    conn.close()

print("\n" + "=" * 70)
