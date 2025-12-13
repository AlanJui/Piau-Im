import sqlite3
import os
from dotenv import load_dotenv

# 載入環境變數
load_dotenv()
DB_HO_LOK_UE = os.getenv('DB_HO_LOK_UE', 'Ho_Lok_Ue.db')

# 連接資料庫
conn = sqlite3.connect(DB_HO_LOK_UE)
cursor = conn.cursor()

try:
    # 建立 UNIQUE INDEX
    cursor.execute('''
        CREATE UNIQUE INDEX IF NOT EXISTS idx_漢字_台羅音標 
        ON 漢字庫 (漢字, 台羅音標)
    ''')
    conn.commit()
    print('✅ UNIQUE INDEX 建立成功！')
    
    # 驗證索引是否存在
    cursor.execute('''
        SELECT sql FROM sqlite_master 
        WHERE type='index' AND name='idx_漢字_台羅音標'
    ''')
    result = cursor.fetchone()
    
    if result:
        print(f'\n索引定義：\n{result[0]}')
    else:
        print('⚠️ 索引未找到')
        
    # 檢查資料表是否有重複資料
    cursor.execute('''
        SELECT 漢字, 台羅音標, COUNT(*) as cnt
        FROM 漢字庫
        GROUP BY 漢字, 台羅音標
        HAVING COUNT(*) > 1
    ''')
    duplicates = cursor.fetchall()
    
    if duplicates:
        print(f'\n⚠️ 發現 {len(duplicates)} 組重複資料：')
        for han_ji, tai_lo, cnt in duplicates[:10]:  # 只顯示前10筆
            print(f'  - 漢字: {han_ji}, 台羅音標: {tai_lo}, 重複次數: {cnt}')
        if len(duplicates) > 10:
            print(f'  ... 還有 {len(duplicates) - 10} 組重複資料')
    else:
        print('\n✅ 資料表無重複資料')
        
except sqlite3.IntegrityError as e:
    print(f'❌ 建立索引失敗（可能有重複資料）：{e}')
    print('\n正在查詢重複資料...')
    cursor.execute('''
        SELECT 漢字, 台羅音標, COUNT(*) as cnt
        FROM 漢字庫
        GROUP BY 漢字, 台羅音標
        HAVING COUNT(*) > 1
    ''')
    duplicates = cursor.fetchall()
    
    if duplicates:
        print(f'\n發現 {len(duplicates)} 組重複資料：')
        for han_ji, tai_lo, cnt in duplicates:
            print(f'  - 漢字: {han_ji}, 台羅音標: {tai_lo}, 重複次數: {cnt}')
except Exception as e:
    print(f'❌ 執行失敗：{e}')
finally:
    conn.close()
