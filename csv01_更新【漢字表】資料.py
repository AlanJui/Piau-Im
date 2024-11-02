import sqlite3

import pandas as pd

# 設定 CSV 文件的檔案名稱
csv_file_name = 'Documents\\廣韻\\廣韻v6_20241101-01_漢字表.csv'

# 使用 pandas 讀取 CSV 文件
data = pd.read_csv(csv_file_name)

# 假設您的 SQLite 資料庫文件名為 'Kong_Un.db'
database_path = 'Kong_Un.db'

# 建立連接到 SQLite 資料庫
conn = sqlite3.connect(database_path)
cursor = conn.cursor()

# 使用 data 迭代每一行，並更新漢字表
for index, row in data.iterrows():
    漢字號 = row['漢字號']
    上字號 = row['上字號']
    下字號 = row['下字號']
    標音 = row['標音']

    # 準備並執行更新語句
    cursor.execute(
        "UPDATE 漢字表 SET 字義號 = ?, 上字號 = ?, 下字號 = ?, 標音 = ? WHERE 漢字號 = ?",
        (漢字號, 上字號, 下字號, 標音, 漢字號)
    )
# 提交更改並關閉連接
conn.commit()
conn.close()

print("更新完成，所有的【下字號】欄位已根據 CSV 文件更新。")
