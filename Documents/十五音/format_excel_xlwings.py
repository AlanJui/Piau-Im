import xlwings as xw

# 開啟工作簿
workbook = xw.Book("十五音_工作簿.xlsx")  # 替換成你的檔名
sheet = workbook.sheets[0]  # 取得第一個工作表

# 讀取 A1 到 M1 的值
values = sheet.range("A1:M1").value

# 格式化成所需字串
formatted_string = ", ".join([f"'{value}'" for value in values])

# 輸出結果
print(f'\n組合 fields 字串：\n\n{formatted_string}\n')
# print(formatted_string)

#---------------------------------------------------------------------
# 製作 SQL Script
#---------------------------------------------------------------------

# 讀取 A1 到 M1 的值
fields = sheet.range("A1:M1").value

# 將欄位名稱格式化成 SELECT 語句
select_clause = ", ".join(fields)

# 組合完整的 SQL Script
sql_script = f"""
SELECT {select_clause}
FROM 漢字表
WHERE 漢字 = ?
ORDER BY COALESCE(常用度, 0) DESC;
"""

# 輸出結果
# print(sql_script)
print(f"SELECT 欄位：\n\n{select_clause}\n")
print(f"SELECT 語句：\n\n{sql_script}\n")