import os

# 要生成超連結的目錄
docs_directory = 'docs'

# 生成超連結的檔案名稱
index_file = os.path.join(docs_directory, 'index.html')

# 開始 index.html 檔案
with open(index_file, 'w', encoding='utf-8') as f:  # 指定 UTF-8 編碼
    # 寫入 HTML 開頭
    f.write('<html><head><title>Index</title></head>\n<body>')
    
    # 遍歷目錄下的檔案
    for filename in os.listdir(docs_directory):
        # 排除 index.html 和 _template.html 檔案
        if filename not in ['index.html', '_template.html']:
            # 只處理 HTML 檔案
            if filename.endswith('.html'):
                # 生成超連結
                link = f'<a href="{filename}">{filename}</a><br>\n'
                # 寫入超連結到 index.html
                f.write(link)
    
    # 寫入 HTML 結尾
    f.write('</body></html>')