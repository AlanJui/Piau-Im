import os

# 要生成超連結的目錄
docs_directory = 'docs'

# 生成超連結的檔案名稱
index_file = os.path.join(docs_directory, 'index.html')

# 開始 index.html 檔案
with open(index_file, 'w', encoding='utf-8') as f:  # 指定 UTF-8 編碼
    # 寫入 HTML 開頭，並指定使用 UTF-8 編碼
    f.write('''<html>
<head>
    <meta charset="UTF-8">
    <title>文章清單</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            margin: 20px;
        }
        h1 {
            color: #333;
        }
        a {
            color: #0066cc;
            text-decoration: none;
        }
        a:hover {
            text-decoration: underline;
        }
    </style>
</head>
<body>
    <h1>文章清單</h1>
    <div id="articles"></div>
    <script src="./assets/javascripts/main.js"></script>\n''')
    
    # 文章名稱與注音方式的字典
    articles = {}
    
    # 遍歷目錄下的檔案
    for filename in os.listdir(docs_directory):
        # 排除 index.html 和 _template.html 檔案
        if filename not in ['index.html', '_template.html']:
            # 取得檔案名稱 (不包含副檔名)
            article_and_phonetic = os.path.splitext(filename)[0]
            
            # 切割檔案名稱，取得文章名稱與注音方式
            parts = article_and_phonetic.split('_')
            if len(parts) == 2:
                article, phonetic_method = parts
                
                # 如果文章名稱不在 articles 字典中，則添加
                if article not in articles:
                    articles[article] = []
                
                # 如果注音方式不在文章名稱對應的列表中，則添加
                if phonetic_method not in articles[article]:
                    articles[article].append(phonetic_method)
    
    # 遍歷每個文章
    for article, phonetic_methods in articles.items():
        # 生成文章標題
        f.write(f'    <h2>{article}</h2>\n')
        
        # 遍歷每個注音方式
        for phonetic_method in phonetic_methods:
            # 組合檔案名稱
            filename = f"{article}_{phonetic_method}.html"
            
            # 生成超連結
            link = f'    <a href="{filename}">{phonetic_method}</a><br>\n'
            
            # 寫入超連結到 index.html
            f.write(link)
    
    # 寫入 HTML 結尾
    f.write('</body>\n</html>\n')
