import os

# 要生成超連結的目錄
docs_directory = 'docs'

# 生成超連結的檔案名稱
index_file = os.path.join(docs_directory, 'index.html')

# HTML 模板檔案名稱
template_file = os.path.join(docs_directory, '_template.html')

# 開始 index.html 檔案
with open(index_file, 'w', encoding='utf-8') as f:  # 指定 UTF-8 編碼
    # 讀取 HTML 模板內容
    with open(template_file, 'r', encoding='utf-8') as template:
        template_content = template.read()
        
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
        
        # 生成文章清單的 HTML 內容
        articles_html = ''
        for article, phonetic_methods in articles.items():
            articles_html += f'<h2>{article}</h2>\n'
            articles_html += '<ul>\n'
            for phonetic_method in phonetic_methods:
                filename = f"{article}_{phonetic_method}.html"
                articles_html += f'<li><a href="{filename}">{phonetic_method}</a></li>\n'
            articles_html += '</ul>\n'
        
        # 將文章清單的 HTML 內容插入模板中
        template_content = template_content.replace('{articles_placeholder}', articles_html)
        
        # 將修改後的模板內容寫入 index.html
        f.write(template_content)
