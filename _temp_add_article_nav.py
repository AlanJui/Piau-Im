import os
import re

doc_dir = 'docs'
ignores = ['_test', '_archived', '金鋼經', 'assets']
ignore_files = ['index.html', '_template.html', 'output_from_excel.html']

nav_template = '''<nav class="main-nav">
  <ul>
    <li><a href="{rel_path}">回到首頁</a></li>
  </ul>
</nav>'''

for root, dirs, files in os.walk(doc_dir):
    dirs[:] = [d for d in dirs if d not in ignores]
    for fn in files:
        if not fn.endswith('.html') or fn in ignore_files:
            continue
            
        fp = os.path.join(root, fn)
        try:
            with open(fp, 'r', encoding='utf-8') as f:
                html = f.read()
                
            # 計算相對路徑
            rel_level = root.replace('\\\\', '/').count('/') - doc_dir.count('/')
            rel_prefix = '../' * rel_level
            rel_path = rel_prefix + 'index.html'
            
            nav_html = nav_template.replace('{rel_path}', rel_path)
            
            # 清除舊的 floating button 與舊的 nav
            html = re.sub(r'<a href=\"[^\"]*index.html\" class=\"floating-home-btn\"[^>]*>.*?</a>', '', html, flags=re.DOTALL)
            html = re.sub(r'<nav class=\"main-nav\">.*?</nav>', '', html, flags=re.DOTALL)
            
            # 在 <body> 後加入 top nav
            # 在 </body> 前加入 bottom nav
            if '<body' in html:
                # 替換前先確保標籤格式
                html = re.sub(r'(<body[^>]*>)', r'\1\n' + nav_html, html, count=1)
                html = re.sub(r'(</body>)', nav_html + r'\n\1', html, count=1)
                
            with open(fp, 'w', encoding='utf-8') as f:
                f.write(html)
        except Exception as e:
            print(f"Failed on {fn}: {str(e)}")

print("Done inserting nav to article pages.")
