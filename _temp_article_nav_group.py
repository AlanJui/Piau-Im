import os
import re

doc_dir = 'docs'
ignores = ['_test', '_archived', '金鋼經', 'assets']
ignore_files = ['index.html', '_template.html', 'output_from_excel.html']

def parse_filename(filename):
    name_no_ext = os.path.splitext(filename)[0]
    if "_" in name_no_ext:
        parts = name_no_ext.split("_")
        phonetic_method = parts[-1]
        article = "_".join(parts[:-1])
    else:
        match = re.search(r"^(.*【.*?】)(.*)$", name_no_ext)
        if match and match.group(2).strip():
            article = match.group(1)
            phonetic_method = match.group(2)
        else:
            article = name_no_ext
            phonetic_method = "開啟"
    return article, phonetic_method

for root, dirs, files in os.walk(doc_dir):
    dirs[:] = [d for d in dirs if d not in ignores]
    
    # gather files in this directory
    valid_files = [f for f in files if f.endswith('.html') and f not in ignore_files]
    
    # group by article base name
    groups = {}
    for fn in valid_files:
        art, phon = parse_filename(fn)
        if art not in groups:
            groups[art] = []
        groups[art].append((fn, phon))
    
    # process each file
    for fn in valid_files:
        fp = os.path.join(root, fn)
        art, current_phon = parse_filename(fn)
        
        siblings = groups.get(art, [])
        # siblings usually list of (filename, phonetic)
        
        # calculate relative path to index
        rel_level = root.replace('\\\\', '/').count('/') - doc_dir.count('/')
        rel_prefix = '../' * rel_level
        rel_path = rel_prefix + 'index.html'
        
        # Build nav html
        nav_html = '<nav class="main-nav">\n  <ul>\n'
        nav_html += f'    <li><a href="{rel_path}" class="nav-home">🏠 首頁</a></li>\n'
        
        # sort siblings so they appear in a consistent order
        siblings.sort(key=lambda x: x[1])
        
        for sib_fn, sib_phon in siblings:
            if sib_fn == fn:
                # current page
                nav_html += f'    <li><span class="nav-current">{sib_phon}</span></li>\n'
            else:
                import urllib.parse
                safe_url = urllib.parse.quote(sib_fn)
                nav_html += f'    <li><a href="{safe_url}" class="nav-sibling">{sib_phon}</a></li>\n'
                
        nav_html += '  </ul>\n</nav>'
        
        try:
            with open(fp, 'r', encoding='utf-8') as f:
                html = f.read()
                
            # 清除舊的 floating button 與舊的 nav
            html = re.sub(r'<a href=\"[^\"]*index.html\" class=\"floating-home-btn\"[^>]*>.*?</a>', '', html, flags=re.DOTALL)
            html = re.sub(r'<nav class=\"main-nav\">.*?</nav>', '', html, flags=re.DOTALL)
            
            # 插入新 nav
            if '<body' in html:
                html = re.sub(r'(<body[^>]*>)', r'\1\n' + nav_html, html, count=1)
                html = re.sub(r'(</body>)', nav_html + r'\n\1', html, count=1)
                
            with open(fp, 'w', encoding='utf-8') as f:
                f.write(html)
        except Exception as e:
            print(f"Failed on {fn}: {str(e)}")

print("Done building grouped nav.")
