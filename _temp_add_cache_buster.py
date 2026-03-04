import os
import re

doc_dir = 'docs'
for root, dirs, files in os.walk(doc_dir):
    for fn in files:
        if not fn.endswith('.html'): continue
            
        fp = os.path.join(root, fn)
        try:
            with open(fp, 'r', encoding='utf-8') as f:
                html = f.read()
                
            # append cache buster to styles.css
            html = re.sub(r'(href=.*styles\.css)(\?v=\d+)?(\"|\')', r'\1?v=9\3', html)
            
            with open(fp, 'w', encoding='utf-8') as f:
                f.write(html)
        except Exception as e:
            pass
