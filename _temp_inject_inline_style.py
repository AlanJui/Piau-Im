import os
import re

css = '''
<style id="nav-horizontal-fix">
/* 絕對橫向排列與防跑版 */
nav.main-nav {
    display: flex !important;
    flex-direction: row !important;
    justify-content: flex-start !important;
    align-items: center !important;
    flex-wrap: nowrap !important; /* 絕對不換行 */
    overflow-x: auto !important; /* 允許橫向捲動 */
    overflow-y: hidden !important;
    width: 100% !important;
    max-width: 100vw !important;
    background: #f8f9fa !important;
    padding: 15px 20px !important;
    margin: 20px 0 !important;
    border-radius: 10px !important;
    box-sizing: border-box !important;
    -webkit-overflow-scrolling: touch !important;
}

nav.main-nav ul {
    display: flex !important;
    flex-direction: row !important;
    flex-wrap: nowrap !important; /* 絕對不換行 */
    justify-content: flex-start !important;
    align-items: center !important;
    width: max-content !important; /* 讓 ul 自適應內容長度 */
    max-width: none !important;
    padding: 0 !important;
    margin: 0 !important;
    list-style: none !important;
    gap: 15px !important;
}

nav.main-nav ul li {
    display: flex !important;
    flex-direction: row !important;
    align-items: center !important;
    margin: 0 !important;
    padding: 0 !important;
    white-space: nowrap !important; /* 單一按鈕不折行 */
    flex-shrink: 0 !important;      /* 防止按鈕被擠壓 */
}

nav.main-nav ul li a, 
nav.main-nav ul li span {
    display: inline-block !important;
    white-space: nowrap !important;
}

/* 如果是手機螢幕太小，強制消除 body 被限制寬度造成的影響 */
body {
    max-width: 100% !important;
}
</style>
'''

doc_dir = 'docs'
for root, dirs, files in os.walk(doc_dir):
    for fn in files:
        if not fn.endswith('.html'): continue
            
        fp = os.path.join(root, fn)
        try:
            with open(fp, 'r', encoding='utf-8') as f:
                html = f.read()
                
            # 移除如果有的舊的 fix style
            html = re.sub(r'<style id=\"nav-horizontal-fix\">.*?</style>', '', html, flags=re.DOTALL)
            
            # 加入緊接著 head 結束前的 inline style
            html = html.replace('</head>', css + '\n</head>')
            
            with open(fp, 'w', encoding='utf-8') as f:
                f.write(html)
        except Exception as e:
            pass

print("Injected inline styles into all HTML files directly")
