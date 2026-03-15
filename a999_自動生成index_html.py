"""
a999_自動生成index_html.py v0.2.1

為 docs 目錄下的 HTML 檔案自動生成 index.html。
修正：加入顯式排序邏輯，確保 Local 與 GitHub Actions 產生的順序一致。
"""

import os
import re

docs_directory = "docs"
ignore_dir_list = ["_archived", "_test", "金鋼經", "__bak"]
ignore_doc_list = ["index.html", "index_bak.html", "_template.html", "output_from_excel.html"]

index_file = os.path.join(docs_directory, "index.html")
template_file = os.path.join(docs_directory, "_template.html")

# 1. 收集所有檔案資訊
all_files_info = []
for root, dirs, files in os.walk(docs_directory):
    dirs[:] = [d for d in dirs if d not in ignore_dir_list]
    for filename in files:
        if not filename.endswith(".html") or filename in ignore_doc_list:
            continue
        
        full_path = os.path.join(root, filename)
        mtime = os.path.getmtime(full_path)
        rel_dir = os.path.relpath(root, docs_directory)
        relative_path = filename if rel_dir == "." else os.path.join(rel_dir, filename).replace("\\", "/")
        
        all_files_info.append({
            "filename": filename,
            "root": root,
            "mtime": mtime,
            "relative_path": relative_path,
            "rel_dir": rel_dir
        })

# 2. 執行排序 (先按 mtime 倒序，同時間則按檔名排序)
# 注意：GitHub Actions 環境下 mtime 可能相同，所以檔名排序是關鍵
all_files_info.sort(key=lambda x: (-x["mtime"], x["filename"]))

# 3. 處理文章字典
articles = {}
for info in all_files_info:
    filename = info["filename"]
    article_and_phonetic = os.path.splitext(filename)[0]
    
    # 解析文章標題與標音方式
    if "_" in article_and_phonetic:
        parts = article_and_phonetic.split("_")
        phonetic_method = parts[-1]
        article = "_".join(parts[:-1])
    else:
        match = re.search(r"^(.*【.*?】)(.*)$", article_and_phonetic)
        if match and match.group(2).strip():
            article = match.group(1)
            phonetic_method = match.group(2)
        else:
            article = article_and_phonetic
            phonetic_method = "開啟"

    # 清理 None
    if "None" in phonetic_method:
        phonetic_method = phonetic_method.replace("None＋", "").replace("＋None", "").replace("None", "").strip()
        if not phonetic_method: phonetic_method = "開啟"

    if info["rel_dir"] != ".":
        article = f"[{info['rel_dir']}] {article}"

    if article not in articles:
        articles[article] = []
    articles[article].append({"method": phonetic_method, "path": info["relative_path"]})

# 4. 生成 HTML 內容
cards_html = ""
# 這裡對字典的 Key 也進行一次排序，確保文章區塊順序一致
for article in sorted(articles.keys()):
    links = "".join([f'<a href="{p["path"]}" class="badge">{p["method"]}</a>' for p in articles[article]])
    cards_html += f"""
    <div class="card">
        <h2 class="card-title">{article}</h2>
        <div class="card-links">
            {links}
        </div>
    </div>
    """

# 5. 寫入檔案
with open(template_file, "r", encoding="utf-8") as t:
    content = t.read().replace("{placeholder}", cards_html)
    with open(index_file, "w", encoding="utf-8") as f:
        f.write(content)

print(f"成功生成索引：{index_file}，包含 {len(articles)} 篇文章。")
