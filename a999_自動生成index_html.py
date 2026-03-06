"""
a999_自動生成index_html.py v0.2

為 docs 目錄下的 HTML 檔案自動生成 index.html，並根據檔案的修改時間排序，距今天越近的檔案擺在越前面。

變更紀錄：
- v0.2 (2024-06-30): 改正了排序邏輯，現在是根據檔案的修改時間（mtime）來排序，而不是檔案名稱。這樣可以確保最新修改的文章會顯示在最前面。
"""

import os
import re

# 要生成超連結的目錄
docs_directory = "docs"

ignore_dir_list = [
    "_archived",
    "_test",
    "金鋼經",
]

ignore_doc_list = [
    "index.html",
    "index_bak.html",
    "_template.html",
    "output_from_excel.html",
]


def create_file_list(directory, extension):
    # 建立檔案清單
    file_list = []

    # 遍歷目錄下的檔案
    for filename in os.listdir(directory):
        # 排除指定的檔案
        if filename not in ignore_doc_list:
            if filename.endswith(extension):
                file_list.append(filename)
    return file_list


# 生成超連結的檔案名稱
index_file = os.path.join(docs_directory, "index.html")

# HTML 模板檔案名稱
template_file = os.path.join(docs_directory, "_template.html")

# 開始 index.html 檔案
with open(index_file, "w", encoding="utf-8") as f:  # 指定 UTF-8 編碼
    # 讀取 HTML 模板內容
    with open(template_file, "r", encoding="utf-8") as template:
        template_content = template.read()

        # 文章名稱與注音方式的字典
        articles = {}

        # 遍歷目錄下的檔案
        for root, dirs, files in os.walk(docs_directory):
            # 移除要忽略的資料夾
            dirs[:] = [d for d in dirs if d not in ignore_dir_list]

            for filename in files:
                # 排除不要處理的檔案 (非html或指定的排除檔案)
                if not filename.endswith(".html") or filename in ignore_doc_list:
                    continue

                full_path = os.path.join(root, filename)
                mtime = os.path.getmtime(full_path)

                # 取得相對路徑（用於建立超連結）
                rel_dir = os.path.relpath(root, docs_directory)
                if rel_dir == ".":
                    relative_path = filename
                else:
                    relative_path = os.path.join(rel_dir, filename).replace("\\", "/")

                # 取得檔案名稱 (不包含副檔名)
                article_and_phonetic = os.path.splitext(filename)[0]

                # 嘗試分割文章名稱與注音方式
                if "_" in article_and_phonetic:
                    parts = article_and_phonetic.split("_")
                    phonetic_method = parts[-1]
                    article = "_".join(parts[:-1])
                else:
                    # 處理如『《歸去來辭》【河洛白話音】十五音.html』這類沒有底線的命名
                    match = re.search(r"^(.*【.*?】)(.*)$", article_and_phonetic)
                    if match and match.group(2).strip():
                        article = match.group(1)
                        phonetic_method = match.group(2)
                    else:
                        article = article_and_phonetic
                        phonetic_method = "開啟"

                # 如果有子目錄，把子目錄名稱加到文章分類前面
                if rel_dir != ".":
                    article = f"[{rel_dir}] " + article

                # 避免抓到莫名其妙名稱為空的檔案
                if not article.strip():
                    continue

                # 如果文章名稱不在 articles 字典中，則添加
                if article not in articles:
                    articles[article] = {"methods": [], "mtime": 0}

                # 如果注音方式不在文章名稱對應的列表中，則添加
                has_added = any(
                    item[0] == phonetic_method for item in articles[article]["methods"]
                )
                if not has_added:
                    articles[article]["methods"].append(
                        (phonetic_method, relative_path)
                    )
                    # 更新這篇文章的最新時間
                    if mtime > articles[article]["mtime"]:
                        articles[article]["mtime"] = mtime

        # 生成文章清單的 HTML 內容
        articles_html = '<div class="card-grid">\n'

        # 針對文章名稱的最新修改時間排序，距今天越近（mtime越大）擺最上面
        sorted_articles = sorted(
            articles.items(), key=lambda x: x[1]["mtime"], reverse=True
        )

        for article, data in sorted_articles:
            phonetic_methods = data["methods"]
            articles_html += '  <div class="card">\n'
            articles_html += f'    <h2 class="card-title">{article}</h2>\n'
            articles_html += '    <div class="card-links">\n'
            for phonetic_method, target_url in phonetic_methods:
                articles_html += f'      <a href="{target_url}" class="badge" target="_blank" title="{article} - {phonetic_method}">{phonetic_method}</a>\n'
            articles_html += "    </div>\n"
            articles_html += "  </div>\n"
        articles_html += "</div>\n"

        # 將文章清單的 HTML 內容插入模板中
        template_content = template_content.replace(
            "{articles_placeholder}", articles_html
        )

        # 將修改後的模板內容寫入 index.html
        f.write(template_content)
