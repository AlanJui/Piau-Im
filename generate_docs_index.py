import os
import re

DOCS_DIR = 'docs'
INDEX_FILE = os.path.join(DOCS_DIR, 'index.html')
TEMPLATE_FILE = os.path.join(DOCS_DIR, '_template.html')
# Files to exclude from the index
EXCLUDE_FILES = {'index.html', '_template.html', '_test.html', '_test2.html'}

def generate_index():
    print(f"Scanning directory: {DOCS_DIR}")
    if not os.path.exists(TEMPLATE_FILE):
        print(f"Error: Template file {TEMPLATE_FILE} not found.")
        return

    html_files = []
    if os.path.exists(DOCS_DIR):
        for filename in os.listdir(DOCS_DIR):
            if filename.endswith(".html") and filename not in EXCLUDE_FILES:
                html_files.append(filename)
    else:
        print(f"Error: Directory {DOCS_DIR} not found.")
        return

    print(f"Found {len(html_files)} HTML files.")
    html_files.sort()

    articles = {} # Key: Article Name, Value: List of (Link Text, Filename)

    for filename in html_files:
        name_no_ext = os.path.splitext(filename)[0]

        article_name = ""
        link_text = name_no_ext

        # 1. Try to extract content inside 《...》
        match = re.search(r'《(.*?)》', name_no_ext)
        if match:
             # Use the full bracketed name as the group key, e.g., 《定風波》
             article_name = f"《{match.group(1)}》"
        else:
             # 2. If no brackets, try splitting by underscore for legacy grouping
             parts = name_no_ext.split('_')
             if len(parts) >= 2:
                 article_name = parts[0]
             else:
                 article_name = "其他 (Others)"

        # Prepare the link text (remove the article name if it's redundant)
        if article_name != "其他 (Others)" and link_text.startswith(article_name):
            cleaned_text = link_text[len(article_name):]
            # Remove leading underscores or whitespace
            cleaned_text = cleaned_text.lstrip('_').strip()
            if cleaned_text:
                link_text = cleaned_text
            else:
                # If nothing left (e.g. file is just "《Thing》.html"), use a default text
                link_text = "全文 (Full Text)"

        if article_name not in articles:
            articles[article_name] = []
        articles[article_name].append((link_text, filename))

    # Sort keys
    sorted_article_names = sorted(articles.keys())

    # Move "Others" to the end if present
    if "其他 (Others)" in sorted_article_names:
        sorted_article_names.remove("其他 (Others)")
        sorted_article_names.append("其他 (Others)")

    content_html = ""
    for article in sorted_article_names:
        content_html += '<div class="article-group">\n'
        content_html += f'  <h2>{article}</h2>\n'
        content_html += '  <ul class="article-links">\n'
        for text, fname in sorted(articles[article]):
             content_html += f'    <li><a href="{fname}">{text}</a></li>\n'
        content_html += '  </ul>\n'
        content_html += '</div>\n'

    # Read template
    with open(TEMPLATE_FILE, 'r', encoding='utf-8') as f:
        template = f.read()

    # Inject content
    output = template.replace('{articles_placeholder}', content_html)

    # Write index.html
    with open(INDEX_FILE, 'w', encoding='utf-8') as f:
        f.write(output)

    print(f"Successfully generated {INDEX_FILE} with {len(articles)} groups.")

if __name__ == "__main__":
    generate_index()
