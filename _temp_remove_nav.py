import re

for filepath in ['docs/_template.html', 'docs/index.html']:
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            html = f.read()

        html = re.sub(r'<nav class=\"main-nav\">.*?</nav>', '', html, flags=re.DOTALL)
        
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(html)
        print(f"Removed nav from {filepath}")
    except Exception as e:
        print(f"Error on {filepath}: {e}")
