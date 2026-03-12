
import os
import sqlite3
import re
from pathlib import Path

# Load database for reverse lookup
DB_PATH = 'Ho_Lok_Ue.db'
DOCS_DIR = 'docs'

def get_db_mapping():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT 漢字, 台羅音標 FROM 漢字庫")
    mapping = {}
    for hanji, tlpa in cursor.fetchall():
        if hanji not in mapping:
            mapping[hanji] = []
        mapping[hanji].append(tlpa)
    conn.close()
    return mapping

def patch_file(file_path, db_mapping):
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()

    # 1. Add JS/CSS if missing
    if 'phonetic_switcher.js' not in content:
        content = content.replace('</head>', '    <script type="text/javascript" src="assets/javascripts/phonetic_switcher.js"></script>\n</head>')

    # 2. Inject data-tlpa into <ruby> tags
    # This is a bit complex as we need to find the character and its current pronunciation
    # and guess the TLPA if it's not there.
    
    def ruby_replacer(match):
        full_ruby = match.group(0)
        if 'data-tlpa' in full_ruby:
            return full_ruby
            
        # Extract Hanji and current pronunciation
        hanji_match = re.search(r'<ruby>\s*([^<>\s\n]+)', full_ruby)
        rt_match = re.search(r'<rt>([^<>]+)</rt>', full_ruby)
        
        if not hanji_match: return full_ruby
        hanji = hanji_match.group(1).strip()
        
        # Try to guess TLPA from DB
        tlpa = ""
        if hanji in db_mapping:
            # If multiple pronunciations, we might be wrong, but better than nothing
            # Or we could try to match the current RT if it's already TL-like
            tlpas = db_mapping[hanji]
            if rt_match:
                curr_rt = rt_match.group(1).strip()
                # If current RT is one of the TLPAs, use it
                if curr_rt in tlpas:
                    tlpa = curr_rt
                else:
                    tlpa = tlpas[0]
            else:
                tlpa = tlpas[0]
        
        if tlpa:
            return full_ruby.replace('<ruby>', f'<ruby data-tlpa="{tlpa}">')
        return full_ruby

    new_content = re.sub(r'<ruby>.*?</ruby>', ruby_replacer, content, flags=re.DOTALL)
    
    if new_content != content:
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(new_content)
        return True
    return False

def main():
    print("Loading database...")
    db_mapping = get_db_mapping()
    
    files = [f for f in os.listdir(DOCS_DIR) if f.endswith('.html') and f != 'index.html']
    print(f"Found {len(files)} files to patch.")
    
    patched_count = 0
    for file in files:
        file_path = os.path.join(DOCS_DIR, file)
        if patch_file(file_path, db_mapping):
            patched_count += 1
            print(f"Patched: {file}")
            
    print(f"Done! Patched {patched_count} files.")

if __name__ == "__main__":
    main()
