import os
import sys

# 從命令列參數取得目錄檔案清單名稱、目錄路徑、副檔名和除外清單檔案
list_file_name = sys.argv[1]
directory_path = sys.argv[2]
file_extension = sys.argv[3]
exclude_file = sys.argv[4] if len(sys.argv) > 4 else None

# 獲取目錄中的所有檔案
all_files = os.listdir(directory_path)

# 過濾出與指定副檔名相符且不以 ~$ 開頭的檔案
filtered_files = [file for file in all_files if file.endswith(file_extension) and not file.startswith('~$')]

# 如果提供了除外清單檔案，則讀取該檔案並從檔案清單中排除在除外清單中的檔案
if exclude_file:
    with open(exclude_file, 'r', encoding='utf-8') as file:
        exclude_files = file.read().splitlines()
    filtered_files = [file for file in filtered_files if file not in exclude_files]

# 將最終的檔案清單寫入到目錄檔案清單檔案中
with open(list_file_name, 'w', encoding='utf-8') as file:
    for file_name in filtered_files:
        file.write(file_name + '\n')