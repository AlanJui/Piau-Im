import subprocess
import sys

# 從命令列參數取得 Python 程式的名稱
python_program = sys.argv[1]

# 獲取當前 Python 環境的路徑
python_path = sys.executable

# 開啟 todo_list.txt 檔案並讀取每一行
with open('todo_list.txt', 'r', encoding='utf-8') as file:
    for line in file:
        # 移除換行符號
        input_file = line.strip()

        # 使用 subprocess 執行指定的 Python 程式，並將該行作為輸入檔案參數
        subprocess.run([python_path, python_program, '-i', input_file])