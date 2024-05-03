# a900_Batch_Process.py a520_製作注音網頁.py -i output xlsx
import argparse
import subprocess
import sys

# 解析命令行參數
parser = argparse.ArgumentParser()
parser.add_argument('executable', help='可在 command line 執行的執行檔名稱')
parser.add_argument('-d', dest='directory', default='output', help='輸入目錄')
parser.add_argument('-t', dest='todo_file_name', default='todo_list.txt', help='待處理清單檔案名稱')
args = parser.parse_args()

# 從命令列參數取得 Python 程式的名稱
python_program = args.executable

# 獲取當前 Python 環境的路徑
python_path = sys.executable

# 取得使用者指定之 TODO List 檔案名稱
todo_list = args.todo_file_name

# 取得使用者指定之目錄路徑
directory_path = args.directory

# 開啟 TODO List 檔案並讀取每一行
with open(todo_list, 'r', encoding='utf-8') as file:
    for line in file:
        # 移除換行符號
        input_file = line.strip()

        # 使用 subprocess 執行指定的 Python 程式，並將該行作為輸入檔案參數
        subprocess.run([python_path, python_program, '-d', directory_path, '-i', input_file])
