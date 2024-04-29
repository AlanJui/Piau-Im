# a900_Batch_Process.py a520_製作注音網頁.py -i output xlsx
import argparse
import glob
import os
import subprocess
import sys

# 解析命令行參數
parser = argparse.ArgumentParser()
parser.add_argument('executable', help='可在 command line 執行的執行檔名稱')
parser.add_argument('-i', dest='input_directory', help='輸入目錄')
parser.add_argument('-e', dest='extension', default='*', help='副檔名')
args = parser.parse_args()

# 收集指定子目錄下的所有指定副檔名的文件
file_paths = glob.glob(os.path.join(args.input_directory, f'*.{args.extension}'))

# 過濾掉 exculude_list 中的文件
# 要排除的文件名稱
exculude_list = ['Piau-Tsu-Im.xlsx', 'env.xlsx', 'env_osX.xlsx']
files = [os.path.basename(file_path) for file_path in file_paths if os.path.basename(file_path) not in exculude_list and not os.path.basename(file_path).startswith('~$')]

# 啟動命令行執行檔並傳入文件列表
for file in files:
    filename = os.path.basename(file)
    subprocess.run([sys.executable, args.executable, '-i', filename])
