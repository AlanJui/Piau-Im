import json
import sys

import xlwings as xw


def load_from_text_file(input_path):
    with open(input_path, 'r', encoding='utf-8') as file:
        # 讀取檔頭部分（JSON 格式）
        env_data = json.loads(file.readline().strip())

        # 讀取檔體部分
        content = file.readlines()
        content = [line.strip() for line in content if line.strip()]

        return env_data, content

def fill_excel_sheets(wb, env_data, content):
    # 回填到 env 工作表
    env_sheet = wb.sheets['env']
    for key, value in env_data.items():
        env_sheet.range(key).value = value

    # 回填到 漢字注音 工作表
    han_ji_sheet = wb.sheets['漢字注音']
    start_row = 5
    start_col = 4
    current_row = start_row
    current_col = start_col

    for line in content:
        if line == "φ":  # 文章結束
            break
        elif line == "\n":  # 換行
            current_row += 4
            current_col = start_col
        else:
            han_ji, tai_gi_im_piau = line.split(" | ")
            # 寫入漢字
            han_ji_sheet.range((current_row, current_col)).value = han_ji
            # 寫入台語音標
            han_ji_sheet.range((current_row - 1, current_col)).value = tai_gi_im_piau if tai_gi_im_piau else None
            current_col += 1

# 使用範例
def main():

    # 利用 sys.argv 取得命令列參數，第一個參數應為 HTML 檔案路徑
    if len(sys.argv) < 2:
        print("請在命令列提供 HTML 檔案路徑參數。")
        sys.exit(1)

    # 讀取純文字檔
    # input_path = 'output.txt'
    input_path = sys.argv[1]
    env_data, content = load_from_text_file(input_path)

    try:
        wb = xw.apps.active.books.active
    except Exception as e:
        print("找不到作用中的 Excel 工作簿！", e)
        sys.exit(1)

    # 回填到 Excel
    fill_excel_sheets(wb, env_data, content)

    # 保存並關閉 Excel 檔案
    wb.save()
    wb.close()

if __name__ == "__main__":
    main()