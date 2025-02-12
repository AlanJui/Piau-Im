import json
import os
import sys

import xlwings as xw


def save_to_text_file(wb, output_dir_path):
    # 讀取 env 工作表的設定
    env_sheet = wb.sheets['env']
    env_data = {
        "FILE_NAME": env_sheet.range("FILE_NAME").value,
        "TITLE": env_sheet.range("TITLE").value,
        "IMAGE_URL": env_sheet.range("IMAGE_URL").value,
        "OUTPUT_PATH": env_sheet.range("OUTPUT_PATH").value,
        "章節序號": env_sheet.range("章節序號").value,
        "顯示注音輸入": env_sheet.range("顯示注音輸入").value,
        "每頁總列數": env_sheet.range("每頁總列數").value,
        "每列總字數": env_sheet.range("每列總字數").value,
        "語音類型": env_sheet.range("語音類型").value,
        "漢字庫": env_sheet.range("漢字庫").value,
        "標音方法": env_sheet.range("標音方法").value,
        "網頁格式": env_sheet.range("網頁格式").value,
        "標音方式": env_sheet.range("標音方式").value,
        "上邊標音": env_sheet.range("上邊標音").value,
        "右邊標音": env_sheet.range("右邊標音").value,
        "網頁每列字數": env_sheet.range("網頁每列字數").value
    }

    # 讀取漢字注音工作表的內容
    han_ji_sheet = wb.sheets['漢字注音']
    start_row = 5
    end_row = int(start_row + env_data["每頁總列數"] * 4)
    start_col = 4
    end_col = int(start_col + env_data["每列總字數"])

    EndOfText = False
    blank_count = 0
    content = []
    for row in range(start_row, end_row, 4):
        for col in range(start_col, end_col):
            han_ji = han_ji_sheet.range((row, col)).value
            tai_gi_im_piau = han_ji_sheet.range((row - 1, col)).value

            if han_ji == 'φ':  # 文章結束
                content.append("φ")
                break
            elif han_ji == '\n':  # 換行
                content.append("\n")
            elif han_ji is None or han_ji.strip() == '':  # 空白
                if blank_count == 2:
                    EndOfText = True
                    break
                else:
                    blank_count += 1
                continue
            else:
                content.append(f"{han_ji} | {tai_gi_im_piau if tai_gi_im_piau else ''}")

            if han_ji == '\n': han_ji = '換行'
            print(f"({row}, {col})：{han_ji} | {tai_gi_im_piau if tai_gi_im_piau else ''}")
            if EndOfText: break

    #--------------------------------------------------------------------------------------
    # 寫入檔案
    #--------------------------------------------------------------------------------------
    # 自 env 工作表取得檔案名稱
    try:
        file_name = str(wb.names['TITLE'].refers_to_range.value).strip()
    except KeyError:
        setting_sheet = wb.sheets["env"]
        file_name = str(setting_sheet.range("C4").value).strip()
    output_dir_path = wb.names['OUTPUT_PATH'].refers_to_range.value
    hue_im = wb.names['語音類型'].refers_to_range.value
    piau_im_huat = wb.names['標音方法'].refers_to_range.value
    im_piat = hue_im[:2]  # 取 hue_im 前兩個字元
    # 檢查檔案名稱是否已包含副檔名
    new_file_path = os.path.join(
        ".\\{0}".format(output_dir_path),
        f"【河洛{im_piat}注音-{piau_im_huat}】{file_name}.txt")

    with open(new_file_path, 'w', encoding='utf-8') as file:
        # 寫入檔頭部分
        file.write(json.dumps(env_data, ensure_ascii=False, indent=2) + "\n")
        # 寫入檔體部分
        file.write("\n".join(content))

if __name__ == "__main__":
    # 利用 sys.argv 取得命令列參數，第一個參數應為 HTML 檔案路徑
    # if len(sys.argv) < 2:
    #     print("請在命令列提供 HTML 檔案路徑參數。")
    #     sys.exit(1)
    # html_file_path = sys.argv[1]

    try:
        wb = xw.apps.active.books.active
    except Exception as e:
        print("找不到作用中的 Excel 工作簿！", e)
        sys.exit(1)

    save_to_text_file(wb, 'output.txt')