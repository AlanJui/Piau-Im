"""
自動製作打字練習表
使用 xlwings 從【漢字注音】工作表製作【打字練習表】
"""

import argparse
import re

import xlwings as xw


def get_tone_key_mapping(tone_map_type='tlpa'):
    """
    取得聲調與按鍵的對照表
    """
    # 台語音標之【拚音輸入法】：【聲調調號】與【鍵盤按鍵】對照表
    # tlpa_phing_im_tiau_key_map = {
    #     '1': ';',   # 陰平
    #     '7': '-',   # 陽去
    #     '3': '_',   # 陰去
    #     '2': '\\',  # 陰上
    #     '5': '/',   # 陽平
    #     '4': '[',   # 陰入
    #     '8': ']'    # 陽入
    # }

    # 閩拚音標之【拚音輸入法】：【聲調調號】與【鍵盤按鍵】對照表
    # bp_phing_im_tiau_key_map = {
    #     '1': ';',   # 陰平
    #     '5': '_',   # 陰去
    #     '6': '-',   # 陽去
    #     '3': '\\',  # 陰上
    #     '2': '/',   # 陽平
    #     '7': '[',   # 陰入
    #     '8': ']'    # 陽入
    # }

    # 閩拚音標之【注音輸入法】：【調符】與【調號】對照表
    # bp_zu_im_hu_tiau_map2 = {
    #     ' ̄': '1'    # 陰平 ==> Unicode: U+0304
    #     ' ̂': '6',   # 陽去 ==> Unicode: U+0302
    #     'ˋ': '5',  # 陰去 ==> Unicode: U+0300
    #     'ˇ': '3',  # 上声 ==> Unicode: U+030C
    #     'ˊ': '2',  # 陽平 ==> Unicode: U+0301
    #     '˫': '7',   # 陽入
    #     '˙': ' ',   # 輕聲（空白鍵）
    # }

    # 閩拚音標之【注音輸入法】：【調符】與【調號】對照表
    bp_zu_im_hu_tiau_map = {
        '˫': '6',   # 陽去
        '˪': '5',   # 陰去
        'ˋ': '3',  # 陰上
        'ˊ': '2',  # 陽平
        '˙': '8',   # 陽入
    }

    # 閩拚音標之【注音輸入法】：【調號】與【按鍵】對照表
    bp_zu_im_tiau_key_map = {
        '1': ':',   # 陰平
        '6': '5',   # 陽去
        '5': '3',   # 陰去
        '3': '4',   # 陰上
        '2': '6',   # 陽平
        '7': '[',   # 陰入
        '8': ']',   # 陽入
        '0': '7',   # 輕聲 ⁰
    }

    # 台語音標之【注音輸入法】：【調符】與【調號】對照表
    tlpa_zu_im_hu_tiau_map = {
        '˫': '6',   # 陽去
        '˪': '5',   # 陰去
        'ˋ': '3',  # 陰上
        'ˊ': '2',  # 陽平
        '˙': '8',   # 陽入
    }

    # 台語音標之【拚音輸入法】：【調號】與【按鍵】對照表
    tlpa_zu_im_tiau_key_map = {
        '1': ':',   # 陰平
        '7': '5',   # 陽去
        '3': '3',   # 陰去
        '2': '4',   # 陰上
        '5': '6',   # 陽平
        '4': '[',   # 陰入
        '8': ']',   # 陽入
        '0': '7',   # 輕聲 ⁰
    }

    if tone_map_type == 'bp':
        return bp_zu_im_hu_tiau_map, bp_zu_im_tiau_key_map
    else:
        return tlpa_zu_im_hu_tiau_map, tlpa_zu_im_tiau_key_map


def is_punctuation(char):
    """
    判斷是否為標點符號
    """
    if char is None or str(char).strip() == '':
        return False

    # 常見的中文標點符號
    chinese_punctuation = '，。！？；：「」『』（）【】《》〈〉、—…～'
    # 常見的英文標點符號
    english_punctuation = ',.!?;:"()[]{}/<>-_=+*&^%$#@`~|\\\'\"'

    return str(char) in chinese_punctuation or str(char) in english_punctuation


def is_line_break(char):
    """
    判斷是否為換行控制字元
    """
    if char is None:
        return False

    return char == '\n' or str(char).strip() == '' or char == 10


def decompose_pronunciation(pronunciation, tone_map_type='tlpa'):
    """
    將注音符號或羅馬拼音分解成個別字元

    Args:
        pronunciation (str): 注音符號或羅馬拼音

    Returns:
        list: 分解後的字元列表
    """

    # 【調符】與【調號】對映：hu_tiau_map
    # 【調號】與【按鍵】對映：tiau_key_map
    hu_tiau_map, tiau_key_map = get_tone_key_mapping(tone_map_type)

    # 檢查【音節】是否【帶有調號】（含數字）
    if re.search(r'\d', pronunciation):
        # 帶調號音節處理
        # 找出調號（聲調）
        tone_match = re.search(r'(\d+)', pronunciation)
        if tone_match:
            tone = tone_match.group(1)
            # 移除聲調數字，取得拼音部分
            letters = pronunciation[:tone_match.start()]

            # 特殊處理：如果是入聲調（4調、8調），需要調整最後字母
            if tone in ['4', '8']:
                # 將 ng 結尾改為 k 結尾（入聲調的特殊處理）
                if letters.endswith('ng'):
                    letters = letters[:-2] + 'k'
                # 將 n 結尾改為 t 結尾
                elif letters.endswith('n') and not letters.endswith('ng'):
                    letters = letters[:-1] + 't'
                # 將 m 結尾改為 p 結尾
                elif letters.endswith('m'):
                    letters = letters[:-1] + 'p'

            # 轉換聲調為按鍵
            tone_key = hu_tiau_map.get(tone, tone)

            # 分解拼音字母並加上聲調按鍵
            result = list(letters) + [tone_key]
        else:
            result = list(pronunciation)
    else:
        # 方音符號處理
        chars = list(pronunciation)
        result = []
        okay = False
        counter = len(chars)
        index = 0

        for i, char in enumerate(chars):
            # hu_tiau_map = ['˪', '˫', 'ˋ', 'ˊ', '˙']
            if char in hu_tiau_map:
                # 是聲調符號，轉換為按鍵
                tiau_ho = hu_tiau_map[char]
                result.append(tiau_key_map[tiau_ho])
                okay = True
            else:
                result.append(char)
            index += 1

        # 如果沒有聲調符號，則可能是：【陰平調】或【陰入調】
        if not okay:
            if index == counter:
                if chars[counter - 1] in ['ㆴ', 'ㆵ', 'ㆻ', 'ㆷ']:
                    # 若最後一個字元，是【入聲韻尾】，則視為【陰入調】
                    if tone_map_type == 'tlpa':
                        result.append(tiau_key_map[str(4)])
                    else:
                        result.append(tiau_key_map[str(7)])
                else:
                    # 若最後一個字元，亦不是【入聲韻尾】，則視為【陰平調】
                    result.append(tiau_key_map[str(1)])

        # # 如果沒有聲調符號，假設是陰平聲（空白鍵）
        # if len(result) == len(chars) and not any(c in bopomofo_tone_map for c in chars):
        #     result.append(' ')

    return result


def _has_meaningful_data(values):
    """Return True if any cell in the provided values contains non-blank data."""
    def _is_blank(cell):
        if cell is None:
            return True
        if isinstance(cell, str) and cell.strip() == '':
            return True
        return False

    if values is None:
        return False

    if not isinstance(values, list):
        return not _is_blank(values)

    for row in values:
        cells = row if isinstance(row, list) else [row]
        for cell in cells:
            if not _is_blank(cell):
                return True
    return False


def calculate_total_rows(sheet, start_col='D', end_col='R', base_row=3, rows_per_group=4):
    """Compute how many row groups exist based on the described worksheet layout."""
    total_rows = 0
    current_base = base_row

    while True:
        han_row = current_base + 2
        pronunciation_row = current_base + 3
        target_range = sheet.range(f'{start_col}{han_row}:{end_col}{pronunciation_row}')
        values = target_range.value

        if not _has_meaningful_data(values):
            break

        total_rows += 1
        current_base += rows_per_group

    return total_rows


def create_typing_practice_sheet(tone_map_type='roman'):
    """
    主函數：製作打字練習表

    Args:
        tone_map_type (str): 聲調對照類型，'roman' 或 'bp'，預設為 'roman'
    """
    try:
        # 取得作用中的活頁簿
        wb = xw.books.active
        print("已取得作用中活頁簿")

        # 取得【漢字注音】工作表
        han_ji_sheet = wb.sheets['漢字注音']
        print("已取得【漢字注音】工作表")

        # 取得或建立【打字練習表】工作表
        try:
            typing_sheet = wb.sheets['打字練習表']
            print("已找到【打字練習表】工作表")
        except Exception:
            typing_sheet = wb.sheets.add('打字練習表')
            print("已建立新的【打字練習表】工作表")

        # 清空打字練習表的內容（從第4行開始）
        # typing_sheet.range('B4:M2000').clear()
        typing_sheet.range('B4:M2000').clear_contents()

        # 不設定 E3:M3 的標題，按需求不透過程式置入

        # 開始處理資料
        current_row = 4  # 從第4行開始填入資料

        print("開始處理漢字注音資料...")

        # 處理所有列的資料
        # 第1列：{D3:R6} - 第3格D5, 第4格D6
        # 第2列：{D7:R10} - 第3格D9, 第4格D10
        # 第3列：{D11:R14} - 第3格D13, 第4格D14
        # 第4列：{D15:R18} - 第3格D17, 第4格D18
        # 第5列：{D19:R22} - 第3格D21, 第4格D22
        # ... 以此類推
        # 根據【漢字注音】工作表，計算【總列數】
        total_rows = calculate_total_rows(han_ji_sheet)
        if total_rows == 0:
            print("【漢字注音】工作表沒有可用資料，結束處理")
            return

        # 計算各列的起始行號：3, 7, 11, 15, 19, 23
        row_starts = [3 + i * 4 for i in range(total_rows)]  # [3, 7, 11, 15, 19, 23]

        for row_group_index, base_row in enumerate(row_starts):
            # print(f"\n處理第 {row_group_index + 1} 列群組，基準行: {base_row}")
            print(f"\n----------------------------------------------------------")
            print(f"第 {row_group_index + 1} 列（漢字行: {base_row+2}）")
            print(f"----------------------------------------------------------")

            # 每列處理 D到R欄 (第4到第18欄)
            for col_index in range(4, 19):  # D(4) 到 R(18)
                try:
                    col_letter = chr(64 + col_index)

                    # 計算漢字和標音的實際行號
                    han_ji_row = base_row + 2    # 第3格
                    pronunciation_row = base_row + 3  # 第4格
                    tai_gi_row = base_row + 1  # 第2格（目前未使用）

                    # 取得當前單元格的資料
                    han_ji = han_ji_sheet.range(f'{col_letter}{han_ji_row}').value
                    pronunciation = han_ji_sheet.range(f'{col_letter}{pronunciation_row}').value
                    tai_gi_piau_im = han_ji_sheet.range(f'{col_letter}{tai_gi_row}').value

                    # print(f"處理 {col_letter}{han_zi_row}/{col_letter}{pronunciation_row}: 漢字={repr(han_zi)}, 標音={repr(pronunciation)}")
                    print(f"{col_index-3}.【{col_letter}{han_ji_row}】: 漢字={repr(han_ji)}, 標音={repr(pronunciation)}")
                    print("\n")

                    # 檢查是否遇到終結符號
                    if han_ji == 'φ':
                        print("    ==> 遇到終結符號，停止處理")
                        break

                    # 檢查是否為換行控制字元
                    if is_line_break(han_ji):
                        print(f"    ==> 欄位 {col_letter} 遇到換行控制字元，在打字練習表留空白行")
                        # 留空白行（不填任何資料）
                        current_row += 1
                        continue

                    # 檢查是否為標點符號
                    if is_punctuation(han_ji):
                        # print(f"    ==> 欄位 {col_letter} 是標點符號: {han_zi}")
                        # 標點符號只填入B欄，C欄及後續欄位留空
                        typing_sheet.range(f'B{current_row}').api.Value2 = str(han_ji)
                        current_row += 1
                        continue

                    # 檢查資料是否有效
                    if han_ji is None or pronunciation is None:
                        print(f"    ==> 欄位 {col_letter} 資料為空，跳過")
                        continue

                    # 處理正常的漢字和標音
                    # print(f"漢字: {han_zi} - {pronunciation}")

                    # 填入純文字資料（不改變格式）
                    typing_sheet.range(f'B{current_row}').api.Value2 = str(han_ji)
                    typing_sheet.range(f'C{current_row}').api.Value2 = str(pronunciation)

                    # 分解標音符號
                    # tone_map_type = 'tfs'
                    decomposed = decompose_pronunciation(str(pronunciation), tone_map_type)
                    print(f"    ==> 鍵盤按鍵: {decomposed}\n")

                    # 將分解後的字元填入 E~M 欄（純文字）
                    for i, char in enumerate(decomposed):
                        if i < 9:  # 最多填入9個字元（E~M欄）
                            col_letter_target = chr(69 + i)  # E=69, F=70, ...
                            typing_sheet.range(f'{col_letter_target}{current_row}').api.Value2 = char

                    current_row += 1

                except Exception as e:
                    print(f"處理欄位 {col_letter} 時發生錯誤: {e}")
                    continue

            # 如果遇到終結符號，跳出外層循環
            if han_ji == 'φ':
                break

        # 使用【打字練習表（模版）】或【打字練習表 (模版)】工作表來統一格式
        template_sheet_names = ['打字練習表（模版）', '打字練習表 (模版)']
        template_sheet = None

        for template_name in template_sheet_names:
            try:
                template_sheet = wb.sheets[template_name]
                print(f"找到【{template_name}】工作表，開始統一格式")
                break
            except Exception:
                continue

        if template_sheet:
            # 取得模版的格式
            template_range = template_sheet.range('B4:M4')
            template_range.api.Copy()

            # 應用到打字練習表的所有資料列
            data_rows = current_row - 4  # 計算實際資料列數
            if data_rows > 0:
                target_range = typing_sheet.range(f'B4:M{3 + data_rows}')
                target_range.api.PasteSpecial(-4122)  # xlPasteFormats
                print(f"已將模版格式應用到 {data_rows} 列資料")

            # 清除剪貼板
            wb.app.api.CutCopyMode = False
        else:
            print("警告：找不到【打字練習表（模版）】或【打字練習表 (模版)】工作表")
            print("將使用預設格式")

        # 設定欄寬以便觀看
        typing_sheet.range('B:M').column_width = 10

        # 啟動打字練習表工作表
        typing_sheet.activate()

        print(f"打字練習表製作完成！共處理了 {current_row - 4} 個漢字")
        print("已切換到【打字練習表】工作表")

    except Exception as e:
        print(f"發生錯誤: {e}")
        return False

    return True


def main():
    """
    主程式入口點
    """
    # 設定命令列參數解析
    parser = argparse.ArgumentParser(description='自動製作打字練習表')
    parser.add_argument(
        'tone_map_type',
        nargs='?',
        default='tlpa',
        choices=['tlpa', 'bp'],
        help='聲調對照類型：tlpa (台語拼音，預設) 或 bp (閩拚)'
    )

    args = parser.parse_args()

    print("=== 自動製作打字練習表 ===")
    print(f"聲調對照類型: {args.tone_map_type}")
    print("請確保:")
    print("1. Excel 已開啟並有作用中的活頁簿")
    print("2. 活頁簿中包含【漢字注音】工作表")
    print("3. 漢字注音工作表的資料格式正確")
    print()

    success = create_typing_practice_sheet(args.tone_map_type)

    if success:
        print("\n✓ 打字練習表製作成功！")
    else:
        print("\n✗ 打字練習表製作失敗！")


if __name__ == "__main__":
    main()