"""
a930_自網頁匯入漢字拼音.py v0.0.2

功能：
    讀取指定的 HTML 檔案，解析其中的 <ruby> 標籤結構，
    將 漢字 (<rb>) 與 標音 (<rt>) 提取出來，
    並依序填入 Excel 工作表。

使用方式：
    python a930_自網頁匯入漢字拼音.py [html_file_path]

需求套件：
    pip install beautifulsoup4 xlwings
"""

import os
import sqlite3
import sys
from pathlib import Path

import xlwings as xw
from bs4 import BeautifulSoup, Comment, NavigableString, Tag

from mod_十五音 import huan_ciat_ca_piau_im, tiau_ho_tng_siann_tiau


def parse_html_to_data(html_content):
    """
    解析 HTML 內容，回傳 (漢字, 標音) 的 list。

    支援結構：
    1. 一般文字 -> (文字, "")
    2. <ruby><rb>漢字</rb><rt>標音</rt>...</ruby> -> (漢字, 標音)
    """
    soup = BeautifulSoup(html_content, "html.parser")
    data_list = []

    # 避免重複處理巢狀標籤：只處理最末端的容器 (不包含其他區塊元素的元素)
    # 或者，只選擇特定的 class，如 content-box 下的 p，title-page 下的 h1, h2

    # 策略修正：
    # 1. 找到所有 div.content-box 下的 p
    # 2. 找到 div.title-page 下的 h1, h2 (及其內部的 p)

    # 簡單做法：使用 select 選擇器定位
    # 注意：h1 裡面可能有 p，這在 HTML 是不標準的，但範例中有。
    # 如果選了 h1 又選了 h1 p，會重複。
    # 我們優先選取最深層的 p。如果 h1 內有 p，則處理 p 時自然會處理到。
    # 但如果 h1 內沒有 p 只有文字和 ruby 呢？

    # 收集所有目標 tag
    target_tags = []

    # 標題區塊
    title_area = soup.find("div", class_="title-page")
    if title_area:
        # 找 h1, h2，但排除內部已經有 p 的 (由 p 自己處理)
        # 其實範例中 h1 內有 p。
        # 我們直接找 title_area 下所有的 p, h1, h2，然後過濾掉「由其他 target 包裹」的
        candidates = title_area.find_all(["h1", "h2", "p"])
        target_tags.extend(candidates)

    # 內容區塊
    content_boxes = soup.find_all("div", class_="content-box")
    for box in content_boxes:
        target_tags.extend(box.find_all("p"))

    # 去重與過濾：如果一個 tag 是另一個 tag 的後代，則只保留後代 (最深層)
    # 例如 h1 > p，我們只處理 p，避免 h1 被當作純文字處理

    final_targets = []
    # 建立 set 方便檢查
    tag_set = set(target_tags)

    for tag in target_tags:
        # 檢查 tag 是否包含其他候選 tag，如果有，則這個 tag 不需要處理 (或是只處理它直接的 text node?)
        # 更好的方式：檢查 tag 的 parent 是否也在 tag_set 中。
        # 如果 parent 在 set 中，表示 parent 會處理到這個 tag (或是 parent 的邏輯需要修改以忽略 Tag children)
        # 上個版本的邏輯是：遇 Tag 則 get_text (錯誤來源)。

        # 修正邏輯：
        # 每個 target 獨立處理。
        # 如果 target A 包含 target B，我們處理 A 時應該跳過 B 嗎？
        # 應該是：處理 A 時，只處理 A 的直接子節點 (direct children)。
        # 如果子節點是 ruby -> 提取。
        # 如果子節點是 text -> 提取。
        # 如果子節點是 p (巢狀) -> 遞迴？但我們已經把 p 加入 targets 了。
        # 所以處理 A 時，若遇到 Tag 且該 Tag 也在 targets 列表中，則跳過 (留給該 Tag 自己處理)。
        # 若 Tag 不在 targets (如 ruby, span, br)，則需處理。
        pass

    # 讓我們重新實作遍歷邏輯，不使用 target_tags 的包含檢查，而是直接走訪 DOM tree
    # 只走訪我們感興趣的 root 節點

    roots = []
    if title_area:
        roots.append(title_area)
    roots.extend(content_boxes)

    extract_data = []

    def process_element(element):
        if element is None:
            return

        # 這裡的 element 必須是一個 Tag，如果是字串，就無法遍歷 children
        # 但在遞迴過程中，如果是 Tag，它應該有 children 屬性
        # 如果 element 本身不是 Tag (例如它是 None 或其他類型)，則跳過
        if not isinstance(element, Tag):
            return

        for child in element.children:
            if isinstance(child, Comment):
                continue
            if isinstance(child, NavigableString):
                text = child.string
                if text:
                    # 濾掉換行與純空白
                    # 但要注意 HTML 中的標點符號可能會跟隨換行符，例如：
                    # <p>
                    #   《
                    #   <ruby>...
                    # </p>
                    # 此時 text 可能是 "\n      《\n      "
                    # 我們只取 "《"
                    stripped_text = text.replace("\n", "").replace("\r", "").strip()
                    for char in stripped_text:
                        extract_data.append((char, ""))

            elif isinstance(child, Tag):
                if child.name == "ruby":
                    # 處理 Ruby
                    rb = child.find("rb")
                    rt = child.find("rt")
                    rb_text = rb.get_text(strip=True) if rb else ""

                    # 若無 rb，嘗試取 child 文字 (扣除 rt/rp)
                    if not rb_text:
                        rt_text_content = rt.get_text(strip=True) if rt else ""
                        full_text = child.get_text(strip=True)
                        # 簡單移除 rt 文字 (不精確但可用)
                        rb_text = full_text.replace(rt_text_content, "")

                        # 去除 rp
                        for rp in child.find_all("rp"):
                            rb_text = rb_text.replace(rp.get_text(strip=True), "")

                    rt_text = rt.get_text(strip=True) if rt else ""

                    if rb_text:
                        extract_data.append((rb_text, rt_text))

                elif child.name == "br":
                    continue

                elif child.name in ["script", "style", "meta", "link"]:
                    continue

                else:
                    # 其他標籤 (如 p, h1, h2, span, div 等)，視為容器繼續遞迴
                    # 關鍵修正：這裡需要遞迴呼叫 process_element(child)
                    process_element(child)

    for root in roots:
        process_element(root)

    return extract_data


def process_phonetic(phonetic_str, cursor):
    """
    將「堅五曾」格式轉換為 (台語音標, 聲, 韻, 調)
    如: '堅五曾' -> ('zian5', 'z', 'ian', '5')
    """
    if not phonetic_str or len(phonetic_str) != 3:
        return ("", "", "", "")

    yun = phonetic_str[0]  # 堅
    tone_code = phonetic_str[1]  # 五
    initial = phonetic_str[2]  # 曾

    siann_tiau = tiau_ho_tng_siann_tiau(tone_code)
    if not siann_tiau:
        # 如果調號轉換失敗，回傳空
        return ("", "", "", "")

    # 查詢資料庫
    # huan_ciat_ca_piau_im(cursor, 字韻, 聲調, 切音)
    results = huan_ciat_ca_piau_im(cursor, yun, siann_tiau, initial)

    if results:
        res = results[0]
        # 取得需要的欄位
        taigi = res.get("漢字標音", "")
        siann = res.get("聲", "")
        yun_val = res.get("韻", "")
        tiau = res.get("調", "")
        return (taigi, siann, yun_val, str(tiau))

    return ("", "", "", "")


def import_to_excel(data, excel_file=None):
    """
    將 data [(漢字, 標音), ...] 寫入 Excel
    """
    wb = None

    if excel_file and os.path.exists(excel_file):
        wb = xw.Book(excel_file)
    else:
        # 如果沒有指定檔案或檔案不存在，嘗試連接當前活動的 Excel，或開新檔
        try:
            wb = xw.books.active
        except:
            wb = xw.Book()

    # 連接資料庫
    db_path = "雅俗通十五音字典.db"
    conn = None
    cursor = None
    if os.path.exists(db_path):
        conn = sqlite3.connect(db_path)
        # 設定 row_factory 為 sqlite3.Row 或直接在 huan_ciat_ca_piau_im 处理
        # mod_十五音 的 huan_ciat_ca_piau_im 內部已經將結果轉為 dict
        cursor = conn.cursor()
    else:
        print(f"警告：找不到資料庫 {db_path}，無法進行台語音標轉換。")

    # 擴展資料：加入台語音標欄位
    extended_data = []

    for row in data:
        han_ji, phonetic = row

        # 預設空字串
        taigi = ""
        siann_val = ""
        yun_val = ""
        tiau_val = ""

        if cursor:
            # 使用 process_phonetic 函式處理
            # 注意：phonetic 可能為空或不符合格式，process_phonetic 內部有檢查
            taigi, siann_val, yun_val, tiau_val = process_phonetic(phonetic, cursor)

        extended_data.append((han_ji, phonetic, taigi, siann_val, yun_val, tiau_val))

    if conn:
        conn.close()

    # 建立或選取工作表
    sheet_name = "網頁匯入"
    if sheet_name in [s.name for s in wb.sheets]:
        sheet = wb.sheets[sheet_name]
        sheet.clear()  # 清除舊資料
    else:
        sheet = wb.sheets.add(sheet_name)

    # 寫入標頭
    sheet.range("A1").value = [
        "漢字",
        "漢字標音",
        "台語音標",
        "台語音標之聲",
        "台語音標之韻",
        "台語音標之調",
    ]

    # 寫入資料
    if extended_data:
        sheet.range("A2").value = extended_data

    # 自動調整欄寬
    sheet.autofit()

    return wb


def main():
    if len(sys.argv) > 1:
        html_path = sys.argv[1]
    else:
        # 預設路徑 (測試用)
        html_path = r"c:\work\Piau-Im\docs\《前赤壁賦》.html"

    if not os.path.exists(html_path):
        print(f"錯誤：找不到檔案 {html_path}")
        return

    print(f"正在讀取並解析：{html_path} ...")

    try:
        with open(html_path, "r", encoding="utf-8") as f:
            content = f.read()

        data = parse_html_to_data(content)

        print(f"解析完成，共 {len(data)} 筆資料。")
        print("正在寫入 Excel ...")

        wb = import_to_excel(data)

        print(f"匯入成功！請查看 Excel 工作表 '網頁匯入'。")

    except Exception as e:
        print(f"發生錯誤：{e}")
        import traceback

        traceback.print_exc()


if __name__ == "__main__":
    main()
