"""
a820_匯出製成中州韻字典檔【台語注音二式】.py v0.1.0

【功能說明】：
用於製成【台語注音二式】字典檔，專供「台語注音二式」輸入方案使用，其羅馬拼音系統採用
「台語注音二式」。

製成作業分兩個次作業流程：
1. 製成【台語注音二式】字典主檔次作業：由 a821_匯出製成【台語注音二式】字典主檔.py 負責完成；
2. 製成【台語注音二式】字典子檔次作業：由 a822_台羅拼音字典檔轉台語注音二式.py 負責完成。

使用此程式之目的，在於避免 a821, a822 因各自獨立執行，結果【遺忘執行】某次作業，
造成【台語注音二式】字典檔之不完整，以致 RIME 輸入方案之執行會有難以預期之錯誤。
由於次作業流程已有現成之程式可供呼叫，故此程式之功能主在以【批次】(Batch)方式，
令 a821 與 a822 程式能整合而連續執行。

【台語注音二式】字典主檔（.yaml）：ji_khoo_bpm2.dict.yaml。

【台語注音二式】字典主檔，包含3個子字典檔：
1. 【閩南話辭彙】：ji_khoo_su_lui_bpm2.dict.yaml
2. 【泉漳厦閩南字/辭】：ji_khoo_ban_lam_bpm2.dict.yaml
3. 【閩南話漢語正字】：ji_khoo_ziann_ji_bpm2.dict.yaml

若需要詳細之「台語注音二式」羅馬拼音系統說明，可參考：
C:/Users/AlanJui/work/rime-tlpa/docs/090_漢字標音轉換指引.md 文件。
"""

# =========================================================================
# 載入程式所需套件/模組
# =========================================================================
import importlib
import logging
import sys

# =========================================================================
# 設定日誌
# =========================================================================
logging.basicConfig(
    filename="process_log.txt",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)

# =========================================================================
# 常數定義
# =========================================================================
EXIT_CODE_SUCCESS = 0
EXIT_CODE_FAILURE = 1

# 次作業模組名稱（對應 .py 檔名，不含副檔名）
MODULE_A821 = "a821_匯出製成【台語注音二式】字典主檔"
MODULE_A822 = "a822_台羅拼音字典檔轉台語注音二式"


# =========================================================================
# 批次執行
# =========================================================================
def run_sub_job(module_name: str, job_title: str) -> int:
    """
    載入次作業模組並呼叫其 main()。
    回傳該次作業之結束代碼；若模組無 main() 或載入失敗則回傳 FAILURE。
    """
    print("=" * 70)
    print(f"▶ 開始次作業：{job_title}")
    print(f"  模組：{module_name}")
    print("=" * 70)
    logging.info("a820 開始次作業：%s（%s）", job_title, module_name)

    try:
        module = importlib.import_module(module_name)
    except Exception as e:
        print(f"❌ 無法載入模組【{module_name}】：{e}")
        logging.error("無法載入模組 %s：%s", module_name, e, exc_info=True)
        return EXIT_CODE_FAILURE

    if not hasattr(module, "main") or not callable(module.main):
        print(f"❌ 模組【{module_name}】未提供可呼叫之 main()。")
        logging.error("模組 %s 缺少 main()", module_name)
        return EXIT_CODE_FAILURE

    try:
        result = module.main()
    except Exception as e:
        print(f"❌ 次作業執行例外【{job_title}】：{e}")
        logging.error("次作業例外 %s：%s", job_title, e, exc_info=True)
        return EXIT_CODE_FAILURE

    exit_code = int(result) if result is not None else EXIT_CODE_SUCCESS
    if exit_code == EXIT_CODE_SUCCESS:
        print(f"✅ 次作業完成：{job_title}")
        logging.info("a820 次作業完成：%s", job_title)
    else:
        print(f"❌ 次作業失敗：{job_title}（結束代碼 {exit_code}）")
        logging.error("a820 次作業失敗：%s（code=%s）", job_title, exit_code)
    return exit_code


def process() -> int:
    """
    依序執行：
      1. a821：自漢字庫匯出字典主檔 ji_khoo_bpm2.dict.yaml
      2. a822：將台羅子字典檔轉成台語注音二式子檔（*_bpm2.dict.yaml）
    任一次作業失敗即中止，避免留下不完整之字典組合。
    """
    jobs = [
        (MODULE_A821, "a821 製成【台語注音二式】字典主檔（ji_khoo_bpm2.dict.yaml）"),
        (
            MODULE_A822,
            "a822 製成【台語注音二式】字典子檔"
            "（su_lui / ban_lam / ziann_ji *_bpm2.dict.yaml）",
        ),
    ]

    for module_name, title in jobs:
        code = run_sub_job(module_name, title)
        if code != EXIT_CODE_SUCCESS:
            print("⛔ 批次作業中止：前一次作業失敗，後續步驟不予執行。")
            return code

    print("=" * 70)
    print("✅ 批次作業全部完成。【台語注音二式】字典主檔與子檔均已更新。")
    print("   主檔：ji_khoo_bpm2.dict.yaml")
    print("   子檔：ji_khoo_su_lui_bpm2.dict.yaml")
    print("         ji_khoo_ban_lam_bpm2.dict.yaml")
    print("         ji_khoo_ziann_ji_bpm2.dict.yaml")
    print("=" * 70)
    logging.info("a820 批次作業全部完成")
    return EXIT_CODE_SUCCESS


# =========================================================================
# 主程式
# =========================================================================
def main() -> int:
    print("<=========== a820 批次作業開始 ===========>")
    print("整合執行 a821（字典主檔）→ a822（字典子檔）")
    result = process()
    print("<=========== a820 批次作業結束 ===========>")
    return result


if __name__ == "__main__":
    sys.exit(main())
