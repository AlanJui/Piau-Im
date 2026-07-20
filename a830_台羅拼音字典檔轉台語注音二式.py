"""
a830_台羅拼音字典檔轉台語注音二式.py v0.1.0
【功能摘要】：
將 yaml 格式的台羅拼音字典檔，轉換成台語注音二式字典檔。

【註】：程式之需求及設計規格，請參考文件：Documents/PRG-a830_台羅拼音字典檔轉台語注音二式.md
"""

# =========================================================================
# 載入程式所需套件/模組
# =========================================================================
import logging
import re
import sys
from pathlib import Path

from mod_convert_TLPA_to_MPS2 import convert_TLPA_to_MPS2
from mod_標音 import convert_tl_to_tlpa

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

# 來源／標的目錄（同目錄）
RIME_DIR = Path(r"C:/Users/AlanJui/work/rime-tlpa")

# 需轉換的字典檔清單（規格文件所列）
SOURCE_DICT_FILES = [
    "ji_khoo_su_lui.dict.yaml",  # 閩南話辭彙
    "ji_khoo_ban_lam.dict.yaml",  # 泉漳厦閩南字/辭
    "ji_khoo_ziann_ji.dict.yaml",  # 閩南話漢語正字
]


# =========================================================================
# 拼音轉換
# =========================================================================
def convert_syllable_tl_to_bpm2(syllable: str) -> str:
    """
    將單一音節自台羅拼音轉成台語注音二式。
    - 保留尾隨的 '%'（部分字典用以標記特殊詞條）
    - 原音節若無調號數字，轉換後亦不強制補上調號
    """
    if not syllable:
        return syllable

    suffix = ""
    body = syllable
    if body.endswith("%"):
        suffix = "%"
        body = body[:-1]

    body_lower = body.lower()
    had_tone = bool(re.search(r"\d+$", body_lower))

    tlpa = convert_tl_to_tlpa(body_lower)
    if not tlpa:
        return syllable

    bpm2 = convert_TLPA_to_MPS2(tlpa)
    if not had_tone:
        # convert_tl_to_tlpa 對無調號音節會自動補 1；還原為無調號寫法
        bpm2 = re.sub(r"\d+$", "", bpm2)

    return bpm2 + suffix


def convert_code_tl_to_bpm2(code: str) -> str:
    """
    將【code】欄轉成台語注音二式。
    多音節以空白分隔（如 'kio3 si7'、'a tshoo'），逐音節轉換後再以空白接回。
    """
    if code is None:
        return ""
    code = str(code)
    if not code.strip():
        return code

    parts = code.split(" ")
    return " ".join(convert_syllable_tl_to_bpm2(p) if p else p for p in parts)


# =========================================================================
# 字典檔轉換
# =========================================================================
def output_filename_for(source_name: str) -> str:
    """ji_khoo_ziann_ji.dict.yaml → ji_khoo_ziann_ji_bpm2.dict.yaml"""
    if not source_name.endswith(".dict.yaml"):
        raise ValueError(f"非預期之字典檔名：{source_name}")
    base = source_name[: -len(".dict.yaml")]
    return f"{base}_bpm2.dict.yaml"


def transform_header_line(line: str) -> str:
    """
    調整檔頭（輸入／輸出皆為不含換行字元之單列文字）：
    1. name: xxx  → name: xxx_bpm2
    2. code 欄註解若提及台羅／TLPA，改為台語注音二式
    其餘（含分隔列「...」）原樣保留，避免破壞 RIME 字典結構。
    """
    stripped = line.strip()

    # name: ji_khoo_ziann_ji  → name: ji_khoo_ziann_ji_bpm2
    if stripped.startswith("name:"):
        name_value = stripped.split(":", 1)[1].strip()
        if name_value and not name_value.endswith("_bpm2"):
            prefix = line[: line.index("name:")]
            return f"{prefix}name: {name_value}_bpm2"
        return line

    # columns 區之 code 註解
    if "code" in stripped and ("TLPA" in stripped or "台羅" in stripped or "台灣音標" in stripped):
        return re.sub(r"#.*$", "# 台語注音二式（BPM2）拼音", line)

    return line


def transform_body_line(line: str) -> str:
    """
    轉換檔身一列：以 Tab 分欄，僅轉換第 2 欄（code）。
    空白列、註解列原樣保留。
    """
    if not line or line.startswith("#"):
        return line

    parts = line.split("\t")
    if len(parts) < 2:
        return line

    parts[1] = convert_code_tl_to_bpm2(parts[1])
    return "\t".join(parts)


def convert_dict_file(source_path: Path, target_path: Path) -> int:
    """
    將單一字典檔自台羅拼音轉成台語注音二式並寫出。
    回傳轉換之詞條列數。
    """
    text = source_path.read_text(encoding="utf-8")
    # 保留是否以換行結尾
    ends_with_newline = text.endswith("\n")
    lines = text.splitlines()

    out_lines: list[str] = []
    in_body = False
    entry_count = 0

    for line in lines:
        if not in_body:
            out_lines.append(transform_header_line(line))
            if line.strip() == "...":
                in_body = True
            continue

        new_line = transform_body_line(line)
        out_lines.append(new_line)
        # 有 Tab 且非註解，視為詞條
        if new_line and not new_line.startswith("#") and "\t" in new_line:
            entry_count += 1

    content = "\n".join(out_lines)
    if ends_with_newline:
        content += "\n"

    target_path.write_text(content, encoding="utf-8", newline="\n")
    return entry_count


def process() -> int:
    if not RIME_DIR.exists():
        print(f"❌ 找不到目錄：{RIME_DIR}")
        return EXIT_CODE_FAILURE

    print(f"📌 工作目錄：{RIME_DIR}")
    total_entries = 0
    converted_files = 0

    for source_name in SOURCE_DICT_FILES:
        source_path = RIME_DIR / source_name
        target_name = output_filename_for(source_name)
        target_path = RIME_DIR / target_name

        if not source_path.exists():
            print(f"⚠️ 跳過（找不到來源檔）：{source_name}")
            logging.warning("找不到來源字典檔：%s", source_path)
            continue

        try:
            count = convert_dict_file(source_path, target_path)
            converted_files += 1
            total_entries += count
            print(f"✅ {source_name} → {target_name}（詞條 {count} 筆）")
            logging.info("已轉換：%s → %s（%s 筆）", source_name, target_name, count)
        except Exception as e:
            print(f"❌ 轉換失敗：{source_name} — {e}")
            logging.error("轉換失敗 %s: %s", source_name, e, exc_info=True)
            return EXIT_CODE_FAILURE

    if converted_files == 0:
        print("❌ 未轉換任何字典檔。")
        return EXIT_CODE_FAILURE

    print(f"✅ 全部完成：{converted_files} 個檔案，合計詞條 {total_entries} 筆。")
    return EXIT_CODE_SUCCESS


# =========================================================================
# 主程式
# =========================================================================
def main() -> int:
    print("<=========== a830 作業開始 ===========>")
    print("台羅拼音字典檔 → 台語注音二式字典檔")
    result = process()
    print("<=========== a830 作業結束 ===========>")
    return result


if __name__ == "__main__":
    sys.exit(main())
