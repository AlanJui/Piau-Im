"""
漢字標音轉換模組
"""

# =========================================================================
# 常數定義
# =========================================================================
# 定義 Exit Code
EXIT_CODE_SUCCESS = 0  # 成功
EXIT_CODE_FAILURE = 1  # 失敗
EXIT_CODE_INVALID_INPUT = 2  # 輸入錯誤
EXIT_CODE_SAVE_FAILURE = 3  # 儲存失敗
EXIT_CODE_PROCESS_FAILURE = 10  # 過程失敗
EXIT_CODE_NO_FILE = 90 # 無法找到檔案
EXIT_CODE_UNKNOWN_ERROR = 99  # 未知錯誤

# =========================================================================
# 設定日誌
# =========================================================================
import re
import unicodedata
from typing import Optional, Tuple

from mod_logging import (
    init_logging,
    logging_exc_error,
    logging_exception,
    logging_process_step,
    logging_warning,
)

init_logging()

#============================================================================
# 音節尾字為調號（數字）擷取函數
#============================================================================
# 常用上標轉換表（補足您可能遇到的上標字元）
_SUPERSCRIPT_MAP = {
    '\u2070': '0',  # ⁰
    '\u00B9': '1',  # ¹
    '\u00B2': '2',  # ²
    '\u00B3': '3',  # ³
    '\u2074': '4',  # ⁴
    '\u2075': '5',  # ⁵
    '\u2076': '6',  # ⁶
    '\u2077': '7',  # ⁷
    '\u2078': '8',  # ⁸
    '\u2079': '9',  # ⁹
}
_SUPER_TRANS = str.maketrans(_SUPERSCRIPT_MAP)

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
