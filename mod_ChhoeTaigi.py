from urllib.parse import quote

import requests


def moe_dict_search(query: str, verbose: bool = True):
    """
    使用萌典台語資料 API 查詢

    Args:
        query: 要查詢的漢字
        verbose: 是否顯示詳細資訊

    Returns:
        list: 包含所有讀音的列表，例如 ["pe̍h", "pi̍k"]
              若查無資料則返回空列表 []
    """
    # 萌典台語 API 端點
    api_endpoints = [
        f"https://www.moedict.tw/t/{quote(query)}.json",
        f"https://moedict.tw/t/{quote(query)}.json",
    ]

    headers = {
        "Accept": "application/json",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
    }

    if verbose:
        print(f"\n查詢：{query}")

    for idx, url in enumerate(api_endpoints, 1):
        try:
            if verbose:
                print(f"\n[嘗試 {idx}] {url}")

            response = requests.get(url, headers=headers, timeout=10)

            if response.status_code != 404 and verbose:
                print(f"狀態碼：{response.status_code}")

            response.raise_for_status()

            content_type = response.headers.get('Content-Type', '')
            if verbose:
                print(f"Content-Type：{content_type}")

            # 檢查是否為 JSON
            if 'json' not in content_type.lower() and not url.endswith('.json'):
                if verbose:
                    print(f"⚠️  回應非 JSON 格式")
                continue

            data = response.json()

            if not data:
                if verbose:
                    print(f"⚠️  查無資料")
                continue

            # 成功找到資料 - 提取所有讀音
            pronunciations = []

            if 'h' in data:
                for heteronym in data['h']:
                    if 'T' in heteronym:
                        pronunciations.append(heteronym['T'])

            if verbose:
                print(f"\n✅ 找到資料")
                print("=" * 70)
                print(f"詞目：{data.get('t', query)}")
                print(f"讀音列表：{pronunciations}")

                # 顯示詳細資訊
                if 'h' in data:
                    heteronyms = data['h']
                    for h_idx, heteronym in enumerate(heteronyms, 1):
                        if len(heteronyms) > 1:
                            print(f"\n【讀音 {h_idx}】")
                        else:
                            print()

                        # 顯示台羅拼音
                        if 'T' in heteronym:
                            print(f"台羅拼音：{heteronym['T']}")

                        # 顯示白話字
                        if 'P' in heteronym:
                            print(f"白話字：{heteronym['P']}")

                        # 顯示釋義
                        if 'd' in heteronym:
                            definitions = heteronym['d']
                            print(f"\n釋義：")
                            for d_idx, definition in enumerate(definitions, 1):
                                if 'f' in definition:
                                    explanation = definition['f']
                                    # 移除標記符號
                                    explanation = explanation.replace('`', '').replace('~', '')
                                    print(f"  {d_idx}. {explanation}")

                                # 顯示例句
                                if 'e' in definition:
                                    examples = definition['e']
                                    for example in examples:
                                        example_text = example.replace('`', '').replace('~', '')
                                        print(f"     例：{example_text}")

                        print("-" * 70)

            return pronunciations  # 返回讀音列表

        except requests.exceptions.Timeout:
            if verbose:
                print(f"⚠️  請求超時")
            continue
        except requests.exceptions.RequestException as e:
            if "404" not in str(e) and verbose:
                print(f"⚠️  錯誤：{e}")
            continue
        except ValueError as e:
            if verbose:
                print(f"⚠️  JSON 解析錯誤：{e}")
            continue
        except Exception as e:
            if verbose:
                print(f"⚠️  錯誤：{e}")
            continue

    if verbose:
        print(f"\n❌ 所有 API 端點都無法使用")

    return []  # 查無資料時返回空列表

def chhoe_taigi(han_ji: str, verbose: bool = True):
    """
    使用萌典台語辭典查詢

    Args:
        han_ji: 要查詢的漢字
        verbose: 是否顯示詳細資訊

    Returns:
        list: 包含所有讀音的列表
    """
    return moe_dict_search(han_ji, verbose)

def ut01_chhoe_taigi(han_ji: str):
    """測試查詢單個漢字"""
    result = chhoe_taigi(han_ji)
    print(f"\n返回結果：{result}")
    return result

def ut02_chhoe_taigi():
    """測試查詢多個漢字"""
    test_chars = ["愛", "白", "隆"]
    for char in test_chars:
        result = chhoe_taigi(char)
        print(f"\n返回結果：{result}")
        print("=" * 100)

def ut03_simple_query():
    """簡單查詢模式（不顯示詳細資訊）"""
    print("\n=== 簡單查詢模式 ===")
    chars = ["白", "愛", "隆"]
    for char in chars:
        result = chhoe_taigi(char, verbose=False)
        print(f"{char}: {result}")

if __name__ == "__main__":
    print("=== 單元測試 01：查詢單個漢字 ===")
    ut01_chhoe_taigi('白')
    print("=" * 100)
    ut01_chhoe_taigi('隆')
    print("=" * 100)
    print("=== 單元測試 01：查詢漢字 ===")
    # ut01_chhoe_taigi('狎')
    # ut01_chhoe_taigi('愛')
    # ut01_chhoe_taigi('隆')
    ut01_chhoe_taigi('白')
    # print("=== 單元測試 02：查詢多個詞彙 ===")
    # ut02_chhoe_taigi()