"""
漢字查字典模組
提供漢字查詢讀音功能
"""

import sqlite3
from contextlib import contextmanager
from pathlib import Path
from typing import Dict, List, Optional, Union

from mod_標音 import split_tai_gi_im_piau

# ============================================================================
# 資料庫連線管理
# ============================================================================

class HanJiTian:
    """漢字字典類別，管理資料庫連線和查詢"""

    def __init__(self, db_path: str = "Ho_Lok_Ue.db"):
        """
        初始化漢字字典

        Args:
            db_path: 資料庫檔案路徑
        """
        self.db_path = db_path

    @contextmanager
    def get_connection(self):
        """
        取得資料庫連線的 Context Manager
        使用方式：
            with han_ji_su_tian.get_connection() as conn:
                # 使用 conn 進行查詢
                pass
        """
        conn = None
        try:
            conn = sqlite3.connect(self.db_path)
            conn.row_factory = sqlite3.Row  # 讓查詢結果可以用欄位名稱存取
            yield conn
        except sqlite3.Error as e:
            print(f"資料庫錯誤: {e}")
            raise
        finally:
            if conn:
                conn.close()

    # ==========================================================
    # 用 `漢字` 查詢《台語音標》的讀音資訊
    # 在【台羅音標漢字庫】資料表結構中，以【常用度】欄位之值，區分【文讀音】與【白話音】。
    # 通用音：常用度 < 1.00；表文、白通用的讀音，最常用的讀音其值為 1.00，次常用的讀音值為 0.90，其餘則使用值為 0.89 ~ 0.81。
    # 文讀音：常用度 > 0.60；最常用的讀音其值為 0.80，次常用的讀音其值為 0.70；其餘則使用數值 0.69 ~ 0.61。
    # 白話音：常用度 > 0.40；最常用的讀音其值為 0.60，次常用的讀音其值為 0.50；其餘則使用數值 0.59 ~ 0.41。
    # 其　它：常用度 > 0.00；使用數值 0.40 ~ 0.01；使用時機為：（1）方言地方腔；(2) 罕見發音；(3) 尚未查證屬文讀音或白話音 。
    # ==========================================================
    def han_ji_ca_piau_im(
        self,
        han_ji: str,
        ue_im_lui_piat: str = "文讀音"
    ) -> Optional[List[Dict[str, Union[str, float]]]]:
        """
        根據漢字查詢其台羅音標及相關讀音資訊，並將台羅音標轉換為台語音標。
        若資料紀錄在常用度欄位儲存值為空值(NULL)，則將其視為 0，因此可排在查詢結果的最後。

        Args:
            han_ji: 欲查詢的漢字
            ue_im_lui_piat: 查詢的讀音類型，可以是 "文讀音"、"白話音"、"其它" 或 "全部"

        Returns:
            包含讀音資訊的字典列表，包含：
            - 識別號: 資料庫識別號
            - 漢字: 漢字
            - 台語音標: 台語音標（聲母+韻母+聲調）
            - 聲母: 聲母
            - 韻母: 韻母
            - 聲調: 聲調
            - 常用度: 常用度
            - 摘要說明: 摘要說明
            若查無資料則回傳 None

        範例:
            >>> su_tian = HanJiTian()
            >>> result = su_tian.han_ji_ca_piau_im("東", "白話音")
            >>> for item in result:
            >>>     print(f"{item['台語音標']} (常用度: {item['常用度']})")
        """
        with self.get_connection() as conn:
            cursor = conn.cursor()

            # 將文白通用音視為第一優選
            common_reading_condition = "常用度 >= 0.81 AND 常用度 <= 1.0"

            # 根據不同讀音類型，添加額外的查詢條件
            if ue_im_lui_piat == "文讀音":
                reading_condition = f"({common_reading_condition}) OR (常用度 >= 0.61 AND 常用度 < 0.81)"
            elif ue_im_lui_piat == "白話音":
                reading_condition = f"({common_reading_condition}) OR (常用度 > 0.40 AND 常用度 < 0.61)"
            elif ue_im_lui_piat == "其它":
                reading_condition = "常用度 > 0.00 AND 常用度 <= 0.40"
            else:
                reading_condition = "1=1"  # 查詢所有

            query = f"""
            SELECT
                識別號,
                漢字,
                台羅音標,
                常用度,
                摘要說明
            FROM
                漢字庫
            WHERE
                漢字 = ? AND ({reading_condition})
            ORDER BY
                COALESCE(常用度, 0) DESC;
            """

            cursor.execute(query, (han_ji,))
            results = cursor.fetchall()

            # 如果沒有找到符合條件的讀音，則查詢所有讀音，並選擇常用度最高者
            if not results:
                query = """
                SELECT
                    識別號,
                    漢字,
                    台羅音標,
                    常用度,
                    摘要說明
                FROM
                    漢字庫
                WHERE
                    漢字 = ?
                ORDER BY
                    COALESCE(常用度, 0) DESC
                LIMIT 1;
                """
                cursor.execute(query, (han_ji,))
                results = cursor.fetchall()

            # 若仍無結果，回傳 None
            if not results:
                return None

            # 將結果轉換為字典列表
            fields = ['識別號', '漢字', '台語音標', '常用度', '摘要說明']

            data = []
            for result in results:
                row_dict = dict(zip(fields, result))
                # 取得台羅音標
                tai_loo_im = row_dict['台語音標']

                # 將台羅音標轉換為台語音標
                split_result = split_tai_gi_im_piau(tai_loo_im)
                row_dict['聲母'] = split_result[0]
                row_dict['韻母'] = split_result[1]
                row_dict['聲調'] = split_result[2]

                # 更新 row_dict 中的台語音標
                row_dict['台語音標'] = f'{row_dict["聲母"]}{row_dict["韻母"]}{row_dict["聲調"]}'

                # 將結果加入列表
                data.append(row_dict)

            return data


# ============================================================================
# 獨立函數介面（方便直接呼叫）
# ============================================================================

def han_ji_ca_piau_im(
    han_ji: str,
    ue_im_lui_piat: str = "文讀音",
    db_path: str = "Ho_Lok_Ue.db"
) -> Optional[List[Dict[str, Union[str, float]]]]:
    """
    查詢漢字的讀音（獨立函數版本）

    Args:
        han_ji: 要查詢的漢字
        ue_im_lui_piat: 讀音類型 ("文讀音"、"白話音"、"其它" 或 "全部")
        db_path: 資料庫檔案路徑

    Returns:
        讀音列表，每個讀音為一個字典

    範例:
        >>> from mod_ca_ji_tian import han_ji_ca_piau_im
        >>> result = han_ji_ca_piau_im("東", "白話音")
        >>> print(result)
    """
    su_tian = HanJiTian(db_path)
    return su_tian.han_ji_ca_piau_im(han_ji, ue_im_lui_piat)


# ============================================================================
# 批次查詢函數
# ============================================================================

def han_ji_ca_piau_im_list(
    han_ji_list: List[str],
    ue_im_lui_piat: str = "文讀音",
    db_path: str = "Ho_Lok_Ue.db"
) -> Dict[str, Optional[List[Dict[str, Union[str, float]]]]]:
    """
    批次查詢多個漢字的讀音

    Args:
        han_ji_list: 要查詢的漢字列表
        ue_im_lui_piat: 讀音類型
        db_path: 資料庫檔案路徑

    Returns:
        字典，key 為漢字，value 為該字的讀音列表

    範例:
        >>> result = han_ji_ca_piau_im_list(["東", "西", "南", "北"], "白話音")
        >>> for han_ji, piau_im_list in result.items():
        >>>     print(f"{han_ji}: {piau_im_list}")
    """
    su_tian = HanJiTian(db_path)
    results = {}

    for han_ji in han_ji_list:
        results[han_ji] = su_tian.han_ji_ca_piau_im(han_ji, ue_im_lui_piat)

    return results


# ============================================================================
# 測試程式
# ============================================================================

def test():
    """測試函數"""

    # 測試 1: 使用類別方式查詢文讀音
    print("=" * 70)
    print("測試 1: 使用 HanJiTian 類別查詢文讀音")
    print("=" * 70)

    su_tian = HanJiTian("Ho_Lok_Ue.db")

    test_chars = ["東", "西", "南", "北", "中"]
    for han_ji in test_chars:
        result = su_tian.han_ji_ca_piau_im(han_ji, ue_im_lui_piat="文讀音")
        print(f"\n漢字: {han_ji} (文讀音)")
        if result:
            for item in result:
                print(f"  台語音標: {item['台語音標']}, 常用度: {item['常用度']}, 說明: {item['摘要說明']}")
        else:
            print(f"  查無資料")

    # 測試 2: 查詢白話音
    print("\n" + "=" * 70)
    print("測試 2: 查詢白話音")
    print("=" * 70)

    for han_ji in ["東", "西"]:
        result = su_tian.han_ji_ca_piau_im(han_ji, ue_im_lui_piat="白話音")
        print(f"\n漢字: {han_ji} (白話音)")
        if result:
            for item in result:
                print(f"  台語音標: {item['台語音標']}, 聲母: {item['聲母']}, 韻母: {item['韻母']}, 聲調: {item['聲調']}")
        else:
            print(f"  查無資料")

    # 測試 3: 使用獨立函數
    print("\n" + "=" * 70)
    print("測試 3: 使用獨立函數 han_ji_ca_piau_im()")
    print("=" * 70)

    result = han_ji_ca_piau_im("東", ue_im_lui_piat="全部")
    print(f"\n查詢「東」的所有讀音:")
    if result:
        for item in result:
            print(f"  {item['台語音標']} (常用度: {item['常用度']})")

    # 測試 4: 批次查詢
    print("\n" + "=" * 70)
    print("測試 4: 批次查詢")
    print("=" * 70)

    results = han_ji_ca_piau_im_list(["東", "西", "南", "北"], ue_im_lui_piat="白話音")
    for han_ji, piau_im_list in results.items():
        print(f"\n{han_ji}:")
        if piau_im_list:
            for item in piau_im_list:
                print(f"  {item['台語音標']} (常用度: {item['常用度']})")


if __name__ == "__main__":
    test()
