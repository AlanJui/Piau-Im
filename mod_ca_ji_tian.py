"""
mod_ca_ji_tian.py V0.2.3

功能說明：
漢字查字典模組，提供漢字查詢讀音功能

更新紀錄：
 - v0.2.1 2026-02-14: 變更讀音查詢作法，預設為：不分文讀/白話音，查詢結果依
    常用度之由大到小排序；新增參數 `display_all_piau_im`，用於控制是否顯示
    所有讀音（包含常用度 > 0.00 的讀音），預設為 True。
 - v0.2.2 2026-03-211: 修訂 `han_ji_ca_piau_im()` 查漢字標音之常用度用法。
 - v0.2.3 2026-07-15: `han_ji_ca_piau_im()` 新增選用參數 `tai_lo_im_piau`，
    可依【台羅音標】篩選；查詢結果於【常用度】相同時，改依【最近揀用時間】
    （由人工校正程式回寫）由新至舊排序，令最近人工揀用之讀音優先。
"""

import sqlite3
from contextlib import contextmanager
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
        self._persistent_conn = None
        self._time_order_column: Optional[str] = None  # 次要排序鍵欄位名稱（快取）

    def connect(self):
        """建立持續性資料庫連線"""
        if self._persistent_conn is None:
            self._persistent_conn = sqlite3.connect(self.db_path)
            self._persistent_conn.row_factory = sqlite3.Row

    def disconnect(self):
        """關閉持續性資料庫連線"""
        if self._persistent_conn:
            self._persistent_conn.close()
            self._persistent_conn = None

    @contextmanager
    def get_connection(self):
        """
        取得資料庫連線的 Context Manager
        使用方式：
            with han_ji_su_tian.get_connection() as conn:
                # 使用 conn 進行查詢
                pass
        """
        # 如果已有持續性連線，則直接使用
        if self._persistent_conn:
            yield self._persistent_conn
            return

        # 否則建立臨時連線
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
    def _get_time_order_column(self, conn) -> str:
        """
        取得查詢結果之次要排序鍵欄位名稱（結果會快取）。
        優先使用【最近揀用時間】（由人工校正程式回寫，代表最近一次人工揀用此讀音之時間）；
        若資料庫尚無此欄位，退而使用【更新時間】，以維持向後相容。
        """
        if self._time_order_column is None:
            cols = [row[1] for row in conn.execute("PRAGMA table_info(漢字庫)").fetchall()]
            self._time_order_column = "最近揀用時間" if "最近揀用時間" in cols else "更新時間"
        return self._time_order_column

    def han_ji_ca_piau_im(
        self,
        han_ji: str,
        ue_im_lui_piat: str = "文讀音",
        display_all_piau_im: bool = False,
        tai_lo_im_piau: Optional[str] = None,
    ) -> Optional[List[Dict[str, Union[str, float]]]]:
        """
        根據漢字查詢其台羅音標及相關讀音資訊，並將台羅音標轉換為台語音標。
        若資料紀錄在常用度欄位儲存值為空值(NULL)，則將其視為 0，因此可排在查詢結果的最後。
        查詢結果排序規則：
        1. 常用度由大至小；
        2. 常用度相同時，依【最近揀用時間】由新至舊（最近人工揀用之讀音優先）。

        Args:
            han_ji: 欲查詢的漢字
            ue_im_lui_piat: 查詢的讀音類型，可以是 "文讀音"、"白話音"、"其它" 或 "全部"
            tai_lo_im_piau: （選用）台羅音標（與資料庫【台羅音標】欄位同格式，如 "tong1"）。
                指定時僅回傳【漢字】與【台羅音標】皆相符之讀音；適用於呼叫端已知音標、
                欲取回該筆讀音完整資訊之場景（如自標音字庫回填）。

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
            common_reading_condition = "常用度 > 0.80 AND 常用度 <= 1.0"

            # 根據不同讀音類型，添加額外的查詢條件
            if display_all_piau_im:
                reading_condition = "常用度 > 0.00 AND 常用度 <= 1.00"
            else:
                if ue_im_lui_piat == "文讀音":
                    reading_condition = f"({common_reading_condition}) OR (常用度 > 0.60 AND 常用度 <= 0.80)"
                elif ue_im_lui_piat == "白話音":
                    reading_condition = f"({common_reading_condition}) OR (常用度 > 0.40 AND 常用度 <= 0.60)"
                elif ue_im_lui_piat == "其它":
                    reading_condition = "常用度 > 0.00 AND 常用度 <= 0.40"
                else:
                    reading_condition = "常用度 > 0.00 AND 常用度 <= 1.00"  # 查詢所有
            # if ue_im_lui_piat == "文讀音":
            #     # reading_condition = f"(常用度 > 0.60 AND 常用度 <= 0.80) OR ({common_reading_condition})"
            #     reading_condition = f"({common_reading_condition}) OR (常用度 > 0.60 AND 常用度 <= 0.80)"
            # elif ue_im_lui_piat == "白話音":
            #     # reading_condition = f"(常用度 > 0.40 AND 常用度 <= 0.60) OR ({common_reading_condition})"
            #     reading_condition = f"({common_reading_condition}) OR (常用度 > 0.40 AND 常用度 <= 0.60)"
            # elif ue_im_lui_piat == "其它":
            #     reading_condition = "常用度 > 0.00 AND 常用度 <= 0.40"
            # else:
            #     reading_condition = "常用度 > 0.00 AND 常用度 <= 1.00"

            # 次要排序鍵：常用度相同時，最近人工揀用之讀音優先
            time_col = self._get_time_order_column(conn)
            order_by = f"COALESCE(常用度, 0) DESC, COALESCE({time_col}, '') DESC"

            # 若呼叫端指定【台羅音標】，加入篩選條件（第一優先：漢字＋台羅音標＋常用度）
            im_piau_condition = " AND 台羅音標 = ?" if tai_lo_im_piau else ""
            params: tuple = (han_ji, tai_lo_im_piau) if tai_lo_im_piau else (han_ji,)

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
                漢字 = ?{im_piau_condition} AND ({reading_condition})
            ORDER BY
                {order_by};
            """

            cursor.execute(query, params)
            results = cursor.fetchall()

            # 若指定【台羅音標】但於該讀音類別中查無資料，放寬讀音類別限制再查一次
            if not results and tai_lo_im_piau:
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
                    漢字 = ? AND 台羅音標 = ?
                ORDER BY
                    {order_by};
                """
                cursor.execute(query, (han_ji, tai_lo_im_piau))
                results = cursor.fetchall()

            # 如果沒有找到符合條件的讀音，則查詢所有讀音，並選擇常用度最高者
            if not results:
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
                    漢字 = ?
                ORDER BY
                    {order_by}
                LIMIT 1;
                """
                cursor.execute(query, (han_ji,))
                results = cursor.fetchall()

            # 若仍無結果，回傳 None
            if not results:
                return None

            # 將結果轉換為字典列表
            fields = ["識別號", "漢字", "台語音標", "常用度", "摘要說明"]

            data = []
            for result in results:
                row_dict = dict(zip(fields, result))
                # 取得台羅音標
                tai_loo_im = row_dict["台語音標"]

                # 將台羅音標轉換為台語音標
                split_result = split_tai_gi_im_piau(tai_loo_im)
                row_dict["聲母"] = split_result[0]
                row_dict["韻母"] = split_result[1]
                row_dict["聲調"] = split_result[2]

                # 更新 row_dict 中的台語音標
                row_dict["台語音標"] = (
                    f'{row_dict["聲母"]}{row_dict["韻母"]}{row_dict["聲調"]}'
                )

                # 將結果加入列表
                data.append(row_dict)

            return data


# ============================================================================
# 獨立函數介面（方便直接呼叫）
# ============================================================================


def han_ji_ca_piau_im(
    han_ji: str,
    ue_im_lui_piat: str = "文讀音",
    db_path: str = "Ho_Lok_Ue.db",
    tai_lo_im_piau: Optional[str] = None,
) -> Optional[List[Dict[str, Union[str, float]]]]:
    """
    查詢漢字的讀音（獨立函數版本）

    Args:
        han_ji: 要查詢的漢字
        ue_im_lui_piat: 讀音類型 ("文讀音"、"白話音"、"其它" 或 "全部")
        db_path: 資料庫檔案路徑
        tai_lo_im_piau: （選用）台羅音標，指定時僅回傳漢字與音標皆相符之讀音

    Returns:
        讀音列表，每個讀音為一個字典

    範例:
        >>> from mod_ca_ji_tian import han_ji_ca_piau_im
        >>> result = han_ji_ca_piau_im("東", "白話音")
        >>> print(result)
    """
    su_tian = HanJiTian(db_path)
    return su_tian.han_ji_ca_piau_im(han_ji, ue_im_lui_piat, tai_lo_im_piau=tai_lo_im_piau)


# ============================================================================
# 批次查詢函數
# ============================================================================


def han_ji_ca_piau_im_list(
    han_ji_list: List[str],
    ue_im_lui_piat: str = "文讀音",
    db_path: str = "Ho_Lok_Ue.db",
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
    ji_tian = HanJiTian(db_path)
    results = {}

    for han_ji in han_ji_list:
        results[han_ji] = ji_tian.han_ji_ca_piau_im(han_ji, ue_im_lui_piat)

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

    ji_tian = HanJiTian("Ho_Lok_Ue.db")

    # test_chars = ["東", "西", "南", "北", "中"]
    test_chars = ["白", "石", "當", "中", "隆"]
    for han_ji in test_chars:
        result = ji_tian.han_ji_ca_piau_im(han_ji, ue_im_lui_piat="文讀音")
        print(f"\n漢字: {han_ji} (文讀音)")
        if result:
            for item in result:
                print(
                    f"  台語音標: {item['台語音標']}, 常用度: {item['常用度']}, 說明: {item['摘要說明']}"
                )
        else:
            print("  查無資料")

    # 測試 2: 查詢白話音
    print("\n" + "=" * 70)
    print("測試 2: 查詢白話音")
    print("=" * 70)

    for han_ji in ["東", "西"]:
        result = ji_tian.han_ji_ca_piau_im(han_ji, ue_im_lui_piat="白話音")
        print(f"\n漢字: {han_ji} (白話音)")
        if result:
            for item in result:
                print(
                    f"  台語音標: {item['台語音標']}, 聲母: {item['聲母']}, 韻母: {item['韻母']}, 聲調: {item['聲調']}"
                )
        else:
            print("  查無資料")

    # 測試 3: 使用獨立函數
    print("\n" + "=" * 70)
    print("測試 3: 使用獨立函數 han_ji_ca_piau_im()")
    print("=" * 70)

    result = han_ji_ca_piau_im("東", ue_im_lui_piat="全部")
    print("\n查詢「東」的所有讀音:")
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
