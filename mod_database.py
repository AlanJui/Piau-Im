"""
資料庫連線管理模組
提供全域資料庫連線管理功能
"""
import os
import sqlite3
from contextlib import contextmanager
from typing import Optional

from dotenv import load_dotenv

# 載入環境變數
load_dotenv()
DB_PATH = os.getenv('DB_HO_LOK_UE', 'Ho_Lok_Ue.db')


class DatabaseManager:
    """資料庫管理器（單例模式）"""
    _instance: Optional['DatabaseManager'] = None

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._conn = None
        return cls._instance

    def connect(self, db_path: Optional[str] = None):
        """
        建立資料庫連線

        Args:
            db_path: 資料庫路徑，若為 None 則使用環境變數中的路徑

        Returns:
            sqlite3.Connection: 資料庫連線物件
        """
        if self._conn is None:
            self._conn = sqlite3.connect(db_path or DB_PATH)
        return self._conn

    def disconnect(self):
        """斷開資料庫連線"""
        if self._conn:
            self._conn.close()
            self._conn = None

    @property
    def connection(self):
        """
        取得資料庫連線（自動建立）

        Returns:
            sqlite3.Connection: 資料庫連線物件
        """
        if self._conn is None:
            self.connect()
        return self._conn

    @property
    def cursor(self):
        """
        取得資料庫游標

        Returns:
            sqlite3.Cursor: 資料庫游標物件
        """
        return self.connection.cursor()

    def execute(self, sql: str, params: tuple = ()):
        """
        執行 SQL 指令

        Args:
            sql: SQL 指令
            params: SQL 參數

        Returns:
            sqlite3.Cursor: 游標物件
        """
        cursor = self.connection.cursor()
        cursor.execute(sql, params)
        return cursor

    def executemany(self, sql: str, params_list: list):
        """
        批次執行 SQL 指令

        Args:
            sql: SQL 指令
            params_list: 參數列表

        Returns:
            sqlite3.Cursor: 游標物件
        """
        cursor = self.connection.cursor()
        cursor.executemany(sql, params_list)
        return cursor

    def commit(self):
        """提交交易"""
        if self._conn:
            self._conn.commit()

    def rollback(self):
        """回滾交易"""
        if self._conn:
            self._conn.rollback()

    @contextmanager
    def transaction(self):
        """
        交易 Context Manager

        使用範例:
            with db_manager.transaction():
                db_manager.execute("INSERT INTO table VALUES (?, ?)", (1, 2))
                # 自動 commit，若發生錯誤則 rollback
        """
        try:
            yield self
            self.commit()
        except Exception:
            self.rollback()
            raise

    def fetchone(self, sql: str, params: tuple = ()):
        """
        查詢單筆資料

        Args:
            sql: SQL 查詢指令
            params: SQL 參數

        Returns:
            tuple: 查詢結果
        """
        cursor = self.execute(sql, params)
        return cursor.fetchone()

    def fetchall(self, sql: str, params: tuple = ()):
        """
        查詢所有資料

        Args:
            sql: SQL 查詢指令
            params: SQL 參數

        Returns:
            list: 查詢結果列表
        """
        cursor = self.execute(sql, params)
        return cursor.fetchall()


# =========================================================================
# 建立全域單例
# =========================================================================
db_manager = DatabaseManager()


# =========================================================================
# 便利函數
# =========================================================================
def get_connection():
    """取得全域資料庫連線"""
    return db_manager.connection


def get_cursor():
    """取得全域資料庫游標"""
    return db_manager.cursor


def execute_query(sql: str, params: tuple = ()):
    """執行 SQL 查詢"""
    return db_manager.execute(sql, params)


def commit():
    """提交交易"""
    db_manager.commit()


def rollback():
    """回滾交易"""
    db_manager.rollback()


def disconnect():
    """斷開連線"""
    db_manager.disconnect()


# =========================================================================
# 測試程式
# =========================================================================
if __name__ == "__main__":
    # 測試連線
    print("測試資料庫連線...")

    # 方式 1：直接使用 db_manager
    cursor = db_manager.execute("SELECT COUNT(*) FROM 漢字庫")
    count = cursor.fetchone()[0]
    print(f"漢字庫記錄數：{count}")

    # 方式 2：使用交易
    try:
        with db_manager.transaction():
            # 這裡的操作會自動 commit 或 rollback
            cursor = db_manager.execute(
                "SELECT 漢字, 台羅音標 FROM 漢字庫 LIMIT 5"
            )
            results = cursor.fetchall()
            print("\n前 5 筆資料：")
            for han_ji, tai_lo in results:
                print(f"  {han_ji}: {tai_lo}")
    except Exception as e:
        print(f"錯誤：{e}")

    # 方式 3：使用便利函數
    cursor = execute_query("SELECT COUNT(*) FROM 漢字庫")
    count = cursor.fetchone()[0]
    print(f"\n使用便利函數查詢，記錄數：{count}")

    # 關閉連線
    db_manager.disconnect()
    print("\n✅ 測試完成")
