import sqlite3

def query_for_tshiat_gu_siong_ji(tshiat_gu_siong_ji):
    # 創建數據庫連接
    conn = sqlite3.connect('.\\Kong_Un.db')

    # 創建一個游標
    cursor = conn.cursor()

    # 資料表欄位：序、發聲部位、聲母、台羅、IPA、清濁、切語下字
    fields = ['id', 'huat_siann_poo_ui', 'siann_bu', 'tai_lo', 'IPA', 'tshing_tok', 'tshiat_gu_siong_ji']

    # 執行 SQL 查詢
    cursor.execute("SELECT * FROM Tshiat_Gu_Siong_Ji WHERE 切語上字 LIKE ?", ('%' + tshiat_gu_siong_ji + '%',))

    # 獲取查詢結果
    results = cursor.fetchall()

    # 將結果轉換為字典列表
    dict_results = [dict(zip(fields, result)) for result in results]

    # 關閉數據庫連接
    conn.close()

    # 回傳字典列表
    return dict_results

if __name__ == "__main__":
    # 測試
    tshiat_gu_siong_ji = "徒"
    results = query_for_tshiat_gu_siong_ji(tshiat_gu_siong_ji)
    print(results)
    