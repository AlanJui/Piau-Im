import sqlite3

def query_tshiat_gu(table_name, fields, query_field, keyword):
    # 創建數據庫連接
    conn = sqlite3.connect('.\\Kong_Un.db')

    # 創建一個游標
    cursor = conn.cursor()

    # 執行 SQL 查詢
    cursor.execute(f"SELECT * FROM {table_name} WHERE {query_field} LIKE ?", ('%' + keyword + '%',))

    # 獲取查詢結果
    results = cursor.fetchall()

    # 將結果轉換為字典列表
    dict_results = [dict(zip(fields, result)) for result in results]

    # 關閉數據庫連接
    conn.close()

    # 回傳字典列表
    return dict_results

if __name__ == "__main__":
    # 測試 "查詢切語上字" 
    table_name = "Tshiat_Gu_Siong_Ji"
    # 資料表欄位：序、發聲部位、聲母、台羅、擬音、清濁、切語上字
    fields = ['id', 'huat_siann_poo_ui', 'siann_bu', 'tai_lo', 'IPA', 'tshing_tok', 'tshiat_gu_siong_ji']
    query_field = "切語上字"
    keyword = "徒"
    results = query_tshiat_gu(table_name, fields, query_field, keyword)
    print(results)

    # 測試 "查詢切語下字" 
    table_name = "Tshiat_Gu_Ha_Ji"
    # 資料表欄位：序、攝、韻系、四聲、等弟、開合、台羅、擬音、切語下字
    fields = ['id', 'liap', 'un_he', 'su_sing', 'ting_te', 'khai_hap', 'tai_lo', 'IPA', 'tshing_tok', 'tshiat_gu_ha_ji']
    query_field = "切語下字"
    keyword = "紅"
    results = query_tshiat_gu(table_name, fields, query_field, keyword)
    print(results)