# Nvim 操作指引

```sh
"""
查詢某漢字的 `小韻` 資料
"""
def han_ji_cha_siau_un(cursor, han_ji):
    # SQL 查詢語句
    query = """
    SELECT *
    FROM 小韻查詢
    WHERE 小韻字 = ?;
    """

    # 執行 SQL 查詢
    cursor.execute(query, (han_ji,))

    # 獲取查詢結果
    results = cursor.fetchall()

    # 將結果轉換為字典列表
    fields = ['小韻字', '切語', '標音', '目次編碼', '小韻字序號', '小韻字集', '字數',
        '發音部位', '聲母', '清濁', '發送收', '聲母拼音碼', '切語上字集',
        '韻系列號', '韻系行號', '韻目索引', '目次', '攝', '韻系',
        '韻目', '調', '呼', '等', '韻母', '切語下字集', '等呼', '韻母拼音碼']

    dict_results = [dict(zip(fields, result)) for result in results]

    # 回傳字典列表
    return dict_results
```

使用 Search and Replace 功能，將以下的字串，透過正規式變更成 `單引號` 所包：

```sh
    小韻字,
  	切語,
	拼音,
    目次編碼,
    小韻字序號,
    小韻字集,
    字數,
    發音部位,
    聲母,
    清濁,
    發送收,
    聲母拼音碼,
    切語上字集,
    韻系列號,
    韻系行號,
    韻目索引,
    目次,
    攝,
    韻系,
    韻目,
    調,
    呼,
    等,
    韻母,
    切語下字集,
    等呼,
    韻母拼音碼
```

在 Nvim 的指令列執行 Search and Replace 功能：

```sh
:%s/\v\s*(\S+),/'\1',/g
```

- \v : 啟用 Very Magic Mode
- \s\* : 將左道的空白字元清空
- (\S+) : 將 , 號左邊的漢字選取
