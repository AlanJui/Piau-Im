# 系統維護指引

## 漢字典

### 標音錯誤更正

【漢字庫】資料表中的【台語音標】，若有標音錯誤，其【更正】之操作方法，說明如下。

- 資料庫系統：SQLite
- 漢字資料庫：Ho_Lok_Ue.db
- 漢字庫資料表：漢字庫
- 資料庫用戶端工具：SQLite Studio 3.4.4

```sql
CREATE TABLE 漢字庫 (
    識別號  INTEGER PRIMARY KEY AUTOINCREMENT,
    漢字   TEXT    NOT NULL,
    台羅音標 TEXT    NOT NULL,
    常用度  REAL    DEFAULT 0.1,
    摘要說明 TEXT    DEFAULT 'NA',
    更新時間 TEXT    DEFAULT (DATETIME('now', 'localtime') )
                 NOT NULL
);
```

使用【SQL編輯器】，以 SQL Script 進行變更：

例如：

- 【千】： cian1 改成 tshian1
- 【志】： zi2 改成 tsi2

1. 更新操作：

```sql
UPDATE 漢字庫
SET 台羅音標 = 'tshian1'
WHERE 漢字 = '千';

UPDATE 漢字庫
SET 台羅音標 = 'tshian1'
WHERE 漢字 = '千';
```

2. 驗證撿視：

```sql
SELECT 漢字, 台羅音標
FROM 漢字庫
WHERE 漢字 IN ('千', '志');
```


### 批次更新（一）

若有多個漢字需要更新，可以使用 CASE 語句進行批次更新。

```sql
UPDATE 漢字庫
SET 台羅音標 = CASE
    WHEN 漢字 = '千' THEN 'tshian1'
    WHEN 漢字 = '志' THEN 'tsi2'
    ELSE 台羅音標
END
WHERE 漢字 IN ('千', '志');
```

### 批次更新（二）

c 開頭的台羅音標：將 c 替換為 tsh。

z 開頭的台羅音標：將 z 替換為 ts。

```sql
UPDATE OR REPLACE 漢字庫
SET 台羅音標 =
    CASE
        WHEN 台羅音標 LIKE 'c%' THEN 'tsh' || SUBSTR(台羅音標, 2)
        WHEN 台羅音標 LIKE 'z%' THEN 'ts' || SUBSTR(台羅音標, 2)
        ELSE 台羅音標
    END
WHERE 台羅音標 LIKE 'c%' OR 台羅音標 LIKE 'z%'
```

【註】：使用下述 Script 會發生【索引】有問題之執行錯誤。

```sql
UPDATE 漢字庫
SET 台羅音標 =
    CASE
        WHEN 台羅音標 LIKE 'c%' THEN 'tsh' || SUBSTR(台羅音標, 2)
        WHEN 台羅音標 LIKE 'z%' THEN 'ts' || SUBSTR(台羅音標, 2)
        ELSE 台羅音標
    END
WHERE 台羅音標 LIKE 'c%' OR 台羅音標 LIKE 'z%';
```

查核驗證

```sql
SELECT 漢字, 台羅音標
FROM 漢字庫
WHERE 台羅音標 LIKE 'c%' OR 台羅音標 LIKE 'z%';
```


## 小工具

### 統計誤用台語音標總數

```sql
SELECT 漢字,
       台羅音標,
       CASE
           WHEN 台羅音標 LIKE 'c%' THEN 'tsh' || SUBSTR(台羅音標, 2)
           WHEN 台羅音標 LIKE 'z%' THEN 'ts' || SUBSTR(台羅音標, 2)
           ELSE 台羅音標
       END AS 更新後音標
FROM 漢字庫
WHERE 台羅音標 LIKE 'c%' OR 台羅音標 LIKE 'z%'
GROUP BY 漢字, 更新後音標
HAVING COUNT(*) > 1;
```

### 檢視那些漢字誤用台語音標

```sql
SELECT 漢字,
       台羅音標,
       CASE
           WHEN 台羅音標 LIKE 'c%' THEN 'tsh' || SUBSTR(台羅音標, 2)
           WHEN 台羅音標 LIKE 'z%' THEN 'ts' || SUBSTR(台羅音標, 2)
           ELSE 台羅音標
       END AS 更新後音標
FROM 漢字庫
WHERE 台羅音標 LIKE 'c%' OR 台羅音標 LIKE 'z%'
GROUP BY 漢字, 更新後音標
HAVING COUNT(*) > 1;
```

### 以【時戳】改正【常用度】錯誤

```sql
UPDATE 漢字庫
SET 常用度 = 0.6
WHERE 更新時間 = '2025-02-12 16:43:41'
```

### 向 AI 請教之參考

[ChatGPT 4o](https://chatgpt.com/share/67ad5e29-caf4-8005-8997-67e07380d0a2)