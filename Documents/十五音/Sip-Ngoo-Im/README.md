# 雅俗通十五音漢字典


## 資料庫結構

### 漢字典（Han_Ji_Tian）

```sh
CREATE TABLE Han_Ji_Tian (
    識別號    INTEGER  NOT NULL
                    UNIQUE,
    漢字     TEXT,
    聲母     TEXT,
    韻母     TEXT,
    聲調     TEXT,
    常用度    REAL,
    台語音標拼音 TEXT,
    方音符號注音 TEXT,
    聲母識別號  INTEGER  NOT NULL,
    韻母識別號  INTEGER  NOT NULL,
    聲調識別號  INTEGER  NOT NULL,
    建立時間   DATETIME,
    PRIMARY KEY (
        識別號 AUTOINCREMENT
    ),
    FOREIGN KEY (
        聲母識別號
    )
    REFERENCES Siann_Bu_Piau (識別號),
    FOREIGN KEY (
        韻母識別號
    )
    REFERENCES Un_Bu_Piau (識別號),
    FOREIGN KEY (
        聲調識別號
    )
    REFERENCES Siann_Tiau_Piau (識別號) 
);
```

### 聲母表（Siann_Bu_Piau）

```sh
CREATE TABLE Siann_Bu_Piau (
    識別號   INTEGER NOT NULL
                  UNIQUE,
    十五音字母 TEXT,
    國際音標  TEXT,
    台語音標  TEXT,
    方音符號  TEXT,
    白話字   TEXT,
    台羅拚音  INTEGER,
    閩拼    TEXT,
    備註    TEXT,
    PRIMARY KEY (
        識別號 AUTOINCREMENT
    )
);
```

### 韻母表（Un_Bu_Piau）

```sh
CREATE TABLE Un_Bu_Piau (
    識別號   INTEGER NOT NULL
                  UNIQUE,
    韻母編碼  TEXT,
    十五音字母 TEXT,
    韻母序   INTEGER,
    舒促    TEXT,
    國際音標  TEXT,
    台語音標  TEXT,
    方音符號  TEXT,
    白話字   TEXT,
    台羅拚音  TEXT,
    閩拼    TEXT,
    PRIMARY KEY (
        識別號 AUTOINCREMENT
    )
);
```

### 聲調表（Siann_Tiau_Piau）

```sh
CREATE TABLE Siann_Tiau_Piau (
    識別號   INTEGER NOT NULL,
    聲調    TEXT,
    四聲調   TEXT,
    舒促聲   TEXT,
    台羅八聲調 INTEGER,
    PRIMARY KEY (
        識別號 AUTOINCREMENT
    )
);
```

## 查詢

### 漢字查詢


1. 建檢視

```sh
CREATE VIEW IF NOT EXISTS Han_Ji_Tian_View AS
SELECT 
    HJT.識別號 AS 識別號,
    HJT.漢字 AS 漢字,
    HJT.聲母 AS 十五音聲母,
    HJT.韻母 AS 十五音韻母,
    HJT.聲調 AS 十五音聲調,
    HJT.常用度 AS 常用度,
    SBP.台語音標 AS 聲母台語音標,
    UBP.台語音標 AS 韻母台語音標,
    SBP.方音符號 AS 聲母方音符號,
    UBP.方音符號 AS 韻母方音符號,
    STP.台羅八聲調 AS 八聲調,
    HJT.建立時間 AS 建立時間
FROM 
    Han_Ji_Tian HJT
LEFT JOIN 
    Siann_Bu_Piau SBP ON HJT.聲母識別號 = SBP.識別號
LEFT JOIN 
    Un_Bu_Piau UBP ON HJT.韻母識別號 = UBP.識別號
LEFT JOIN 
    Siann_Tiau_Piau STP ON HJT.聲調識別號 = STP.識別號;
```

2. 使用 SELECT 查詢

```sh
SELECT *
FROM Han_Ji_Tian_View
WHERE [漢字] = '不'
ORDER BY [建立時間] DESC, [常用度] DESC;
```

### 進階漢字讀音查詢

1. 建立【查漢字讀音檢視】

```sh
CREATE VIEW IF NOT EXISTS 查漢字讀音檢視 AS 
SELECT *
FROM Han_Ji_Tian_View
```

2. 使用【查漢字讀音檢視】，並要求查詢結果需依指定之欄位排序

```sh
SELECT *
FROM 查漢字讀音檢視
ORDER BY 建立時間 DESC, 常用度 DESC;
```



## 新增時間戳記欄位

1. 原資料表結構

```sh
CREATE TABLE Han_Ji_Tian (
    識別號    INTEGER NOT NULL
                   UNIQUE,
    漢字     TEXT,
    聲母     TEXT,
    韻母     TEXT,
    聲調     TEXT,
    常用度    REAL,
    台語音標聲母 TEXT,
    台語音標韻母 TEXT,
    台語音標拼音 TEXT,
    方音符號聲母 TEXT,
    方音符號韻母 TEXT,
    方音符號注音 TEXT,
    聲母識別號  INTEGER NOT NULL,
    韻母識別號  INTEGER NOT NULL,
    聲調識別號  INTEGER NOT NULL,
    PRIMARY KEY (
        識別號 AUTOINCREMENT
    ),
    FOREIGN KEY (
        聲母識別號
    )
    REFERENCES Siann_Bu_Piau (識別號),
    FOREIGN KEY (
        聲調識別號
    )
    REFERENCES Siann_Tiau_Piau (識別號),
    FOREIGN KEY (
        韻母識別號
    )
    REFERENCES Un_Bu_Piau (識別號) 
);
```

2. 新增時間戳記欄位

```sh
ALTER TABLE Han_Ji_Tian
ADD COLUMN 建立時間 DATETIME;
```

3. 為資料表所有紀錄，填入現在時刻

```sh
UPDATE Han_Ji_Tian
SET 建立時間 = CURRENT_TIMESTAMP
WHERE 建立時間 IS NULL;
```

4. 新增紀錄時，自動填入當時時刻

```sh
INSERT INTO Han_Ji_Tian (識別號, 漢字, 聲母, 韻母, 聲調, 常用度, 台語音標聲母, 台語音標韻母, 台語音標拼音, 方音符號聲母, 方音符號韻母, 方音符號注音, 聲母識別號, 韻母識別號, 聲調識別號, 建立時間)
VALUES (NULL, '漢字範例', 's', 'an', '1', 0.8, 's', 'an', 'san', 's', 'an', 'san', 1, 1, 1, CURRENT_TIMESTAMP);
```


## 刪除不需要的欄位

1. 建立原資料表的工作暫存表結構

```sh
CREATE TABLE Han_Ji_Tian_New (
    識別號    INTEGER NOT NULL UNIQUE,
    漢字     TEXT,
    聲母     TEXT,
    韻母     TEXT,
    聲調     TEXT,
    常用度    REAL,
    台語音標拼音 TEXT,
    方音符號注音 TEXT,
    聲母識別號  INTEGER NOT NULL,
    韻母識別號  INTEGER NOT NULL,
    聲調識別號  INTEGER NOT NULL,
    建立時間    DATETIME,
    PRIMARY KEY (識別號 AUTOINCREMENT),
    FOREIGN KEY (聲母識別號) REFERENCES Siann_Bu_Piau (識別號),
    FOREIGN KEY (韻母識別號) REFERENCES Un_Bu_Piau (識別號),
    FOREIGN KEY (聲調識別號) REFERENCES Siann_Tiau_Piau (識別號)
);
```

2. 自原資料表抄紀錄內容到工作暫存資料表中

```sh
INSERT INTO Han_Ji_Tian_New (識別號, 漢字, 聲母, 韻母, 聲調, 常用度, 台語音標拼音, 方音符號注音, 聲母識別號, 韻母識別號, 聲調識別號, 建立時間)
SELECT 識別號, 漢字, 聲母, 韻母, 聲調, 常用度, 台語音標拼音, 方音符號注音, 聲母識別號, 韻母識別號, 聲調識別號, 建立時間
FROM Han_Ji_Tian;
```

3. 將原資料表刪除

```sh
DROP TABLE IF EXISTS Han_Ji_Tian;
```

4. 將工作暫存資料表名稱改成原資料表使用之名稱

```sh
ALTER TABLE Han_Ji_Tian_New RENAME TO Han_Ji_Tian;
```

【備註】

在 SQLite 中，資料表 Han_Ji_Tian 可能有關聯的索引或外鍵，因此刪除該表時，需確保所有相關聯的對象（如索引、外鍵、觸發器）已被一併處理。

如果執行 DROP TABLE 或 ALTER TABLE 時遇到外鍵約束或其他問題，請檢查是否需要臨時禁用外鍵檢查：

```sh
PRAGMA foreign_keys = OFF;
```

執行完成後再重新開啟外鍵檢查：

```sh
PRAGMA foreign_keys = ON;
```

## 修訂資料紀錄

### 將某欄位的資料清空

資料表所有的【聲母】欄位，若其內容為 'q' ，需將之清除。

```sh
UPDATE Han_Ji_Tian
SET 聲母 = ''
WHERE 聲母 = 'q';
```

或

```sh
UPDATE Han_Ji_Tian
SET 聲母 = NULL
WHERE 聲母 = 'q';
```

### 將欄位開頭的第一個字元去除

若資料表中【台語音標拼音」欄位的第一個字元為 'q' ，需自記錄去除。

如：【由】字的【台語音標拼音】為： qiu5，改成：iu5 。

```sh
UPDATE Han_Ji_Tian
SET 台語音標拼音 = SUBSTR(台語音標拼音, 2)
WHERE 台語音標拼音 LIKE 'q%';
```

## 台羅音標漢字庫

### 資料表結構（Schema）

```bash
CREATE TABLE 台羅音標漢字庫 (
    識別號  INTEGER NOT NULL
                 UNIQUE,
    漢字   TEXT,
    台羅音標 TEXT,
    常用度  TEXT,
    摘要說明 TEXT,
    建立時間 TEXT    DEFAULT (DATETIME('now', 'localtime') ) 
                 NOT NULL,
    更新時間 TEXT    NOT NULL
                 DEFAULT (DATETIME('now', 'localtime') ),
    PRIMARY KEY (
        識別號 AUTOINCREMENT
    )
);
```


### 資料更新觸發器

```bash
DROP TRIGGER IF EXISTS 紀錄更新觸發器;

CREATE TRIGGER 紀錄更新觸發器
AFTER UPDATE ON 台羅音標漢字庫
FOR EACH ROW
WHEN NEW.更新時間 = OLD.更新時間
BEGIN
    UPDATE 台羅音標漢字庫
    SET 更新時間 = DATETIME('now', 'localtime')
    WHERE 識別號 = NEW.識別號;
END;
```

