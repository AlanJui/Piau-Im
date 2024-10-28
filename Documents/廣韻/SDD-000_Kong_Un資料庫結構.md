# Kong_Un.db 資料庫結構

## 廣韻漢字庫

```sql
CREATE TABLE 廣韻漢字庫 (
    字號    INTEGER NOT NULL
                  UNIQUE,
    漢字    TEXT,
    標音    TEXT,
    常用度   REAL,
    上字    TEXT,
    下字    TEXT,
    上字號   INTEGER,
    聲母    TEXT,
    聲母標音  TEXT,
    七聲類   TEXT,
    清濁    TEXT,
    發送收   TEXT,
    下字號   INTEGER,
    韻母    TEXT,
    韻母標音  TEXT,
    韻目列號  INTEGER,
    攝     TEXT,
    調     TEXT,
    目次    TEXT,
    韻目    TEXT,
    等呼    TEXT,
    等     INTEGER,
    呼     TEXT,
    廣韻調名  TEXT,
    台羅聲調  INTEGER,
    字義識別號 INTEGER,
    PRIMARY KEY (
        字號 AUTOINCREMENT
    ),
    FOREIGN KEY (
        上字號
    )
    REFERENCES 切語上字表 (識別號)
);
```

## 字義資料表

```sql
CREATE TABLE 字義資料表 (
    識別號   INTEGER NOT NULL
                  UNIQUE,
    廣韻小字序 REAL,
    漢字    TEXT,
    字義摘要  TEXT,
    PRIMARY KEY (
        識別號 AUTOINCREMENT
    )
);
```

## 切語上字表

```sql
CREATE TABLE 切語上字表 (
    識別號   INTEGER NOT NULL
                  UNIQUE,
    七聲類   TEXT,
    發音部位  TEXT,
    聲母    TEXT,
    清濁    TEXT,
    發送收   TEXT,
    聲母標音  TEXT,
    切語上字集 TEXT,
    備註    TEXT,
    PRIMARY KEY (
        識別號 AUTOINCREMENT
    )
);
```

## 聲母對照表

```sql
CREATE TABLE 聲母對照表 (
    識別號    TEXT,
    國際音標聲母 TEXT,
    台語音標聲母 TEXT,
    台羅聲母   TEXT,
    白話字聲母  TEXT,
    閩拼聲母   TEXT,
    方音聲母   TEXT,
    十五音聲母  TEXT
);
```

## 切語下字表

```sql
CREATE TABLE 切語下字表 (
    識別號      INTEGER NOT NULL
                     UNIQUE,
    韻目識別號    INTEGER,
    韻目列號     INTEGER,
    攝        TEXT,
    四聲調號     INTEGER,
    調        TEXT,
    目次       TEXT,
    韻目       TEXT,
    韻類       TEXT,
    等呼       TEXT,
    等        INTEGER,
    呼        TEXT,
    韻母       TEXT,
    韻目標音對照識別 INTEGER,
    韻母標音     TEXT,
    切語下字集    TEXT,
    備註       TEXT,
    PRIMARY KEY (
        識別號 AUTOINCREMENT
    ),
    FOREIGN KEY (
        韻目識別號
    )
    REFERENCES 韻目資料表 (韻目識別號)
);
```

## 韻目標音對照表

```sql
CREATE TABLE 韻目標音對照表 (
    識別號   INTEGER NOT NULL
                  UNIQUE,
    韻目    TEXT,
    韻目識別號 INTEGER,
    韻母    TEXT,
    韻目列號  INTEGER,
    韻攝    TEXT,
    韻類    TEXT,
    四聲韻目  TEXT,
    等呼    TEXT,
    等     INTEGER,
    呼     TEXT,
    舒聲標音  TEXT,
    促聲標音  TEXT,
    PRIMARY KEY (
        識別號 AUTOINCREMENT
    )
);
```

## 韻目資料表

```sql
CREATE TABLE 韻目資料表 (
    韻目識別號   INTEGER NOT NULL
                    UNIQUE,
    韻目      TEXT,
    韻目方陣識別號 INTEGER,
    韻目列號    INTEGER,
    韻攝      TEXT,
    四聲調號    INTEGER,
    調       TEXT,
    廣韻聲調    TEXT,
    舒促聲     TEXT,
    目次序號    INTEGER,
    目次      TEXT,
    PRIMARY KEY (
        韻目識別號 AUTOINCREMENT
    ),
    FOREIGN KEY (
        韻目列號
    )
    REFERENCES 韻攝清單 (韻目列號),
    FOREIGN KEY (
        韻目方陣識別號
    )
    REFERENCES 韻目方陣表 (識別)
);
```


## 韻母對照表

```sql
CREATE TABLE 韻母對照表 (
    識別號    INTEGER,
    國際音標韻母 TEXT,
    台語音標韻母 TEXT,
    台羅韻母   TEXT,
    白話字韻母  TEXT,
    閩拼韻母   TEXT,
    方音韻母   TEXT,
    十五音韻母  TEXT,
    十五音舒促聲 TEXT,
    十五音序   INTEGER
);
```

## 韻攝清單

```sql
CREATE TABLE 韻攝清單 (
    韻目列號 INTEGER NOT NULL
                 UNIQUE,
    韻攝   TEXT,
    四聲韻目 TEXT,
    PRIMARY KEY (
        韻目列號 AUTOINCREMENT
    )
);
```

## 韻母方陣表

```sql
CREATE TABLE 韻目方陣表 (
    識別   INTEGER NOT NULL
                 UNIQUE,
    韻目   TEXT,
    韻目列號 INTEGER,
    韻攝   TEXT,
    廣韻聲調 TEXT,
    目次序號 INTEGER,
    PRIMARY KEY (
        識別 AUTOINCREMENT
    )
);
```

## 聲調表

```sql
CREATE TABLE 聲調表 (
    識別號   INTEGER NOT NULL,
    台羅八聲調 INTEGER,
    調值    TEXT,
    方音符號  TEXT,
    四聲調   TEXT,
    雅俗通聲調 TEXT,
    舒促聲   TEXT,
    PRIMARY KEY (
        識別號 AUTOINCREMENT
    )
);
```


```sql
CREATE TABLE 韻母對照表 (
    識別號    INTEGER,
    國際音標韻母 TEXT,
    台語音標韻母 TEXT,
    台羅韻母   TEXT,
    白話字韻母  TEXT,
    閩拼韻母   TEXT,
    方音韻母   TEXT,
    十五音韻母  TEXT,
    十五音舒促聲 TEXT,
    十五音序   INTEGER
);
```
