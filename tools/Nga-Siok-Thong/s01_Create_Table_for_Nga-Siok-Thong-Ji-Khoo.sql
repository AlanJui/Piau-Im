--建立 "雅俗通字庫" 表格
CREATE TABLE 雅俗通字庫 (
    識別號  INTEGER  PRIMARY KEY AUTOINCREMENT UNIQUE NOT NULL,
    漢字    TEXT    NOT NULL,
    切音    TEXT    NOT NULL,
    字韻    TEXT    NOT NULL,
    聲調    INTEGER NOT NULL,
    原始拼音  TEXT,
    舒促聲   TEXT    NOT NULL,
    聲     TEXT    NOT NULL,
    韻     TEXT    NOT NULL,
    調     INTEGER NOT NULL,
    拼音碼   TEXT    NOT NULL,
    雅俗通標音 TEXT    NOT NULL,
    十五音標音 TEXT    NOT NULL,
    常用度   REAL    DEFAULT (0.0) 
);
