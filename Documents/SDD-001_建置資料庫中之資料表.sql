CREATE TABLE 韻母對照表 (
    識別號    INTEGER PRIMARY KEY,
    韻母碼    TEXT,
    韻母國際音標 TEXT,
    白話字韻母  TEXT,
    閩拼韻母   TEXT,
    台羅韻母   TEXT,
    方音韻母   TEXT,
    十五音韻母  TEXT,
    舒促聲    TEXT,
    十五音序   INTEGER
);

CREATE TABLE 廣韻韻母對照表 (
    識別號    INTEGER PRIMARY KEY,
    韻母識別號  INTEGER REFERENCES 韻母對照表 (識別號),
    廣韻韻母   TEXT,
    雅俗通韻母  TEXT,
    舒促聲    TEXT,
    韻母拼音碼  TEXT,
    韻母國際音標 TEXT,
    林進三拚音碼 TEXT
);

CREATE TABLE 切語下字表 (
    識別號     INTEGER PRIMARY KEY,
    廣韻韻母識別號 INTEGER REFERENCES 廣韻韻母對照表 (識別號),
    韻系列號    INTEGER,
    韻系行號    INTEGER,
    韻目索引    TEXT,
    目次識別號   INTEGER,
    目次      TEXT,
    攝       TEXT,
    韻系      TEXT,
    韻目      TEXT,
    調       TEXT,
    呼       TEXT,
    等       INTEGER,
    韻母      TEXT,
    切語下字集   TEXT,
    等呼      TEXT,
    韻母拼音碼   TEXT,
    備註      TEXT
);

CREATE TABLE 聲母對照表 (
    識別號    INTEGER PRIMARY KEY,
    聲母碼    TEXT,
    聲母國際音標 TEXT,
    白話字聲母  TEXT,
    閩拼聲母   TEXT,
    台羅聲母   TEXT,
    方音聲母   TEXT,
    十五音聲母  TEXT
);

CREATE TABLE 廣韻聲母對照表 (
    識別號    INTEGER PRIMARY KEY,
    聲母識別號  INTEGER REFERENCES 聲母對照表 (識別號),
    廣韻聲母   TEXT,
    雅俗通聲母  TEXT,
    聲母拼音碼  TEXT,
    聲母國際音標 TEXT
);

CREATE TABLE 切語上字表 (
    識別號     INTEGER PRIMARY KEY,
    廣韻聲母識別號 INTEGER REFERENCES 廣韻聲母對照表 (識別號),
    七聲類     TEXT,
    發音部位    TEXT,
    聲母      TEXT,
    清濁      TEXT,
    發送收     TEXT,
    聲母拼音碼   TEXT,
    切語上字集   TEXT,
    備註      TEXT
);

CREATE TABLE 小韻表 (
    識別號    INTEGER PRIMARY KEY,
    上字表識別號 INTEGER REFERENCES 切語上字表 (識別號),
    下字表識別號 INTEGER REFERENCES 切語下字表 (識別號),
    切語     TEXT,
    拼音     TEXT,
    小韻字    TEXT,
    目次編碼   TEXT,
    小韻字序號  INTEGER,
    小韻字集   TEXT,
    字數     INTEGER,
    聲母     TEXT,
    聲母拼音碼  TEXT,
    發音部位   TEXT,
    清濁     TEXT,
    發送收    TEXT,
    韻母     TEXT,
    韻母拼音碼  TEXT,
    調      TEXT,
    調號     INTEGER,
    備註     TEXT,
    原有備註   TEXT,
    異體字    TEXT,
    其它備註   TEXT
);

CREATE TABLE 字表 (
    識別號   INTEGER PRIMARY KEY,
    字     TEXT,
    同音字序  INTEGER,
    切語    TEXT,
    谷歌小韻號 INTEGER,
    小韻識別號 INTEGER REFERENCES 小韻表 (識別號),
    拼音    TEXT,
    字義    TEXT,
    備註    TEXT
);

---------------------------------------------------------------------


CREATE TABLE 字表 (
    識別號   INTEGER PRIMARY KEY,
    字     TEXT,
    同音字序  INTEGER,
    切語    TEXT,
    谷歌小韻號 INTEGER,
    小韻識別號 INTEGER REFERENCES 小韻表 (識別號) ON DELETE CASCADE,
    拼音    TEXT,
    字義    TEXT,
    備註    TEXT
)