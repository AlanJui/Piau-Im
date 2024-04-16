-- 變更小韻表：要求識別號改成 INTEGER ，並能自動加一

PRAGMA foreign_keys = 0;

CREATE TABLE sqlitestudio_temp_table AS SELECT *
                                          FROM 小韻表;

DROP TABLE 小韻表;

CREATE TABLE 小韻表 (
    識別號    INTEGER PRIMARY KEY ASC AUTOINCREMENT
                   UNIQUE,
    上字表識別號         REFERENCES 切語上字表 (識別號) ON DELETE CASCADE
                                          ON UPDATE CASCADE,
    下字表識別號         REFERENCES 切語下字表 (識別號) ON DELETE CASCADE
                                          ON UPDATE CASCADE,
    切語,
    拼音,
    小韻字,
    目次編碼,
    小韻字序號,
    小韻字集,
    字數,
    聲母,
    聲母拼音碼,
    發音部位,
    清濁,
    發送收,
    韻母,
    韻母拼音碼,
    調,
    調號,
    備註,
    原有備註,
    異體字,
    其它備註,
    ""
);

INSERT INTO 小韻表 (
                    識別號,
                    上字表識別號,
                    下字表識別號,
                    切語,
                    拼音,
                    小韻字,
                    目次編碼,
                    小韻字序號,
                    小韻字集,
                    字數,
                    聲母,
                    聲母拼音碼,
                    發音部位,
                    清濁,
                    發送收,
                    韻母,
                    韻母拼音碼,
                    調,
                    調號,
                    備註,
                    原有備註,
                    異體字,
                    其它備註,
                    ""
                )
                SELECT 識別號,
                       上字表識別號,
                       下字表識別號,
                       切語,
                       拼音,
                       小韻字,
                       目次編碼,
                       小韻字序號,
                       小韻字集,
                       字數,
                       聲母,
                       聲母拼音碼,
                       發音部位,
                       清濁,
                       發送收,
                       韻母,
                       韻母拼音碼,
                       調,
                       調號,
                       備註,
                       原有備註,
                       異體字,
                       其它備註,
                       ""
                  FROM sqlitestudio_temp_table;

DROP TABLE sqlitestudio_temp_table;

PRAGMA foreign_keys = 1;
