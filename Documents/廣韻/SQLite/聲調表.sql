--
-- SQLiteStudio v3.4.4 生成的檔案，週四 十月 24 15:48:16 2024
--
-- 所用的文字編碼：System
--
PRAGMA foreign_keys = off;
BEGIN TRANSACTION;

-- 表：聲調表
DROP TABLE IF EXISTS 聲調表;

CREATE TABLE IF NOT EXISTS 聲調表 (
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

INSERT INTO 聲調表 (識別號, 台羅八聲調, 調值, 方音符號, 四聲調, 雅俗通聲調, 舒促聲) VALUES (1, 1, '44', 'ˉ', '平', '上平', '舒');
INSERT INTO 聲調表 (識別號, 台羅八聲調, 調值, 方音符號, 四聲調, 雅俗通聲調, 舒促聲) VALUES (2, 2, '51', 'ˋ', '上', '上上', '舒');
INSERT INTO 聲調表 (識別號, 台羅八聲調, 調值, 方音符號, 四聲調, 雅俗通聲調, 舒促聲) VALUES (3, 3, '31', '˪', '去', '上去', '舒');
INSERT INTO 聲調表 (識別號, 台羅八聲調, 調值, 方音符號, 四聲調, 雅俗通聲調, 舒促聲) VALUES (4, 4, '30', '', '入', '上入', '促');
INSERT INTO 聲調表 (識別號, 台羅八聲調, 調值, 方音符號, 四聲調, 雅俗通聲調, 舒促聲) VALUES (5, 5, '24', 'ˊ', '平', '下平', '舒');
INSERT INTO 聲調表 (識別號, 台羅八聲調, 調值, 方音符號, 四聲調, 雅俗通聲調, 舒促聲) VALUES (6, 6, '00', 'ˋ', '上', '下上', '舒');
INSERT INTO 聲調表 (識別號, 台羅八聲調, 調值, 方音符號, 四聲調, 雅俗通聲調, 舒促聲) VALUES (7, 7, '33', '˫', '去', '下去', '舒');
INSERT INTO 聲調表 (識別號, 台羅八聲調, 調值, 方音符號, 四聲調, 雅俗通聲調, 舒促聲) VALUES (8, 8, '50', '˙', '入', '下入', '促');
INSERT INTO 聲調表 (識別號, 台羅八聲調, 調值, 方音符號, 四聲調, 雅俗通聲調, 舒促聲) VALUES (9, 0, NULL, '˙', '', NULL, NULL);

COMMIT TRANSACTION;
PRAGMA foreign_keys = on;
