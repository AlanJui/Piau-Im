-- 先將現有的字表更名為臨時表
ALTER TABLE 字表 RENAME TO 字表_tmp;

-- 建立一個新的字表，並包含所有的欄位以及新的常用率欄位的預設值為 0.0
CREATE TABLE 字表 (
    識別號   INTEGER PRIMARY KEY ASC AUTOINCREMENT
                  UNIQUE
                  NOT NULL,
    小韻識別號 INTEGER REFERENCES 小韻表 (識別號) ON DELETE NO ACTION
                                       ON UPDATE NO ACTION,
    同音字序,
    字,
    切語,
    拼音,
    常用率   REAL    DEFAULT (0.0),  -- 修改預設值為 0.0
    字義,
    備註,
    谷歌小韻號
);

-- 從臨時表中將資料插入新的字表中
INSERT INTO 字表 SELECT * FROM 字表_tmp;

-- 刪除臨時表
DROP TABLE 字表_tmp;
