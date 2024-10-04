# 羅馬拼音

## 需求

### 功能摘要

- 功能名稱：TL_Tng_Zu_Im (台羅轉注音)

- 功能描述：將傳入之【台羅拼音】羅馬字母轉換成【方音符號】之注音符號。

- 範例說明：

    - 漢字：不
    - 羅馬拼音：put4
    - 聲母：p
    - 韻母：ut
    - 聲調：4

    ```python
	def TL_Tng_Zu_Im(siann_bu, un_bu, siann_tiau):
        # 處理作業
        return {
            '聲母': zu_im_siann_bu,
            '韻母': zu_im_un_bu,
            '聲調': zu_im_siann_tiau,
        }



    zu_im_fu_ho = TL_Tng_Zu_Im(siann_bu='p', un_bu='ut', siann_tiau=4)
    # 進行斷言
	assert zu_im_fu_ho['聲母'] == 'ㄅ', "聲母不正確"
	assert zu_im_fu_ho['韻母'] == 'ㄨㆵ', "韻母不正確"
	assert zu_im_fu_ho['聲調'] == '', "聲調不正確"
    ```


### 資料表結構

#### 【韻母表】

```sh
CREATE TABLE 韻母表 (
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

#### 【聲母表】

```sh
CREATE TABLE 聲母表 (
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

#### 【聲調表】

```sh
CREATE TABLE 聲調表 (
    識別號   INTEGER NOT NULL,
    台羅八聲調 INTEGER,
    方音符號  TEXT,
    四聲調   TEXT,
    雅俗通聲調 TEXT,
    舒促聲   TEXT,
    PRIMARY KEY (
        識別號 AUTOINCREMENT
    )
);
```


## 參考

### 變更資料表

```sh
PRAGMA foreign_keys = 0;

CREATE TABLE sqlitestudio_temp_table AS SELECT *
                                          FROM 聲調表;

DROP TABLE 聲調表;

CREATE TABLE 聲調表 (
    識別號   INTEGER NOT NULL,
    台羅八聲調 INTEGER,
    方音符號  TEXT,
    四聲調   TEXT,
    雅俗通聲調 TEXT,
    舒促聲   TEXT,
    PRIMARY KEY (
        識別號 AUTOINCREMENT
    )
);

INSERT INTO 聲調表 (
                    識別號,
                    台羅八聲調,
                    四聲調,
                    雅俗通聲調,
                    舒促聲
                )
                SELECT 識別號,
                       台羅八聲調,
                       四聲調,
                       聲調,
                       舒促聲
                  FROM sqlitestudio_temp_table;

DROP TABLE sqlitestudio_temp_table;

PRAGMA foreign_keys = 1;
```
