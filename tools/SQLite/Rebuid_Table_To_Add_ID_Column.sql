CREATE TABLE "Han_Ji_Phing_Im_Ji_Tian_new" (
    "ID" INTEGER PRIMARY KEY AUTOINCREMENT,
    "Han_Ji" TEXT NOT NULL,
    "TL_Phing_Im" TEXT NOT NULL,
    "freq" REAL,
    "NST_ID" TEXT,
    "Siann" TEXT,
    "Un" TEXT,
    "Tiau" TEXT
);

-- 將原表格的數據複製到新表格中
INSERT INTO "Han_Ji_Phing_Im_Ji_Tian_new" ("Han_Ji", "TL_Phing_Im", "freq", "NST_ID", "Siann", "Un", "Tiau")
SELECT "Han_Ji", "TL_Phing_Im", "freq", "NST_ID", "Siann", "Un", "Tiau" FROM "Han_Ji_Phing_Im_Ji_Tian";

-- 刪除原表格
DROP TABLE "Han_Ji_Phing_Im_Ji_Tian";

-- 將新表格重命名為原表格的名稱
ALTER TABLE "Han_Ji_Phing_Im_Ji_Tian_new" RENAME TO "Han_Ji_Phing_Im_Ji_Tian";