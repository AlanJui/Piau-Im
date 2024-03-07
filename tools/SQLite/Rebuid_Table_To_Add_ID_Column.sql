-- 創建一個新的表格，其結構與原表格相同，但多了一個帶有 AUTOINCREMENT 屬性的 ID 列
CREATE TABLE "Lui_Tsip_Nga_Siok_Thong_new" (
    "ID" INTEGER PRIMARY KEY AUTOINCREMENT,
    "Ji" TEXT NOT NULL,
    "Siann" TEXT NOT NULL,
    "Un" TEXT NOT NULL,
    "Tiau" TEXT NOT NULL,
    "Phing_Im" TEXT NOT NULL
);

-- 將原表格的數據複製到新表格中
INSERT INTO "Lui_Tsip_Nga_Siok_Thong_new" ("Ji", "Siann", "Un", "Tiau", "Phing_Im")
SELECT "Ji", "Siann", "Un", "Tiau", "Phing_Im" FROM "Lui_Tsip_Nga_Siok_Thong";

-- 刪除原表格
DROP TABLE "Lui_Tsip_Nga_Siok_Thong";

-- 將新表格重命名為原表格的名稱
ALTER TABLE "Lui_Tsip_Nga_Siok_Thong_new" RENAME TO "Lui_Tsip_Nga_Siok_Thong";