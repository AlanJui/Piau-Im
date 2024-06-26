DROP VIEW  IF EXISTS 小韻速檢;
CREATE VIEW 小韻速檢 AS SELECT
           小韻表.識別號 AS 小韻識別號,
		   小韻表.小韻字,
           小韻表.切語 AS 小韻切語,
           小韻表.拼音 AS 小韻標音,
           小韻表.目次編碼 AS 小韻目次,
 
           切語上字檢視.廣韻聲母,
           切語上字檢視.聲母碼,
           切語上字檢視.七聲類,
           切語上字檢視.發音部位,
           切語上字檢視.清濁,
           切語上字檢視.發送收,

           切語下字檢視.廣韻韻母,
           切語下字檢視.韻母碼,
           切語下字檢視.攝,
           切語下字檢視.韻系,
           切語下字檢視.韻目,
           切語下字檢視.調,
           切語下字檢視.呼,
           切語下字檢視.等,
           切語下字檢視.等呼
 
      FROM 小韻表
           LEFT JOIN 切語上字檢視 ON 小韻表.上字表識別號 = 切語上字檢視.識別號
           LEFT JOIN 切語下字檢視 ON 小韻表.下字表識別號 = 切語下字檢視.識別號;