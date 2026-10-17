/* =====================================================================
   SEMA KESIF - ETA / Logo / Netsis veritabaninin gercek tablo ve alanlarini listeler
   Kullanim (SQL Server Management Studio):
     1) Ust menuden ilgili sirket veritabanini secin (USE ...)
     2) F5 ile calistirin
     3) Sonuc ekraninda sag tik > "Save Results As..." ile CSV kaydedin
   Sadece OKUR, hicbir veriyi degistirmez.
   ===================================================================== */

/* 1) Hangi program? Tablo adlarina gore tahmin */
SELECT
  CASE
    WHEN EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME LIKE 'LG[_]___[_]CLCARD') THEN 'LOGO'
    WHEN EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'TBLCASABIT')          THEN 'NETSIS'
    WHEN EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'CARKART')             THEN 'ETA'
    ELSE 'BILINMIYOR'
  END AS PROGRAM,
  DB_NAME() AS VERITABANI;

/* 2) Cari / stok / muhasebe / banka / personel kart tablolarinin tum alanlari */
SELECT c.TABLE_NAME, c.ORDINAL_POSITION, c.COLUMN_NAME, c.DATA_TYPE,
       c.CHARACTER_MAXIMUM_LENGTH AS UZUNLUK, c.IS_NULLABLE
FROM INFORMATION_SCHEMA.COLUMNS c
WHERE
  /* Logo */
     c.TABLE_NAME LIKE 'LG[_]___[_]CLCARD'   OR c.TABLE_NAME LIKE 'LG[_]___[_]ITEMS'
  OR c.TABLE_NAME LIKE 'LG[_]___[_]UNITSET%' OR c.TABLE_NAME LIKE 'LG[_]___[_]ITMUNITA'
  OR c.TABLE_NAME LIKE 'LG[_]___[_]UNITBARCODE' OR c.TABLE_NAME LIKE 'LG[_]___[_]PRCLIST'
  OR c.TABLE_NAME LIKE 'LG[_]___[_]EMUHACC'  OR c.TABLE_NAME LIKE 'LG[_]___[_]BNCARD'
  OR c.TABLE_NAME LIKE 'LG[_]___[_]BANKACC'  OR c.TABLE_NAME LIKE 'LG[_]___[_]SRVCARD'
  OR c.TABLE_NAME LIKE 'LG[_]___[_]SHIPINFO'
  /* Netsis */
  OR c.TABLE_NAME IN ('TBLCASABIT','TBLCASABITEK','TBLCARIEMAIL','TBLSTSABIT','TBLSTSABITEK',
                      'TBLMUPLAN','TBLBNKHESSABIT','TBLSTOKBARKOD','TBLSTOKFIAT')
  OR c.TABLE_NAME LIKE 'TBL%PERSONEL%' OR c.TABLE_NAME LIKE 'TBL%DEMIRBAS%' OR c.TABLE_NAME LIKE 'TBL%BARKOD%'
  /* ETA */
  OR c.TABLE_NAME IN ('CARKART','ADRESLER','STKKART','STKBARKOD','STKFIYAT','STKFIYTIP','DEPSTKKART')
  OR c.TABLE_NAME LIKE 'MUH%KART%' OR c.TABLE_NAME LIKE 'BNK%KART%' OR c.TABLE_NAME LIKE 'PER%KART%'
  OR c.TABLE_NAME LIKE 'DMR%KART%'
ORDER BY c.TABLE_NAME, c.ORDINAL_POSITION;

/* 3) Tablo listesi ve satir sayilari (hangi tablolar dolu?) */
SELECT t.name AS TABLO, SUM(p.rows) AS SATIR
FROM sys.tables t
JOIN sys.partitions p ON p.object_id = t.object_id AND p.index_id IN (0, 1)
GROUP BY t.name
HAVING SUM(p.rows) > 0
ORDER BY t.name;
