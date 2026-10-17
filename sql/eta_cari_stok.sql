/* ETA:SQL / ETA V8 SQL - Cari ve stok (sadece okur)
   ETA'nin alan adlari kamuya acik dokumanda cok az; sadece dogrulanmis alanlar yazildi.
   Once 01_sema_kesif.sql calistirin, CARKART / STKKART / ADRESLER alanlarini gonderin;
   sorgu vergi no, il/ilce, telefon alanlariyla tamamlanacak. */

-- Cari
SELECT
  K.CARKOD    AS [Cari Kodu],
  K.CARUNVAN  AS [Unvan]
FROM CARKART K
ORDER BY K.CARKOD;

-- Stok
SELECT
  S.STKKOD    AS [Stok Kodu],
  S.STKCINSI  AS [Stok Adi],
  S.STKBIRIM  AS [Ana Birim],     -- (?) dogrulanmali
  (SELECT TOP 1 B.STKBARKOD FROM STKBARKOD B WHERE B.STKBARSTKKOD = S.STKKOD) AS [Barkod No]
FROM STKKART S
ORDER BY S.STKKOD;
