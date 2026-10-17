/* Logo Tiger / Go - Stok (malzeme) kart listesi (sadece okur)
   Firma 001 degilse 001'leri degistirin. "stok_listesi.xlsx" olarak kaydedin. */
SELECT
  I.CODE       AS [Stok Kodu],
  I.NAME       AS [Stok Adi],
  U.CODE       AS [Ana Birim],
  I.VAT        AS [KDV],
  I.STGRPCODE  AS [Grup Kodu],
  I.SPECODE    AS [Ozel Kod 1],
  (SELECT TOP 1 B.BARCODE FROM LG_001_ITMUNITA B
     WHERE B.ITEMREF = I.LOGICALREF AND B.BARCODE <> '' ORDER BY B.LINENR) AS [Barkod No],
  (SELECT TOP 1 P.PRICE FROM LG_001_PRCLIST P
     WHERE P.CARDREF = I.LOGICALREF AND P.PTYPE = 2 ORDER BY P.PRIORITY) AS [Satis Fiyati],
  (SELECT TOP 1 P.PRICE FROM LG_001_PRCLIST P
     WHERE P.CARDREF = I.LOGICALREF AND P.PTYPE = 1 ORDER BY P.PRIORITY) AS [Alis Fiyati],
  I.CARDTYPE   AS [LOGO_KART_TURU]   -- 1 TM, 10 HM, 11 YM, 12 MM, 13 TK ...
FROM LG_001_ITEMS I
LEFT JOIN LG_001_UNITSETL U ON U.UNITSETREF = I.UNITSETREF AND U.MAINUNIT = 1
WHERE I.ACTIVE = 0
  AND I.CARDTYPE NOT IN (20, 21)     -- malzeme siniflari haric
ORDER BY I.CODE;
