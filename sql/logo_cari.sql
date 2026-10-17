/* Logo Tiger / Go - Cari kart listesi (sadece okur)
   Firma numarasi 001 degilse tablo adlarindaki 001'i degistirin.
   Sonucu Excel'e "cari_listesi.xlsx" olarak kaydedin; kolon adlari
   DIA_Veri_Temizleme'nin ESLESME kurallarina uygundur. */
SELECT
  C.CODE        AS [Cari Kodu],
  C.DEFINITION_ AS [Unvan],
  C.TAXNR       AS [Vergi No],
  C.TAXOFFICE   AS [Vergi Dairesi],
  C.TCKNO       AS [TC Kimlik],      -- eski surumlerde yoksa bu satiri silin
  C.ADDR1       AS [Adres 1],
  C.ADDR2       AS [Adres 2],
  C.CITY        AS [Il],
  C.TOWN        AS [Ilce],
  C.POSTCODE    AS [postakodu],
  C.COUNTRY     AS [Ulke],
  C.TELNRS1     AS [Telefon 1],
  C.TELNRS2     AS [Telefon 2],
  C.CELLPHONE   AS [Cep Telefonu],
  C.FAXNR       AS [Fax],
  C.EMAILADDR   AS [Email],
  C.WEBADDR     AS [Web Adresi],
  C.SPECODE     AS [Ozel Kod 1],
  C.CARDTYPE    AS [LOGO_KART_TURU]  -- 1 Alici, 2 Satici, 3 Alici+Satici
FROM LG_001_CLCARD C
WHERE C.ACTIVE = 0          -- sadece aktif kartlar
ORDER BY C.CODE;
