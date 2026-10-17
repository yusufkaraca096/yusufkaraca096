/* Netsis - Cari kart listesi (sadece okur)
   Ilgili sirket veritabaninda calistirin. "cari_listesi.xlsx" olarak kaydedin.
   (?) isaretli alanlar her surumde olmayabilir: once 01_sema_kesif.sql ile kontrol edin. */
SELECT
  C.CARI_KOD        AS [Cari Kodu],
  C.CARI_ISIM       AS [Unvan],
  C.VERGI_NUMARASI  AS [Vergi No],
  C.VERGI_DAIRESI   AS [Vergi Dairesi],
  C.CARI_ADRES      AS [Adres 1],
  C.CARI_IL         AS [Il],
  C.CARI_ILCE       AS [Ilce],
  C.POSTAKODU       AS [postakodu],
  C.CARI_TEL        AS [Telefon 1],
  C.FAX             AS [Fax],
  C.GRUP_KODU       AS [Grup Kodu],
  C.CARI_TIP        AS [NETSIS_CARI_TIP]
  -- , C.EMAIL       AS [Email]        -- (?) bazi surumlerde TBLCARIEMAIL tablosunda
  -- , E.TCKIMLIKNO  AS [TC Kimlik]    -- (?) TBLCASABIT veya TBLCASABITEK'te olabilir
FROM TBLCASABIT C
-- LEFT JOIN TBLCASABITEK E ON E.CARI_KOD = C.CARI_KOD
ORDER BY C.CARI_KOD;
