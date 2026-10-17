/* Netsis - Stok kart listesi (sadece okur). "stok_listesi.xlsx" olarak kaydedin.
   (?) isaretli alanlari 01_sema_kesif.sql ciktisina gore acin. */
SELECT
  S.STOK_KODU    AS [Stok Kodu],
  S.STOK_ADI     AS [Stok Adi],
  S.OLCU_BR1     AS [Ana Birim],
  S.GRUP_KODU    AS [Grup Kodu],
  S.KOD_1        AS [Ozel Kod 1],
  S.SATIS_FIAT1  AS [Satis Fiyati]
  -- , S.KDV_ORANI   AS [KDV]        -- (?)
  -- , S.ALIS_FIAT1  AS [Alis Fiyati] -- (?)
FROM TBLSTSABIT S
ORDER BY S.STOK_KODU;
