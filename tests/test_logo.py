"""Logo kaynagi testleri: Logo tablo yapisini taklit eden SQLite veritabani ile.
Son test Logo -> temizleme -> DIA (sahte sunucu) akisini uctan uca calistirir."""
import copy
import os
import sqlite3
import sys
import tempfile
import unittest

import pandas as pd

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))
sys.path.insert(0, os.path.dirname(__file__))

from logo_kaynak import LogoKaynak, firma_no_dogrula, excele_yaz  # noqa: E402


def logo_db(tckno=True):
    c = sqlite3.connect(':memory:')
    cari_kolon = ('LOGICALREF, ACTIVE, CARDTYPE, CODE, DEFINITION_, SPECODE, ADDR1, ADDR2, CITY, TOWN, '
                  'POSTCODE, COUNTRY, TELNRS1, TELNRS2, FAXNR, TAXNR, TAXOFFICE, EMAILADDR, WEBADDR, '
                  'CELLPHONE' + (', TCKNO, ISPERSCOMP' if tckno else ''))
    c.execute('CREATE TABLE LG_001_CLCARD (' + cari_kolon + ')')
    satirlar = [
        (1, 0, 1, 'C001  ', 'ABC GIDA SAN. TİC. A.Ş.', 'TOPTAN', 'Moda Cad. 1', '', 'Kadıköy/İstanbul', '',
         '34710', 'TÜRKİYE', '0532 123 45 67', '', '', '0123456789', 'KADIKÖY', 'info@abc.com', '', '',
         '', 0),
        (2, 0, 3, 'C002', 'Ali Veli', '', '', '', 'Siverek', '', '', '', '5321234568', '', '', '', '', '',
         '', '05339998877', '10000000146', 1),
        (3, 1, 1, 'C003', 'Pasif Kart', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 0),
    ]
    if not tckno:
        satirlar = [s[:-2] for s in satirlar]
    c.executemany('INSERT INTO LG_001_CLCARD VALUES (' + ','.join('?' * len(satirlar[0])) + ')', satirlar)
    c.execute('CREATE TABLE LG_001_ITEMS (LOGICALREF, ACTIVE, CARDTYPE, CODE, NAME, STGRPCODE, SPECODE, '
              'VAT, UNITSETREF)')
    c.executemany('INSERT INTO LG_001_ITEMS VALUES (?,?,?,?,?,?,?,?,?)', [
        (10, 0, 1, 'S001', 'Çay 1 kg', 'GIDA', 'A', 20, 5),
        (11, 0, 20, 'SINIF', 'Malzeme sınıfı', '', '', 0, 5),
    ])
    c.execute('CREATE TABLE LG_001_UNITSETL (LOGICALREF, CODE, UNITSETREF, MAINUNIT)')
    c.executemany('INSERT INTO LG_001_UNITSETL VALUES (?,?,?,?)', [(1, 'ADET', 5, 1), (2, 'KOLI', 5, 0)])
    c.execute('CREATE TABLE LG_001_ITMUNITA (ITEMREF, LINENR, BARCODE)')
    c.execute("INSERT INTO LG_001_ITMUNITA VALUES (10, 1, '8690000000017')")
    c.execute('CREATE TABLE LG_001_PRCLIST (CARDREF, PTYPE, PRIORITY, PRICE)')
    c.executemany('INSERT INTO LG_001_PRCLIST VALUES (?,?,?,?)', [(10, 2, 1, 150.5), (10, 1, 1, 100)])
    return LogoKaynak(c, 'sqlite')


class LogoTest(unittest.TestCase):
    def test_firma_no(self):
        self.assertEqual(firma_no_dogrula('1'), '001')
        with self.assertRaises(ValueError):
            firma_no_dogrula("1; DROP TABLE x")

    def test_firmalar(self):
        self.assertEqual(logo_db().firmalar(), ['001'])

    def test_cari(self):
        df = logo_db().cari_cek(1)
        self.assertEqual(df['Cari Kodu'].tolist(), ['C001', 'C002'])       # pasif kart yok, bosluk kirpildi
        self.assertEqual(df['Vergi No'][0], '0123456789')
        self.assertEqual(df['TC Kimlik'][1], '10000000146')

    def test_eski_surum_tckno_yok(self):
        df = logo_db(tckno=False).cari_cek(1)
        self.assertNotIn('TC Kimlik', df.columns)
        self.assertEqual(len(df), 2)

    def test_stok(self):
        df = logo_db().stok_cek('001')
        self.assertEqual(df['Stok Kodu'].tolist(), ['S001'])
        self.assertEqual(df['Ana Birim'][0], 'ADET')
        self.assertEqual(df['Barkod No'][0], '8690000000017')
        self.assertEqual(df['Satis Fiyati'][0], 150.5)

    def test_olmayan_firma(self):
        with self.assertRaises(RuntimeError):
            logo_db().cari_sorgusu(2)


class UcTanUcaTest(unittest.TestCase):
    def test_logo_temizle_dia(self):
        import dia_isle
        from dia_aktar import aktar, VARSAYILAN_ESLESME
        from test_dia import SahteDia, istemci
        with tempfile.TemporaryDirectory() as d:
            excele_yaz(logo_db().cari_cek(1), d, 'cari_listesi.xlsx')
            dia_isle.KLASOR = d
            dia_isle.HAFIZA_KLASORU = d
            dia_isle.modul_isle('cari', dia_isle.kurallari_yukle(), mukerrer_mod='otomatik')
            hazir = pd.read_excel(os.path.join(d, 'cari_dia_hazir.xlsx'), dtype=str)
        self.assertEqual(hazir['verginumarasi'][0], '0123456789')
        self.assertEqual((hazir['sehir'][0], hazir['ilce'][0]), ('İSTANBUL', 'KADIKÖY'))
        self.assertEqual(hazir['sehir'][1], 'ŞANLIURFA')
        sahte = SahteDia()
        rapor = aktar(istemci(sahte), hazir, 'cari', copy.deepcopy(VARSAYILAN_ESLESME), log=lambda m: None)
        self.assertEqual(rapor['sonuc'].tolist(), ['OK', 'OK'])
        kart = sahte.istek('scf_carikart_ekle')[0]['kart']
        self.assertEqual(kart['verginumarasi'], '0123456789')
        self.assertEqual(sahte.istek('scf_carikart_adresleri_ekle')[0]['kart']['_key_sis_sehirler'], 34)


if __name__ == '__main__':
    unittest.main()
