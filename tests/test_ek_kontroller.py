"""Ek kontrol testleri.  Calistirma:  python -m unittest discover -s tests"""
import os
import sys
import tempfile
import unittest

import pandas as pd

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

import ek_kontroller as ek  # noqa: E402
from _ortak import ayar, modul_ayar_tanimlari, kurallari_excele_yaz, kurallari_excelden_oku, KAYNAK_SATIR  # noqa: E402


def uyarilar(rapor, alan=None):
    return [r['durum'] for r in rapor if alan is None or r['alan'] == alan]


class OrtakTest(unittest.TestCase):
    def test_hucre_on_temizlik(self):
        df = pd.DataFrame({'a': ['#YOK', 'ABC_x000D_ LTD', 'Ä°STANBUL', '-', 'x​y', 'ÝZMÝR'], 'b': [0, 1, 2, 3, 4, 5]})
        df = ek.hucre_on_temizlik(df, {}, 'cari')
        self.assertTrue(pd.isna(df['a'][0]) and pd.isna(df['a'][3]))
        self.assertEqual(df['a'][1:3].tolist(), ['ABC  LTD', 'İSTANBUL'])
        self.assertEqual(df['a'][4:].tolist(), ['xy', 'İZMİR'])
        self.assertEqual(df['b'][0], 0)          # sayi 0 bosaltilmaz

    def test_ayar_kapatilinca_calismaz(self):
        df = pd.DataFrame({'a': ['#YOK']})
        self.assertEqual(ek.hucre_on_temizlik(df, {'hata_degerleri_temizle': False}, 'cari')['a'][0], '#YOK')

    def test_bilimsel(self):
        self.assertTrue(ek.bilimsel_isle('8,69E+12')[1].startswith('GEÇERSİZ'))
        self.assertEqual(ek.bilimsel_isle('8.691234567890E+12')[0], '8691234567890')
        self.assertEqual(ek.bilimsel_isle('ABC')[1], 'OK')

    def test_zorunlu_ve_uzunluk(self):
        df = pd.DataFrame({'carikartkodu': ['C1', '', 'X' * 30], 'unvan': ['A', 'B', None]})
        rapor = []
        ek.zorunlu_uzunluk_kontrol(df, {}, 'cari', rapor)
        self.assertEqual([r['satir'] for r in rapor if 'ZORUNLU' in r['durum']], [3, 4])
        self.assertEqual([r['satir'] for r in rapor if 'UZUNLUK' in r['durum']], [4])
        rapor = []
        ek.zorunlu_uzunluk_kontrol(df, {'zorunlu_alanlar': 'YOK', 'alan_uzunluklari': 'YOK'}, 'cari', rapor)
        self.assertEqual(rapor, [])

    def test_benzer_kayit(self):
        ciftler = ek.benzer_kayitlar(['ABC Gıda Ltd Şti', 'ABC GIDA LİMİTED ŞİRKETİ', 'XYZ Yapı', 'ABC Gida Ltd. Sti.',
                                     'ABC Gıda Ltd Şti'])
        self.assertIn((0, 1, 100), ciftler)
        self.assertNotIn((0, 4, 100), ciftler)        # birebir ayni: mukerrer kontrolunun isi
        self.assertFalse([c for c in ciftler if 2 in c[:2]])
        self.assertTrue(ek.benzer_kayitlar(['Mehmet Yılmaz', 'Mehmet Yilmazz'], 85))

    def test_degisiklikler(self):
        once = pd.DataFrame({'a': ['x', 'y', None], KAYNAK_SATIR: [2, 3, 4]})
        sonra = pd.DataFrame({'a': ['X', 'y', 'yeni'], KAYNAK_SATIR: [2, 3, None]})
        fark = ek.degisiklikler(once, sonra)
        self.assertEqual(fark['Yeni Değer'].tolist(), ['X', 'YENİ SATIR EKLENDİ', 'SATIR SİLİNDİ (boş/özet satır)'])

    def test_esleme_hafizasi(self):
        with tempfile.TemporaryDirectory() as d:
            self.assertEqual(ek.esleme_hafizasi_kaydet(d, 'cari', {'Firma Ünvanı': 'unvan', 'unvan': 'unvan'}), 1)
            self.assertEqual(ek.esleme_hafizasi_kaydet(d, 'cari', {'Firma Ünvanı': 'unvan'}), 0)
            self.assertEqual(ek.esleme_hafizasi_yukle(d), {'cari': {'Firma Ünvanı': 'unvan'}})

    def test_modul_bazli_ayarlar_excele_yazilir(self):
        from kurallar import KURALLAR
        with tempfile.TemporaryDirectory() as d:
            yol = os.path.join(d, 'kurallar.xlsx')
            kurallari_excele_yaz(KURALLAR, yol)
            K = kurallari_excelden_oku(yol)
        self.assertEqual(K['stok']['benzer_kayit_sutunu'], 'stok adı')
        self.assertIs(K['stok']['barkod_kontrol'], True)
        self.assertNotIn('barkod_kontrol', K['cari'])
        self.assertIn('unvan_standart', [a['ayar'] for a in modul_ayar_tanimlari('cari')])
        self.assertNotIn('unvan_standart', [a['ayar'] for a in modul_ayar_tanimlari('stok')])
        self.assertEqual(ayar({}, 'iban_sutunu', 'banka'), 'IBAN No')


class CariTest(unittest.TestCase):
    def test_unvan(self):
        c = ek.cift_ayari(ayar({}, 'unvan_kisaltmalari', 'cari'))
        self.assertEqual(ek.kisaltma_uygula('ABC Gıda Ltd.Şti', c), 'ABC Gıda LTD. ŞTİ.')
        self.assertEqual(ek.kisaltma_uygula('Deniz Sanayi ve Ticaret Limited Şirketi', c), 'Deniz SAN. ve TİC. LTD. ŞTİ.')
        self.assertEqual(ek.kisaltma_uygula('XYZ AŞ', c), 'XYZ A.Ş.')
        self.assertEqual(ek.kisaltma_uygula('AS YAPI', c), 'AS YAPI')      # sadece sonda
        self.assertEqual(ek.kisaltma_uygula('KASAP ALİ', c), 'KASAP ALİ')

    def test_vergi_dairesi(self):
        self.assertEqual(ek.vergi_dairesi_isle('Kadikoy V.D.', [])[0], 'KADIKÖY')
        self.assertEqual(ek.vergi_dairesi_isle('Kozyatağı Vergi Dairesi Müdürlüğü', [])[0], 'KOZYATAĞI')
        self.assertEqual(ek.vergi_dairesi_isle('buyuk mukellefler', [('BUYUK MUKELLEFLER', 'BÜYÜK MÜKELLEFLER')])[0],
                         'BÜYÜK MÜKELLEFLER')

    def test_adresten_il_ilce(self):
        self.assertEqual(ek.adresten_il_ilce('Moda Cad. No:5 Kadıköy/İstanbul'), ('İSTANBUL', 'KADIKÖY'))
        self.assertEqual(ek.adresten_il_ilce('Atatürk Mah. Ankara Cad. No:1'), ('', ''))
        self.assertEqual(ek.adresten_il_ilce('Bahçelievler Mah. Çankaya ANKARA'), ('ANKARA', 'ÇANKAYA'))

    def test_cari_akis(self):
        modul = {'telefon_alanlari': ['telefon1', 'telefon2', 'ceptel', 'fax']}
        df = pd.DataFrame({'unvan': ['ABC Ltd Şti', 'Ali Veli'], 'vergidairesi': ['Kadikoy VD', None],
                           'adres1': ['Moda Cad. Kadıköy İstanbul', None], 'sehir': [None, 'Ankara'],
                           'ilce': [None, None], 'telefon1': ['5321234567', '3122223344'],
                           'tckimlikno2': ['10000000146', None], 'postakodu': ['06100', '06100']})
        rapor = []
        df = ek.on_islem(df, modul, 'cari', rapor)
        df = ek.son_kontroller(df, modul, 'cari', rapor)
        self.assertEqual(df['unvan'][0], 'ABC LTD. ŞTİ.')
        self.assertEqual((df['sehir'][0], df['ilce'][0]), ('İSTANBUL', 'KADIKÖY'))
        self.assertEqual((df['ceptel'][0], df['telefon1'][0]), ('5321234567', ''))
        self.assertEqual(df['telefon1'][1], '3122223344')
        d = ' | '.join(uyarilar(rapor))
        self.assertIn('VKN YOK, SADECE TC', d)
        self.assertIn('POSTA KODU 06100 İSTANBUL', d)


class StokTest(unittest.TestCase):
    def test_stok_akis(self):
        modul = {'kdv_alanlari': ['KDV Satış Oran']}
        df = pd.DataFrame({'stok kart kodu': ['S1', 'S2', 'S3'], 'stok adı': ['Süt 1l', 'Peynir 500gr', 'Çay'],
                           'Ana Birim': ['Adt', 'kilo', 'Çuval'], 'Barkod No': ['4006381333931', '4006381333932', 'IC01'],
                           'KDV Satış Oran': [18, 1, 20], 'Fiyat 1': ['1.250,50 TL', '10 USD', 5.0],
                           'Fiyat 2': [2000, 5, 6]})
        rapor = []
        df = ek.on_islem(df, modul, 'stok', rapor)
        df = ek.son_kontroller(df, modul, 'stok', rapor)
        self.assertEqual(df['Ana Birim'].tolist(), ['ADET', 'KG', 'Çuval'])
        self.assertEqual(df['stok adı'].tolist(), ['Süt 1 LT', 'Peynir 500 GR', 'Çay'])
        self.assertEqual(df['Fiyat 1'].tolist(), [1250.5, 10.0, 5.0])
        self.assertEqual(df['Döviz 1'][1], 'USD')
        d = uyarilar(rapor)
        self.assertTrue(any('TANIMSIZ BİRİM' in x for x in d))
        self.assertEqual(len([x for x in d if 'BARKOD' in x]), 1)
        self.assertTrue(any('ESKİ KDV' in x for x in d))
        self.assertEqual(len([x for x in d if 'ALIŞ' in x]), 2)


class BankaPersonelTest(unittest.TestCase):
    def test_banka(self):
        df = pd.DataFrame({'IBAN No': ['TR120006400000112345678901', 'TR120006400000112345678901'],
                           'Banka adı': [None, 'Garanti'], 'Hesap NO': [None, '1'],
                           'Hesap Adı': ['İş Bankası USD', 'Garanti TL'], 'Döviz Türü': ['TL', 'TL']})
        rapor = []
        df = ek.on_islem(df, {}, 'banka', rapor)
        df = ek.son_kontroller(df, {}, 'banka', rapor)
        self.assertEqual(df['Banka adı'][0], 'İŞ BANKASI')
        self.assertEqual(df['Hesap NO'].tolist(), ['112345678901', '1'])
        d = ' | '.join(uyarilar(rapor))
        self.assertIn('IBAN İLE UYUŞMUYOR', d)
        self.assertIn('HESAP ADINDA USD VAR', d)

    def test_personel(self):
        df = pd.DataFrame({'Adı': ['Ali'], 'Soyadı': ['Veli'], 'Cinsiyet': ['Bay'], 'Kan Grubu': ['a rh pozitif'],
                           'Doğum Tarihi': ['01.01.2015'], 'İşe Başlama Tarihi': ['01.01.2024']})
        rapor = []
        df = ek.on_islem(df, {}, 'personel', rapor)
        df = ek.son_kontroller(df, {}, 'personel', rapor)
        self.assertEqual((df['Cinsiyet'][0], df['Kan Grubu'][0]), ('ERKEK', 'A RH+'))
        self.assertTrue(any('İŞE GİRİŞTE YAŞ' in x for x in uyarilar(rapor)))


class DemirbasHesapTest(unittest.TestCase):
    def test_demirbas(self):
        df = pd.DataFrame({'TARİH': ['01.01.2099', '01.01.2020'], 'TUTAR': [100, 0], 'AMORTİSMAN YIL': [5, 5],
                           'AMORTİSMAN ORANI': [20, 25], 'BİRİKMİŞ AMORTİSMAN': [150, None]})
        rapor = []
        ek.son_kontroller(df, {}, 'demirbas', rapor)
        d = ' | '.join(uyarilar(rapor))
        for parca in ('GELECEKTE', 'TUTAR 0', 'UYUMSUZ', 'TUTARDAN BÜYÜK'):
            self.assertIn(parca, d)
        self.assertEqual(len([x for x in uyarilar(rapor) if 'UYUMSUZ' in x]), 1)

    def test_hesap_plani(self):
        df = pd.DataFrame({'Hesap Kodu': ['120-01-001', '125.01', '12001002'], 'Hesap Adı': ['a', 'b', 'c']})
        rapor = []
        modul = {'hesap_seviye_uzunluklari': '3,2,3', 'eksik_ust_hesap_ekle': True}
        df = ek.on_islem(df, modul, 'hesapplan', rapor)
        df = ek.son_kontroller(df, modul, 'hesapplan', rapor)
        self.assertEqual(df['Hesap Kodu'].tolist()[:3], ['120.01.001', '125.01', '120.01.002'])
        self.assertEqual(sorted(df['Hesap Kodu'].tolist()[3:]), ['120', '120.01', '125'])
        self.assertTrue(any('125 TEK DÜZEN' in x for x in uyarilar(rapor)))


class CaprazTest(unittest.TestCase):
    def test_stok_grup_kodu_urun_agacinda(self):
        from kurallar import KURALLAR
        with tempfile.TemporaryDirectory() as d:
            pd.DataFrame({'Grup1': ['GIDA'], 'GRUP2': ['SUT']}).to_excel(
                os.path.join(d, 'urunagaci_listesi.xlsx'), index=False)
            df = pd.DataFrame({'Grup Kodu': ['SUT', 'ICECEK', None]})
            rapor = ek.capraz_kontrol(df, {}, 'stok', KURALLAR, d)
        self.assertEqual([r['satir'] for r in rapor], [3])


if __name__ == '__main__':
    unittest.main()
