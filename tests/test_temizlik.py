"""Calistirma:  python -m unittest discover -s tests"""
import os
import sys
import tempfile
import unittest

import pandas as pd

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

from _ortak import (telefon_isle, vkn_isle, tc_isle, vkn_gecerli, tc_gecerli, buyuk_harf,  # noqa: E402
                    cari_kayitturu_ata, ayar, ayar_cevir, kurallari_excele_yaz,
                    kurallari_excelden_oku)
from il_ilce import il_ilce_coz, il_ilce_duzelt  # noqa: E402
from il_ilce_veri import IL_ILCE  # noqa: E402


class IlIlceTest(unittest.TestCase):
    def test_veri(self):
        self.assertEqual(len(IL_ILCE), 81)
        self.assertEqual(sum(len(v) for v in IL_ILCE.values()), 973)

    def test_dogru_giris(self):
        self.assertEqual(il_ilce_coz('İstanbul', 'Kadıköy'), ('İSTANBUL', 'KADIKÖY', 'OK'))

    def test_yer_degistirmis(self):
        self.assertEqual(il_ilce_coz('Kadıköy', 'İstanbul')[:2], ('İSTANBUL', 'KADIKÖY'))

    def test_ayni_hucre(self):
        for il, ilce in [('Kadıköy/İstanbul', ''), ('', 'Kadıköy / İstanbul'),
                         ('İstanbul - Kadıköy', None), ('İstanbul Kadıköy', '')]:
            self.assertEqual(il_ilce_coz(il, ilce)[:2], ('İSTANBUL', 'KADIKÖY'), (il, ilce))

    def test_il_doldur(self):
        self.assertEqual(il_ilce_coz('', 'Siverek')[:2], ('ŞANLIURFA', 'SİVEREK'))
        self.assertEqual(il_ilce_coz('', 'Siverek', il_bos_ise_doldur=False)[0], '')

    def test_belirsiz_ilce(self):
        il, ilce, durum = il_ilce_coz('', 'Gölbaşı')
        self.assertEqual((il, ilce), ('', 'GÖLBAŞI'))
        self.assertTrue(durum.startswith('UYARI'))

    def test_takma_ad_ve_ascii(self):
        self.assertEqual(il_ilce_coz('Afyon', 'merkez')[:2], ('AFYONKARAHİSAR', 'MERKEZ'))
        self.assertEqual(il_ilce_coz('ISTANBUL', 'kadikoy')[:2], ('İSTANBUL', 'KADIKÖY'))

    def test_uyumsuz(self):
        self.assertIn('ILINDE YOK', il_ilce_coz('Ankara', 'Kadıköy')[2])

    def test_dataframe_baslik(self):
        df = pd.DataFrame({'sehir': ['Kadıköy'], 'ilce': ['İstanbul']})
        df, rapor = il_ilce_duzelt(df, yazim='BASLIK')
        self.assertEqual(df.iloc[0].tolist(), ['İstanbul', 'Kadıköy'])
        self.assertEqual(len(rapor), 1)


class TelefonTest(unittest.TestCase):
    def test_bastaki_sifir_silinir(self):
        self.assertEqual(telefon_isle('0532 123 45 67')[0], '5321234567')
        self.assertEqual(telefon_isle('+90 532 123 45 67')[0], '5321234567')
        self.assertEqual(telefon_isle(5321234567.0), ('5321234567', 'OK'))

    def test_formatlar(self):
        self.assertEqual(telefon_isle('5321234567', '11')[0], '05321234567')
        self.assertEqual(telefon_isle('5321234567', '90')[0], '905321234567')
        self.assertEqual(telefon_isle('5321234567', '+90')[0], '+905321234567')

    def test_dahili(self):
        tel, durum = telefon_isle('0212 555 44 33 dahili 12')
        self.assertEqual(tel, '2125554433')
        self.assertTrue(durum.startswith('UYARI'))

    def test_gecersiz(self):
        self.assertTrue(telefon_isle('123')[1].startswith('GEÇERSİZ'))


class VknTcTest(unittest.TestCase):
    def test_sifir_tamamla(self):
        self.assertEqual(vkn_isle(123456789)[0], '0123456789')
        self.assertEqual(vkn_isle('123456789.0')[0], '0123456789')
        self.assertEqual(vkn_isle('0123456789'), ('0123456789', 'OK'))
        self.assertEqual(vkn_isle(123456789, sifir_tamamla=False, algoritma_kontrol=False)[0],
                         '123456789')

    def test_kontrol_hanesi(self):
        self.assertTrue(vkn_gecerli('1234567890'))
        self.assertFalse(vkn_gecerli('1234567891'))
        self.assertTrue(vkn_isle('1234567891')[1].startswith('UYARI'))

    def test_tc(self):
        self.assertTrue(tc_gecerli('10000000146'))
        self.assertFalse(tc_gecerli('12345678901'))
        self.assertEqual(tc_isle(10000000146.0), ('10000000146', 'OK'))


class DigerTest(unittest.TestCase):
    def test_turkce_buyuk_harf(self):
        self.assertEqual(buyuk_harf('izmir ışık'), 'İZMİR IŞIK')

    def test_cari_kayitturu(self):
        unvan = ['KASAP AHMET', 'DENİZ AŞ', 'XYZ LTD.ŞTİ', 'Hasan Taş']
        df = pd.DataFrame({'unvan': unvan, 'verginumarasi': [''] * 4, 'tckimlikno2': [''] * 4})
        self.assertEqual(cari_kayitturu_ata(df)['cari kayitturu'].tolist(),
                         ['SHS', 'KRLS', 'KRLS', 'SHS'])

    def test_ayar(self):
        self.assertEqual(ayar({}, 'telefon_format'), '10')
        self.assertEqual(ayar_cevir('telefon_format', 11.0), '11')
        self.assertFalse(ayar_cevir('vkn_sifir_tamamla', 'Hayir'))

    def test_kurallar_excel_gidis_donus(self):
        from kurallar import KURALLAR
        with tempfile.TemporaryDirectory() as d:
            yol = os.path.join(d, 'kurallar.xlsx')
            kurallari_excele_yaz(KURALLAR, yol)
            K = kurallari_excelden_oku(yol)
        self.assertEqual(K['cari']['eslesme'], KURALLAR['cari']['eslesme'])
        self.assertEqual(K['cari']['telefon_alanlari'], KURALLAR['cari']['telefon_alanlari'])
        self.assertEqual(K['cari']['telefon_format'], '10')
        self.assertIs(K['cari']['il_ilce_duzelt'], True)


class UcTanUcaTest(unittest.TestCase):
    def test_cari(self):
        import dia_isle
        with tempfile.TemporaryDirectory() as d:
            pd.DataFrame({
                'Cari Kodu': ['C1', 'C2', 'C3'],
                'Unvan': ['ABC A.Ş.', 'Ali Veli', 'Ayşe'],
                'Vergi No': [123456789, None, None],
                'Il': ['Kadıköy/İstanbul', 'Siverek', 'Kadıköy'],
                'Ilce': [None, None, 'İstanbul'],
                'Telefon': ['0532 123 45 67', 5321234568, None],
                'Ulke': ['', None, 'TURKIYE'],
            }).to_excel(os.path.join(d, 'cari_listesi.xlsx'), index=False)
            dia_isle.KLASOR = d
            dia_isle.HAFIZA_KLASORU = d
            dia_isle.modul_isle('cari', dia_isle.kurallari_yukle(), mukerrer_mod='otomatik')
            df = pd.read_excel(os.path.join(d, 'cari_dia_hazir.xlsx'), dtype=str)
            fark = pd.read_excel(os.path.join(d, 'cari_dia_hazir_DEGISIKLIKLER.xlsx'), dtype=str)
        self.assertIn('0532 123 45 67', fark['Eski Değer'].tolist())
        self.assertEqual(df['verginumarasi'][0], '0123456789')
        self.assertEqual(df['sehir'].tolist(), ['İSTANBUL', 'ŞANLIURFA', 'İSTANBUL'])
        self.assertEqual(df['ilce'].tolist(), ['KADIKÖY', 'SİVEREK', 'KADIKÖY'])
        self.assertEqual(df['ulke'].tolist(), ['TURKIYE'] * 3)      # varsayilan bos hucreye yazildi
        # cep numarasi telefon1'den ceptel'e tasinir (cep_ayir)
        self.assertEqual(df['ceptel'][0], '5321234567')
        self.assertTrue(pd.isna(df['telefon1'][0]))
        self.assertEqual(df['cari kayitturu'][0], 'KRLS')


if __name__ == '__main__':
    unittest.main()
