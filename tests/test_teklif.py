"""Teklif motoru testleri (uydurma fiyatlarla; gercek fiyat listesi depoda yoktur)."""
import os
import sqlite3
import sys
import tempfile
import unittest
from unittest import mock

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

import teklif_cikti  # noqa: E402
from teklif_hesap import hesapla, ybh_hesapla, kademe_katsayi  # noqa: E402
from teklif_veri import TeklifVeri, AYAR_VARSAYILAN, varsayilan_klasor, veri_klasoru_degistir  # noqa: E402


def motor_excel(yol):
    """RACA_DIA_Teklif_Motoru.xlsx ile ayni yapida kucuk dosya."""
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = 'PaketListesi'
    ws.append(['PAKET LİSTESİ', 'Kodu', 'Adı', 'Std. Fiyat', 'Ent. Fiyat', 'Kullanıcı %', 'YBH %', 'Açıklama',
               'Ürünler'])
    ws.append(['Masaüstü', 'MAS1', 'Muhasebe', 1000, 2000, 10, 15, 'açıklama', 'a-b-c'])
    ws.append(['Masaüstü', 'MAS2', 'Ticari', 500, 0, 20, 15, '', 'a-b'])
    ws.append(['Mobil', 'MOB1', 'Mobil Satış', 300, 400, 50, 10, '', 'm'])
    ws.append(['Diğer', 'DIG1', 'E-Fatura', 800, 900, 0, 20, '', ''])
    a = wb.create_sheet('ADAM SAAT')
    a.append(['ADAM/SAAT'])
    a.append([])
    a.append([])
    a.append(['Paket Kodu', 'Modül', 'Standart', 'Enterprise'])
    for r, (std, ent) in enumerate([(2, 3), (1, None), (1, 2), (4, 5)], start=5):
        a.cell(r, 1, '=IF(PaketListesi!B%d="","",PaketListesi!B%d)' % (r - 3, r - 3))   # Excel'deki gibi formul
        a.cell(r, 3, std)
        a.cell(r, 4, '=ROUND(C%d*1.5,0)' % r if ent is None else ent)
    s = wb.create_sheet('AYARLAR')
    s.append(['AYARLAR'])
    s.append([])
    s.append(['Ayar', 'Değer', 'Açıklama'])
    for satir in [('Enterprise Masaüstü Eşiği', 10, ''), ('Enterprise Mobil Eşiği', 10, ''),
                  ('Lisans Yüzdesi (%)', 10, ''), ('Saatlik Danışman Ücreti', 100, ''),
                  ('Minimum Uyarlama', 500, 'Alt sınır'), ('Maksimum Uyarlama', 0, ''),
                  ('--- SABİT FİYATLAR ---', None, None), ('B2B Premium Uyarlama', 700, 'Sabit fiyat'),
                  ('KDV Oranı (%)', 20, '')]:
        s.append(list(satir))
    k = wb.create_sheet('KAMPANYALAR')
    k.append(['KAMPANYA'])
    k.append([])
    k.append(['Kampanya Adı', 'İndirim %', 'Durum', 'Açıklama'])
    k.append(['YOK', 0, 'AKTİF', ''])
    k.append(['Rakip Geçiş', 15, 'AKTİF', ''])
    k.append(['Eski', 30, 'PASİF', ''])
    r = wb.create_sheet('RAKİP')
    for _ in range(3):
        r.append([])
    r.append(['Firma', 'Adam/Saat', 'Saatlik TL', 'Toplam'])
    r.append(['RACA (Biz)', 8, 100])
    r.append(['Logo', 8, 150])
    r.append(['Mikro', 8, 50])
    g = wb.create_sheet('TEKLİF GEÇMİŞİ')
    for _ in range(3):
        g.append([])
    g.append(['Teklif No', 'Tarih', 'Müşteri', 'Paket', 'Kullanıcı', 'Lisans', 'Uyarlama', 'KDV Dahil', 'Kampanya',
              'Sonuç', 'Red Nedeni'])
    g.append(['ESKI-1', '2025-05-01', 'Eski Müşteri', 'STANDART', '3', 1000, 500, 1800, 'YOK', 'KAZANDI', ''])
    wb.save(yol)


class TeklifTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.d = self.tmp.name
        motor_excel(os.path.join(self.d, 'motor.xlsx'))
        self.v = TeklifVeri(os.path.join(self.d, 'veri', 'teklif.db'))
        self.rapor = self.v.excelden_ice_aktar(os.path.join(self.d, 'motor.xlsx'))
        self.ayar = self.v.ayarlar()
        self.pk = self.v.paketler()

    def tearDown(self):
        self.v.db.close()
        self.tmp.cleanup()

    def h(self, **g):
        return hesapla(dict({'masaustu': 3, 'mobil': 2}, **g), self.pk, self.ayar, self.v.sabit_kalemler(),
                       self.v.rakipler())

    # ---------------------------------------------------------------- ice aktarma
    def test_ice_aktarma(self):
        self.assertEqual(self.rapor[0], '4 paket')
        self.assertEqual((self.pk['MAS1']['std_saat'], self.pk['MAS1']['ent_saat']), (2, 3))
        self.assertEqual(self.pk['MAS2']['ent_saat'], 2)        # =ROUND(C*1.5) formulu
        self.assertEqual(self.ayar['saatlik_ucret'], 100)
        self.assertEqual(self.v.sabit_kalemler(), {'B2B Premium Uyarlama': 700})
        self.assertEqual([k['ad'] for k in self.v.kampanyalar()], ['YOK', 'Rakip Geçiş'])   # PASIF gelmez
        self.assertEqual([r['firma'] for r in self.v.rakipler()], ['Logo', 'Mikro'])
        self.assertEqual(self.v.teklifler()[0]['no'], 'ESKI-1')

    # ---------------------------------------------------------------- hesap
    def test_kullanici_ve_paket_tipi(self):
        s = self.h(kalemler=[{'kod': 'MAS1'}, {'kod': 'MOB1'}, {'kod': 'DIG1'}])
        self.assertEqual(s['paket_tipi'], 'STANDART')
        mas, mob, dig = s['satirlar']
        self.assertEqual((mas['adet'], mas['toplam']), (3, 1000 + 2 * 1000 * 0.10))
        self.assertEqual((mob['adet'], mob['toplam']), (2, 300 + 300 * 0.5))
        self.assertEqual((dig['adet'], dig['toplam']), (1, 800))
        s = self.h(masaustu=11, kalemler=[{'kod': 'MAS1'}, {'kod': 'MAS2'}])
        self.assertEqual(s['paket_tipi'], 'ENTERPRISE')
        self.assertEqual(s['satirlar'][0]['birim'], 2000)
        self.assertEqual(s['satirlar'][1]['birim'], 500)          # Ent. fiyat 0 -> Std.
        self.assertTrue(any('Enterprise fiyatı yok' in u for u in s['uyarilar']))

    def test_indirim_tutari_kdv_haric_ve_kdvyi_etkiler(self):
        """Kullanici sikayeti: Ind. Tutar KDV haric toplami ve KDV alt toplamini etkilemiyordu."""
        s = self.h(kalemler=[{'kod': 'DIG1', 'ind_yuzde': 10, 'ind_tutar': 100}])
        self.assertEqual(s['lisans'], {'net': 620.0, 'kdv': 124.0, 'dahil': 744.0, 'toplam': 800.0,
                                       'indirim': 180.0})
        bos = self.h(kalemler=[{'kod': 'DIG1'}])
        self.assertEqual(s['uyarlama']['sinir'], 'ALT SINIR')       # uyarlama iki durumda da 500
        self.assertEqual(round(bos['net_toplam'] - s['net_toplam'], 2), 180)
        self.assertEqual(round(bos['kdv'] - s['kdv'], 2), 36)
        self.assertEqual(round(bos['genel_toplam'] - s['genel_toplam'], 2), 216)

    def test_uyarlama_tum_satirlarin_saati(self):
        """Excel'de SUM(O17:O24) ilk 7 satiri saymiyordu."""
        s = self.h(kalemler=[{'kod': k} for k in ('MAS1', 'MAS2', 'MOB1', 'DIG1')])
        self.assertEqual(s['uyarlama']['saat'], 2 + 1 + 1 + 4)
        self.assertEqual(s['uyarlama']['saat_bedeli'], 800)
        self.assertEqual(s['uyarlama']['net'], round(s['lisans']['net'] * 0.10 + 800, 2))
        s = self.h(kalemler=[{'kod': 'MAS1', 'saat': 0}])     # elle saat
        self.assertEqual(s['uyarlama']['saat'], 0)
        self.assertEqual((s['uyarlama']['net'], s['uyarlama']['sinir']), (500, 'ALT SINIR'))

    def test_kampanya_ve_odeme_plani_net_uzerinden(self):
        """Excel'de odeme plani net yerine kampanya INDIRIM tutarindan hesaplaniyordu."""
        s = self.h(kalemler=[{'kod': 'DIG1'}], sabitler=['B2B Premium Uyarlama'], kampanya='Rakip Geçiş',
                   kampanya_yuzde=15, maliyet=1000)
        ara = s['ara_toplam']
        self.assertEqual(ara, 800 + 500 + 700)                 # lisans + min uyarlama + sabit
        self.assertEqual(s['net_toplam'], ara * 0.85)
        self.assertEqual(s['genel_toplam'], round(ara * 0.85 * 1.2, 2))
        pesin, iki = s['odeme'][0], s['odeme'][1]
        self.assertEqual(pesin['dahil'], round(ara * 0.85 * 0.95 * 1.2, 2))
        self.assertEqual(iki['dahil'], round(ara * 0.85 / 2 * 1.2, 2))
        self.assertEqual(s['kar']['brut'], ara * 0.85 - 1000)

    def test_rakip(self):
        s = self.h(kalemler=[{'kod': 'DIG1'}])                  # 4 saat x 100 = 400
        r = s['rakip']
        self.assertEqual(r['biz'], 400)
        self.assertEqual([x['toplam'] for x in r['satirlar']], [600, 200])     # ayni 4 saat x rakip ucreti
        self.assertEqual(r['ortalama'], 400)
        self.ayar['rakip_ayni_saat'] = 0
        self.assertEqual([x['toplam'] for x in self.h(kalemler=[{'kod': 'DIG1'}])['rakip']['satirlar']], [1200, 400])

    def test_bilinmeyen_kod_sirayi_kaydirmaz(self):
        s = self.h(kalemler=[{'kod': 'YOK'}, {'kod': 'DIG1'}])
        self.assertEqual([x['sira'] for x in s['satirlar']], [1])
        self.assertTrue(s['uyarilar'])

    def test_ybh(self):
        self.assertEqual(kademe_katsayi(5, '5:1; 10:1.1; *:1.65'), 1)
        self.assertEqual(kademe_katsayi(7, '5:1; 10:1.1; *:1.65'), 1.1)
        self.assertEqual(kademe_katsayi(70, '5:1; 10:1.1; *:1.65'), 1.65)
        y = ybh_hesapla('STANDART', [{'kod': 'MAS1', 'kullanici': 3}, {'kod': 'MAS2', 'kullanici': 1}],
                        self.pk, self.ayar)
        self.assertEqual(y['modul'], 5)
        self.assertEqual(y['baz'], round(1200 * 0.15 + 500 * 0.15, 2))
        self.assertEqual(y['yillik']['net'], round(y['baz'] * 1 * 1, 2))

    # ---------------------------------------------------------------- kayit / yedek
    def test_teklif_kaydet_ve_numara(self):
        g = {'musteri': 'A', 'masaustu': 1, 'mobil': 0, 'kalemler': [{'kod': 'MAS1'}]}
        s = hesapla(g, self.pk, self.ayar)
        no1 = self.v.teklif_kaydet(g, s)
        no2 = self.v.teklif_kaydet(g, s)
        self.assertTrue(no1.endswith('-0001') and no2.endswith('-0002'))
        self.v.teklif_sonuc_yaz(no1, 'KAZANDI')
        self.v.teklif_kaydet(dict(g, musteri='A2'), s, no1)          # duzenleme sonucu silmez
        t = {x['no']: x for x in self.v.teklifler()}
        self.assertEqual((t[no1]['musteri'], t[no1]['sonuc']), ('A2', 'KAZANDI'))
        self.assertEqual(self.v.teklif_girdisi(no1)['kalemler'], [{'kod': 'MAS1'}])

    def test_kayitlar_kapatip_acinca_kalir_ve_yedeklenir(self):
        """Kullanici sikayeti: kapatip acinca onceki kayitlar siliniyordu."""
        yol = self.v.yol
        g = {'musteri': 'Kalıcı', 'masaustu': 1, 'mobil': 0, 'kalemler': [{'kod': 'MAS1'}]}
        no = self.v.teklif_kaydet(g, hesapla(g, self.pk, self.ayar))
        self.v.kapat()
        self.v = TeklifVeri(yol)
        self.assertIn(no, [t['no'] for t in self.v.teklifler()])
        son = self.v.yedekler()[0][0]
        self.assertIn('kapanis', son)                       # kaydedilen teklif kapanis yedeginde de var
        y = sqlite3.connect(son)
        self.assertEqual(y.execute('SELECT musteri FROM teklif WHERE no=?', (no,)).fetchone()[0], 'Kalıcı')
        y.close()
        # yedekten geri donus: silinen teklif geri gelir
        self.v.yedek_al('test')
        self.v.teklif_sil(no)
        self.v.yedekten_don(self.v.yedekler()[0][0] if 'test' in self.v.yedekler()[0][1] else
                            [y for y, a in self.v.yedekler() if 'test' in a][0])
        self.assertIn(no, [t['no'] for t in self.v.teklifler()])

    def test_yedek_sayisi_sinirli(self):
        import teklif_veri
        with mock.patch.object(teklif_veri, 'YEDEK_SAYISI', 3):
            for n in range(5):
                with mock.patch('teklif_veri.datetime') as dt:
                    from datetime import datetime
                    dt.now.return_value = datetime(2026, 1, 1, 10, 0, n)
                    self.v.yedek_al('n%d' % n)
        self.assertEqual(len(self.v.yedekler()), 3)

    def test_gecersiz_yedek_reddedilir(self):
        bozuk = os.path.join(self.d, 'bozuk.db')
        sqlite3.connect(bozuk).close()
        with self.assertRaises(ValueError):
            self.v.yedekten_don(bozuk)

    def test_veri_exe_yaninda_tutulmaz(self):
        with mock.patch.dict(os.environ, {'APPDATA': os.path.join(self.d, 'appdata')}), \
                mock.patch('os.path.expanduser', return_value=os.path.join(self.d, 'ev')):
            k = varsayilan_klasor()
            self.assertTrue(k.startswith(os.path.join(self.d, 'ev')))
            self.assertNotIn('src', os.path.relpath(k, self.d))
            os.makedirs(k)
            open(os.path.join(k, 'teklif.db'), 'w').write('x')
            hedef = veri_klasoru_degistir(os.path.join(self.d, 'ortak'))
            self.assertTrue(os.path.exists(hedef))
            self.assertEqual(varsayilan_klasor(), os.path.join(self.d, 'ortak'))

    def test_fiyat_listesi_guncelleme_saatleri_korur(self):
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.append(['DIA fiyat listesi'])
        ws.append(['Tür', 'Kodu', 'Adı', 'Std. Fiyat', 'Ent. Fiyat', 'Kullanıcı %', 'YBH %', 'Açıklama', 'Ürünler'])
        ws.append(['Masaüstü', 'MAS1', 'Muhasebe', 1100, 2200, 10, 15, '', 'a-b-c'])
        ws.append(['Diğer', 'YENI1', 'Yeni', 50, 0, 0, 0, '', ''])
        yol = os.path.join(self.d, 'dia.xlsx')
        wb.save(yol)
        r = self.v.paket_listesi_guncelle(yol)
        self.assertEqual((r['fiyati_degisen'], r['yeni']), (['MAS1'], ['YENI1']))
        self.assertEqual(sorted(r['listeden_kalkan']), ['DIG1', 'MAS2', 'MOB1'])
        pk = self.v.paketler()
        self.assertEqual((pk['MAS1']['std_fiyat'], pk['MAS1']['std_saat']), (1100, 2))
        self.assertEqual(pk['MAS2']['aktif'], 0)
        self.assertEqual(len(self.v.paketler(sadece_aktif=True)), 2)

    def test_ayar_varsayilanlari_ticari_deger_icermez(self):
        for ad in ('saatlik_ucret', 'min_uyarlama', 'max_uyarlama'):
            self.assertEqual(AYAR_VARSAYILAN[ad], 0)

    # ---------------------------------------------------------------- cikti
    def test_pdf_ve_excel(self):
        g = {'teklif_no': 'RT-2026-0001', 'musteri': 'Çağlar Şeker İnş. Ltd. Şti.', 'masaustu': 3, 'mobil': 2,
             'kalemler': [{'kod': 'MAS1', 'ind_tutar': 50}, {'kod': 'DIG1'}], 'kampanya': 'Rakip Geçiş',
             'kampanya_yuzde': 15, 'notlar': 'Özel not'}
        s = hesapla(g, self.pk, self.ayar, self.v.sabit_kalemler())
        pdf = teklif_cikti.pdf_olustur(os.path.join(self.d, 't.pdf'), g, s, self.ayar)
        with open(pdf, 'rb') as f:
            self.assertTrue(f.read(5).startswith(b'%PDF'))
        xl = teklif_cikti.excel_olustur(os.path.join(self.d, 't.xlsx'), g, s, self.ayar)
        from openpyxl import load_workbook
        hucreler = [c.value for r in load_workbook(xl).active.iter_rows() for c in r if c.value is not None]
        self.assertIn(s['genel_toplam'], hucreler)
        self.assertIn('Çağlar Şeker İnş. Ltd. Şti.', hucreler)
        self.assertFalse([h for h in hucreler if isinstance(h, str) and 'Maliyet' in h])   # ic bilgi yok
        self.assertEqual(teklif_cikti.para(1234567.8), '1.234.567,80 TL')


if __name__ == '__main__':
    unittest.main()
