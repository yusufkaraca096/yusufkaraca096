"""DIA istemcisi ve aktarim testleri (sahte DIA sunucusu ile; internet gerekmez)."""
import copy
import json
import os
import sys
import tempfile
import unittest

import pandas as pd

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

from dia_ws import DiaIstemci, DiaHatasi, yeni_key, sunucu_kodu_temizle  # noqa: E402
from dia_aktar import aktar, kart_olustur, eslesme_yukle, VARSAYILAN_ESLESME  # noqa: E402


class SahteDia:
    """Gelen istekleri kaydeder, DIA benzeri cevap uretir."""

    def __init__(self, mevcut_cariler=None, oturum_dusur=0, kontor=10 ** 9, yetkisiz=()):
        self.yetkisiz = set(yetkisiz)
        self.istekler = []
        self.mevcut = dict(mevcut_cariler or {})   # kod -> _key
        self.oturum_dusur = oturum_dusur
        self.kontor = kontor
        self.sonraki_key = 1000
        self.session = None

    def __call__(self, url, govde):
        self.istekler.append((url, copy.deepcopy(govde)))
        servis, p = next(iter(govde.items()))
        if servis == 'login':
            if p['password'] != 'dogru':
                return {'code': '401', 'msg': 'Kullanici adi veya sifre hatali'}
            self.session = 'S' + str(len(self.istekler))
            return {'code': '200', 'msg': self.session}
        if servis == 'logout':
            return {'code': '200', 'msg': ''}
        if p.get('session_id') != self.session:
            return {'code': '419', 'msg': 'Invalid session'}
        if self.oturum_dusur:
            self.oturum_dusur -= 1
            self.session = None
            return {'code': '419', 'msg': 'LOGIN_TIMEOUT'}
        if servis in self.yetkisiz:
            return {'code': '401', 'msg': 'INSUFFICIENT_PRIVILEGES -> Web Servis kullanıcınızın yetkisini ayarlayınız!'}
        if servis == 'sis_kontor_sorgula':
            return {'code': '200', 'result': self.kontor}
        if servis.endswith('_listele'):
            if servis == 'sis_sehirler_listele':
                rows = [{'_key': 34, 'sehiradi': 'İSTANBUL'}, {'_key': 6, 'sehiradi': 'ANKARA'}]
            elif servis == 'scf_carikart_listele':
                rows = [{'_key': k, 'carikartkodu': c} for c, k in self.mevcut.items()]
            else:
                rows = []
            return {'code': '200', 'result': rows[p['offset']:p['offset'] + p['limit']]}
        self.kontor -= 1
        if self.kontor < 0:
            return {'code': '406', 'msg': 'CREDIT_ERROR'}
        if servis.endswith('_ekle'):
            if not p['kart'].get('unvan') and servis == 'scf_carikart_ekle':
                return {'code': '400', 'msg': 'unvan zorunlu'}
            self.sonraki_key += 1
            return {'code': '200', 'key': self.sonraki_key}
        if servis.endswith('_guncelle'):
            return {'code': '200', 'msg': 'ok'}
        return {'code': '404', 'msg': 'bilinmeyen servis ' + servis}

    def servisler(self):
        return [next(iter(g)) for _, g in self.istekler]

    def istek(self, servis):
        return [next(iter(g.values())) for _, g in self.istekler if next(iter(g)) == servis]


def istemci(sahte, sifre='dogru'):
    return DiaIstemci('test', 'kullanici', sifre, 'APIKEY', 1, 1,
                      temel_url='https://test.ws.dia.com.tr/api/v3', gonderici=sahte)


def cari_df():
    return pd.DataFrame({
        'carikartkodu': ['C1', 'C2', 'C3'],
        'unvan': ['ABC A.Ş.', 'Ali Veli', None],
        'verginumarasi': ['0123456789', None, None],
        'tckimlikno2': [None, '10000000146', None],
        'sehir': ['İSTANBUL', 'ANKARA', 'BİLİNMEYEN'],
        'ilce': ['KADIKÖY', 'ÇANKAYA', None],
        'adres1': ['Moda Cad. 1', None, 'x'],
        'telefon1': ['5321234567', None, None],
        'cari kayitturu': ['KRLS', 'SHS', 'SHS'],
    })


class IstemciTest(unittest.TestCase):
    def test_login_ve_url(self):
        s = SahteDia()
        d = istemci(s)
        d.login()
        url, govde = s.istekler[0]
        self.assertEqual(url, 'https://test.ws.dia.com.tr/api/v3/sis/json')
        self.assertEqual(govde['login']['params']['apikey'], 'APIKEY')
        self.assertEqual(d.session_id, s.session)

    def test_hatali_sifre(self):
        with self.assertRaises(DiaHatasi) as h:
            istemci(SahteDia(), sifre='yanlis').login()
        self.assertEqual(h.exception.code, '401')

    def test_oturum_duserse_yeniden_login(self):
        s = SahteDia(oturum_dusur=1)
        d = istemci(s)
        self.assertEqual(d.kontor_sorgula(), 10 ** 9)
        self.assertEqual(s.servisler().count('login'), 2)

    def test_listele_sayfalama(self):
        s = SahteDia({'C' + str(i): i for i in range(25)})
        rows = istemci(s).listele('scf', 'scf_carikart_listele', sayfa=10)
        self.assertEqual(len(rows), 25)
        self.assertEqual(len(s.istek('scf_carikart_listele')), 3)

    def test_sunucu_kodu_her_bicimde(self):
        for g in ['racabilisim', ' RacaBilisim ', 'racabilisim.ws.dia.com.tr',
                  'https://racabilisim.ws.dia.com.tr/api/v3/', 'https://racabilisim.ws.dia.com.tr/api/v3/scf/json',
                  'racabilisim.ws.dia.com.tr:443']:
            self.assertEqual(sunucu_kodu_temizle(g), 'racabilisim', g)
        d = DiaIstemci('https://racabilisim.ws.dia.com.tr/api/v3/', 'u', 'p', 'k', 1, 1)
        self.assertEqual(d.temel_url, 'https://racabilisim.ws.dia.com.tr/api/v3')

    def test_yetki_hatasi_yeniden_login_yapmaz(self):
        from dia_ws import DiaYetkiHatasi
        s = SahteDia(yetkisiz={'scf_carikart_listele'})
        with self.assertRaises(DiaYetkiHatasi):
            istemci(s).listele('scf', 'scf_carikart_listele')
        self.assertEqual(s.servisler().count('login'), 1)

    def test_yetki_hatasi_aktarimi_durdurur(self):
        s = SahteDia(yetkisiz={'scf_carikart_ekle'})
        rapor = aktar(istemci(s), cari_df(), 'cari', copy.deepcopy(VARSAYILAN_ESLESME), log=lambda m: None)
        self.assertEqual(len(rapor), 1)
        self.assertTrue(rapor['mesaj'][0].startswith('YETKI'))

    def test_yeni_key(self):
        self.assertEqual(yeni_key({'key': 5}), 5)
        self.assertEqual(yeni_key({'msg': '7'}), 7)
        self.assertEqual(yeni_key({'result': {'_key': 9}}), 9)


class AktarimTest(unittest.TestCase):
    def test_kart_olustur_bos_alan_gondermez(self):
        kart = kart_olustur(cari_df().iloc[1], VARSAYILAN_ESLESME['cari']['kart'], {'carikarttipi': 'AL'})
        self.assertEqual(kart, {'carikartkodu': 'C2', 'unvan': 'Ali Veli',
                                'tckimlikno': '10000000146', 'carikarttipi': 'AL'})

    def test_ekle_guncelle_ve_adres(self):
        s = SahteDia({'C2': 55})
        rapor = aktar(istemci(s), cari_df(), 'cari', copy.deepcopy(VARSAYILAN_ESLESME), log=lambda m: None)
        self.assertEqual(rapor['islem'].tolist(), ['EKLE', 'GUNCELLE', 'EKLE'])
        self.assertEqual(rapor['sonuc'].tolist(), ['OK', 'OK', 'HATA'])
        self.assertIn('unvan zorunlu', rapor['mesaj'][2])
        guncel = s.istek('scf_carikart_guncelle')[0]['kart']
        self.assertEqual(guncel['_key'], 55)
        self.assertNotIn('carikarttipi', guncel)       # sabit alanlar guncellemede ezilmez
        adres = s.istek('scf_carikart_adresleri_ekle')
        self.assertEqual(len(adres), 1)
        self.assertEqual(adres[0]['kart']['_key_scf_carikart'], 1001)
        self.assertEqual(adres[0]['kart']['_key_sis_sehirler'], 34)
        self.assertEqual(adres[0]['kart']['ilce'], 'KADIKÖY')

    def test_kuru_calistir_hicbir_sey_yazmaz(self):
        s = SahteDia()
        rapor = aktar(istemci(s), cari_df(), 'cari', copy.deepcopy(VARSAYILAN_ESLESME), kuru_calistir=True,
                      log=lambda m: None)
        self.assertEqual(set(rapor['sonuc']), {'KURU'})
        self.assertFalse([x for x in s.servisler() if x.endswith(('_ekle', '_guncelle'))])
        self.assertIn('0123456789', rapor['istek'][0])

    def test_kontor_bitince_durur(self):
        s = SahteDia(kontor=1)
        rapor = aktar(istemci(s), cari_df(), 'cari', copy.deepcopy(VARSAYILAN_ESLESME), log=lambda m: None)
        self.assertIn('KONTOR', rapor['mesaj'].iloc[-1])
        self.assertLess(len(rapor), 3)

    def test_mevcut_ise_atla(self):
        s = SahteDia({'C1': 1, 'C2': 2, 'C3': 3})
        es = copy.deepcopy(VARSAYILAN_ESLESME)
        es['cari']['mevcut_ise'] = 'atla'
        rapor = aktar(istemci(s), cari_df(), 'cari', es, log=lambda m: None)
        self.assertEqual(set(rapor['islem']), {'ATLA'})

    def test_eksik_dosya_mesaji(self):
        from dia_aktar import eksik_dosya_mesaji
        modul = {'cikti_dosya': 'cari_dia_hazir.xlsx', 'girdi_dosya': 'cari_listesi.xlsx'}
        with tempfile.TemporaryDirectory() as d:
            self.assertIn('Logo', eksik_dosya_mesaji(d, modul))
            open(os.path.join(d, 'cari_listesi.xlsx'), 'w').close()
            self.assertIn('henuz temizlenmemis', eksik_dosya_mesaji(d, modul))
            open(os.path.join(d, 'cari_dia_hazir_INCELEME.xlsx'), 'w').close()
            self.assertIn('Inceleme', eksik_dosya_mesaji(d, modul))

    def test_eslesme_dosyasi(self):
        with tempfile.TemporaryDirectory() as d:
            es = eslesme_yukle(d)
            self.assertTrue(os.path.exists(os.path.join(d, 'dia_eslesme.json')))
            with open(os.path.join(d, 'dia_eslesme.json'), encoding='utf-8') as f:
                self.assertEqual(json.load(f), es)


if __name__ == '__main__':
    unittest.main()
