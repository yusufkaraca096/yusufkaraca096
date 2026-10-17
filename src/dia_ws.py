"""DIA Web Servis (WS API v3) istemcisi.

Adres : https://SUNUCUKODU.ws.dia.com.tr/api/v3/<modul>/json   (POST, JSON)
Login : sis/json -> {"login": {"username", "password", "disconnect_same_user", "lang", "params": {"apikey"}}}
        cevap {"code": "200", "msg": "<session_id>"}
Servis: <modul>/json -> {"<servis_adi>": {"session_id", "firma_kodu", "donem_kodu", ...}}
Kaynak: https://doc.dia.com.tr/doku.php?id=gelistirici:wsapi:anasayfa
"""
import json
import time
import urllib.error
import urllib.request


class DiaHatasi(Exception):
    def __init__(self, code, msg, servis=''):
        self.code = str(code)
        self.msg = str(msg)
        self.servis = servis
        super().__init__((servis + ': ' if servis else '') + self.code + ' - ' + self.msg)


class DiaKontorHatasi(DiaHatasi):
    """406 CREDIT_ERROR: kontor yetersiz."""


# Oturumun dustugunu gosteren cevaplar (419 LOGIN_TIMEOUT, 'Invalid session')
_OTURUM_KODLARI = ('419', '401')


def _oturum_dustu_mu(code, msg):
    return str(code) in _OTURUM_KODLARI or 'invalid session' in str(msg).lower()


class DiaIstemci:
    def __init__(self, sunucu_kodu, kullanici, sifre, apikey, firma_kodu, donem_kodu,
                 zaman_asimi=60, temel_url=None, gonderici=None):
        """temel_url: test icin (varsayilan https://<sunucu>.ws.dia.com.tr/api/v3)
        gonderici : test icin (url, sozluk) -> sozluk cagrilabilir nesnesi"""
        self.temel_url = (temel_url or
                          'https://' + str(sunucu_kodu).strip() + '.ws.dia.com.tr/api/v3').rstrip('/')
        self.kullanici = kullanici
        self.sifre = sifre
        self.apikey = apikey
        self.firma_kodu = int(firma_kodu)
        self.donem_kodu = int(donem_kodu)
        self.zaman_asimi = zaman_asimi
        self._gonderici = gonderici or self._http_gonder
        self.session_id = None
        self._login_zamani = 0
        self.cagri_sayisi = 0

    # ------------------------------------------------------------
    def _http_gonder(self, url, govde):
        veri = json.dumps(govde, ensure_ascii=False).encode('utf-8')
        istek = urllib.request.Request(url, data=veri, method='POST',
                                       headers={'Content-Type': 'application/json;charset=UTF-8'})
        try:
            with urllib.request.urlopen(istek, timeout=self.zaman_asimi) as cevap:
                return json.loads(cevap.read().decode('utf-8'))
        except urllib.error.HTTPError as e:
            metin = e.read().decode('utf-8', 'replace')
            try:
                return json.loads(metin)
            except ValueError:
                raise DiaHatasi(e.code, metin[:300])
        except urllib.error.URLError as e:
            raise DiaHatasi('BAGLANTI', 'Sunucuya baglanilamadi: ' + str(e.reason))

    def _gonder(self, modul, govde):
        self.cagri_sayisi += 1
        return self._gonderici(self.temel_url + '/' + modul + '/json', govde)

    # ------------------------------------------------------------
    def login(self):
        cevap = self._gonder('sis', {'login': {
            'username': self.kullanici,
            'password': self.sifre,
            'disconnect_same_user': 'True',
            'lang': 'tr',
            'params': {'apikey': self.apikey},
        }})
        if str(cevap.get('code')) != '200':
            raise DiaHatasi(cevap.get('code'), cevap.get('msg'), 'login')
        self.session_id = cevap.get('msg')
        self._login_zamani = time.time()
        return self.session_id

    def logout(self):
        if not self.session_id:
            return
        try:
            self._gonder('sis', {'logout': {'session_id': self.session_id}})
        except DiaHatasi:
            pass
        self.session_id = None

    def _oturum(self):
        # Oturum 1 saatte duser; 50 dakikada bir tazele
        if not self.session_id or time.time() - self._login_zamani > 50 * 60:
            self.login()
        return self.session_id

    # ------------------------------------------------------------
    def cagir(self, modul, servis, firma_donem=True, **parametreler):
        """Servisi cagirir, basariliysa cevap sozlugunu dondurur. Oturum duserse bir kez yeniden dener."""
        for deneme in range(2):
            govde = {'session_id': self._oturum()}
            if firma_donem:
                govde['firma_kodu'] = self.firma_kodu
                govde['donem_kodu'] = self.donem_kodu
            govde.update(parametreler)
            cevap = self._gonder(modul, {servis: govde})
            code = str(cevap.get('code', '200'))
            if code == '200':
                return cevap
            msg = cevap.get('msg', '')
            if deneme == 0 and _oturum_dustu_mu(code, msg):
                self.session_id = None
                continue
            if code == '406':
                raise DiaKontorHatasi(code, msg, servis)
            raise DiaHatasi(code, msg, servis)

    def kontor_sorgula(self):
        cevap = self.cagir('sis', 'sis_kontor_sorgula', firma_donem=False)
        return cevap.get('result', cevap.get('msg'))

    def listele(self, modul, servis, filters=None, sorts=None, params=None, sayfa=1000):
        """Tum kayitlari sayfa sayfa getirir."""
        sonuc = []
        offset = 0
        while True:
            cevap = self.cagir(modul, servis, filters=filters or [], sorts=sorts or [],
                               params=params or {}, limit=sayfa, offset=offset)
            satirlar = cevap.get('result', [])
            if not isinstance(satirlar, list):
                break
            sonuc.extend(satirlar)
            if len(satirlar) < sayfa:
                break
            offset += sayfa
        return sonuc

    def ekle(self, modul, servis, kart):
        """Ekler, yeni kaydin _key degerini dondurur."""
        cevap = self.cagir(modul, servis, kart=kart)
        return yeni_key(cevap)

    def guncelle(self, modul, servis, kart):
        return self.cagir(modul, servis, kart=kart)


def yeni_key(cevap):
    """Ekleme cevabindaki _key: 'key', 'result' veya 'msg' alaninda gelebilir."""
    for alan in ('key', '_key', 'result', 'msg'):
        deger = cevap.get(alan)
        if isinstance(deger, dict):
            deger = deger.get('_key', deger.get('key'))
        if isinstance(deger, (int, float)) or (isinstance(deger, str) and deger.isdigit()):
            return int(deger)
    return None
