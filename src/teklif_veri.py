"""Teklif motoru verileri: paket listesi, ayarlar, sabit kalemler, kampanyalar, rakipler, teklif gecmisi.

Veriler yerel SQLite dosyasinda (teklif.db) tutulur. Fiyatlar ve ucretler kodda YOKTUR
(depo herkese acik); ilk calistirmada mevcut teklif Excel'i ice aktarilir.
"""
import glob
import json
import os
import re
import shutil
import sqlite3
from datetime import datetime

YEDEK_SAYISI = 30

# ad, varsayilan, aciklama  (ticari degerler 0: ice aktarilinca dolar)
AYARLAR = [
    ('ent_masaustu_esik', 10, 'Masaüstü kullanıcı bu sayıyı geçerse paket ENTERPRISE olur'),
    ('ent_mobil_esik', 10, 'Mobil kullanıcı bu sayıyı geçerse paket ENTERPRISE olur'),
    ('lisans_yuzdesi', 10, 'Uyarlama = lisans net toplamı × bu % + adam/saat bedeli'),
    ('saatlik_ucret', 0, 'Saatlik danışman ücreti (KDV hariç)'),
    ('min_uyarlama', 0, 'Uyarlama alt sınırı (KDV hariç)'),
    ('max_uyarlama', 0, 'Uyarlama üst sınırı (0 = sınır yok)'),
    ('kdv', 20, 'KDV oranı (%)'),
    ('pesin_indirim', 5, 'Peşin ödemede indirim (%)'),
    ('taksitler', '2,3,4', 'Ödeme planındaki taksit sayıları (virgülle)'),
    ('rakip_ayni_saat', 1, '1: rakip = bizim adam/saat × rakibin saatlik ücreti (aynı iş). 0: Rakipler tablosundaki saat'),
    ('rakip_ucuz_esik', 10, 'Rakip ortalamasından bu % kadar ucuzsak "Rekabetçisiniz"'),
    ('rakip_pahali_esik', 15, 'Rakip ortalamasından bu % kadar pahalıysak "Pahalısınız"'),
    ('gecerlilik_gun', 15, 'Teklifin geçerlilik süresi (gün)'),
    ('teklif_no_onek', 'RT', 'Teklif numarası ön eki: RT-2026-0001'),
    ('firma_adi', 'RACA TEKNOLOJİ', 'PDF/Excel teklifin başlığındaki firma adı'),
    ('firma_adres', '', 'Firma adresi'),
    ('firma_iletisim', '', 'Telefon / e-posta / web (tek satır)'),
    ('logo_yolu', '', 'Logo dosyası (PNG/JPG). Boşsa logo basılmaz'),
    ('teklif_notlari', 'Fiyatlar KDV hariç belirtilmiş olup KDV ayrıca gösterilmiştir.\n'
                       'Uyarlama bedeli tahmini adam/saat üzerinden hesaplanmıştır.',
     'Teklifin altına yazılacak şartlar / notlar'),
    ('ybh_modul_kademe', '5:1; 10:1.1; 20:1.25; 30:1.45; *:1.65', 'YBH modül sayısı katsayıları (en fazla:katsayı)'),
    ('ybh_kullanici_kademe', '5:1; 10:1.1; 25:1.2; 50:1.35; *:1.55', 'YBH kullanıcı sayısı katsayıları'),
]
AYAR_VARSAYILAN = {a: v for a, v, _ in AYARLAR}
SAYI_AYARLARI = {a for a, v, _ in AYARLAR if isinstance(v, (int, float))}

SEMA = """
CREATE TABLE IF NOT EXISTS paket (
    kod TEXT PRIMARY KEY, tur TEXT, ad TEXT, std_fiyat REAL, ent_fiyat REAL,
    kull_yuzde REAL, ybh_yuzde REAL, aciklama TEXT, urunler TEXT,
    std_saat REAL DEFAULT 0, ent_saat REAL DEFAULT 0, aktif INTEGER DEFAULT 1);
CREATE TABLE IF NOT EXISTS ayar (ad TEXT PRIMARY KEY, deger TEXT);
CREATE TABLE IF NOT EXISTS sabit_kalem (ad TEXT PRIMARY KEY, fiyat REAL, aktif INTEGER DEFAULT 1);
CREATE TABLE IF NOT EXISTS kampanya (ad TEXT PRIMARY KEY, yuzde REAL, durum TEXT, aciklama TEXT);
CREATE TABLE IF NOT EXISTS rakip (firma TEXT PRIMARY KEY, saat REAL, ucret REAL);
CREATE TABLE IF NOT EXISTS teklif (
    no TEXT PRIMARY KEY, tarih TEXT, musteri TEXT, paket_tipi TEXT, kullanici TEXT,
    lisans REAL, uyarlama REAL, net REAL, genel_toplam REAL, kampanya TEXT,
    sonuc TEXT DEFAULT 'BEKLİYOR', red_nedeni TEXT, veri TEXT);
"""


def _sayi(x, varsayilan=0.0):
    if x is None or (isinstance(x, str) and not x.strip()):
        return varsayilan
    if isinstance(x, (int, float)):
        return float(x)
    try:
        return float(str(x).replace(' ', '').replace('.', '').replace(',', '.')) if ',' in str(x) \
            else float(str(x).replace(' ', ''))
    except ValueError:
        return varsayilan


def _metin(x):
    return '' if x is None else str(x).strip()


class TeklifVeri:
    def __init__(self, yol, otomatik_yedek=True):
        """yol: teklif.db dosyasi (yoksa olusturulur). Acilista gunde bir otomatik yedek alinir."""
        self.yol = os.path.abspath(yol)
        self.klasor = os.path.dirname(self.yol)
        self.yedek_klasoru = os.path.join(self.klasor, 'Yedekler')
        os.makedirs(self.klasor, exist_ok=True)
        self.db = sqlite3.connect(self.yol)
        self.degisti = False         # son yedekten beri veri degisti mi (kapanista yedek icin)
        self.db.row_factory = sqlite3.Row
        self.db.executescript(SEMA)
        for ad, deger, _ in AYARLAR:
            self.db.execute('INSERT OR IGNORE INTO ayar VALUES (?, ?)', (ad, str(deger)))
        if not self.db.execute('SELECT 1 FROM kampanya').fetchone():
            self.db.execute("INSERT INTO kampanya VALUES ('YOK', 0, 'AKTİF', 'İndirim uygulanmaz')")
        self.db.commit()
        if otomatik_yedek and not self.bugun_yedek_var_mi():
            self.yedek_al('otomatik')

    # ------------------------------------------------------------ yedek
    def yedekler(self):
        """En yeni once: [(yol, ad)]"""
        dosyalar = sorted(glob.glob(os.path.join(self.yedek_klasoru, 'teklif_*.db')), reverse=True)
        return [(d, os.path.basename(d)) for d in dosyalar]

    def bugun_yedek_var_mi(self):
        bugun = datetime.now().strftime('%Y-%m-%d')
        return any(ad.startswith('teklif_' + bugun) for _, ad in self.yedekler())

    def yedek_al(self, neden='elle', en_az_dakika=0):
        """SQLite yedekleme API'si: program acikken de tutarli kopya alir. Eski yedeklerden son 30'u kalir.
        en_az_dakika: son yedek bundan yeniyse tekrar alinmaz (sik duzenlemede gunluk yedekler silinmesin)."""
        if en_az_dakika and self.yedekler():
            son = os.path.getmtime(self.yedekler()[0][0])
            if (datetime.now().timestamp() - son) < en_az_dakika * 60:
                return None
        if self.bos_mu() and not self.db.execute('SELECT 1 FROM teklif LIMIT 1').fetchone():
            return None     # bos veritabaninin yedegi alinmaz (iyi yedekleri silip yer kaplamasin)
        os.makedirs(self.yedek_klasoru, exist_ok=True)
        hedef = os.path.join(self.yedek_klasoru, 'teklif_%s_%s.db' % (
            datetime.now().strftime('%Y-%m-%d_%H%M%S'), re.sub(r'\W', '', neden)[:20] or 'yedek'))
        kopya = sqlite3.connect(hedef)
        with kopya:
            self.db.backup(kopya)
        kopya.close()
        self.degisti = False
        for eski, _ in self.yedekler()[YEDEK_SAYISI:]:
            try:
                os.remove(eski)
            except OSError:
                pass
        return hedef

    def yedekten_don(self, yedek_yolu):
        """Once mevcut hali yedekler, sonra secilen yedegi geri yukler."""
        kaynak = sqlite3.connect(yedek_yolu)
        try:
            kaynak.execute('SELECT 1 FROM teklif LIMIT 1')     # gecerli teklif veritabani mi
        except sqlite3.DatabaseError:
            kaynak.close()
            raise ValueError('Seçilen dosya teklif yedeği değil: ' + yedek_yolu)
        self.yedek_al('geri_donus_oncesi')
        with self.db:
            kaynak.backup(self.db)
        kaynak.close()

    def kapat(self):
        """Kapanista, son yedekten beri degisiklik varsa yedek alinir."""
        try:
            if self.degisti:
                self.yedek_al('kapanis')
        finally:
            self.db.close()

    # ------------------------------------------------------------ okuma
    def bos_mu(self):
        return not self.db.execute('SELECT 1 FROM paket LIMIT 1').fetchone()

    def ayarlar(self):
        sonuc = dict(AYAR_VARSAYILAN)
        for r in self.db.execute('SELECT ad, deger FROM ayar'):
            sonuc[r['ad']] = _sayi(r['deger'], AYAR_VARSAYILAN.get(r['ad'], 0)) \
                if r['ad'] in SAYI_AYARLARI else r['deger']
        return sonuc

    def paketler(self, sadece_aktif=False):
        sql = 'SELECT * FROM paket' + (' WHERE aktif=1' if sadece_aktif else '') + ' ORDER BY tur, kod'
        return {r['kod']: dict(r) for r in self.db.execute(sql)}

    def sabit_kalemler(self):
        return {r['ad']: r['fiyat'] for r in self.db.execute('SELECT * FROM sabit_kalem WHERE aktif=1 ORDER BY ad')}

    def kampanyalar(self, sadece_aktif=True):
        sql = 'SELECT * FROM kampanya' + (" WHERE durum='AKTİF'" if sadece_aktif else '') + ' ORDER BY yuzde, ad'
        return [dict(r) for r in self.db.execute(sql)]

    def rakipler(self):
        return [dict(r) for r in self.db.execute('SELECT * FROM rakip ORDER BY firma')]

    # ------------------------------------------------------------ yazma
    def ayar_yaz(self, ad, deger):
        self.degisti = True
        self.db.execute('INSERT OR REPLACE INTO ayar VALUES (?, ?)', (ad, str(deger)))
        self.db.commit()

    def tablo_oku(self, tablo):
        return [dict(r) for r in self.db.execute('SELECT * FROM ' + _tablo_adi(tablo))]

    def tablo_yaz(self, tablo, satirlar):
        """Tabloyu tamamen degistirir (duzenleme penceresinden)."""
        self.degisti = True
        tablo = _tablo_adi(tablo)
        self.yedek_al(tablo + '_duzenleme', en_az_dakika=30)
        kolonlar = [r[1] for r in self.db.execute('PRAGMA table_info(%s)' % tablo)]
        with self.db:
            self.db.execute('DELETE FROM ' + tablo)
            for s in satirlar:
                self.db.execute('INSERT OR REPLACE INTO %s (%s) VALUES (%s)' % (
                    tablo, ','.join(kolonlar), ','.join('?' * len(kolonlar))), [s.get(k) for k in kolonlar])

    # ------------------------------------------------------------ teklif gecmisi
    def yeni_teklif_no(self):
        onek = self.ayarlar()['teklif_no_onek'] or 'RT'
        yil = datetime.now().year
        bas = '%s-%d-' % (onek, yil)
        en_buyuk = 0
        for r in self.db.execute('SELECT no FROM teklif WHERE no LIKE ?', (bas + '%',)):
            m = re.search(r'(\d+)$', r['no'])
            if m:
                en_buyuk = max(en_buyuk, int(m.group(1)))
        return bas + str(en_buyuk + 1).zfill(4)

    def teklif_kaydet(self, girdi, sonuc, no=None):
        self.degisti = True
        no = no or girdi.get('teklif_no') or self.yeni_teklif_no()
        girdi = dict(girdi, teklif_no=no)
        eski = self.db.execute('SELECT sonuc, red_nedeni FROM teklif WHERE no=?', (no,)).fetchone()
        with self.db:
            self.db.execute('INSERT OR REPLACE INTO teklif VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)', (
                no, girdi.get('tarih') or datetime.now().strftime('%Y-%m-%d'), girdi.get('musteri', ''),
                sonuc['paket_tipi'], '%d masaüstü / %d mobil' % (girdi.get('masaustu', 0), girdi.get('mobil', 0)),
                sonuc['lisans']['net'], sonuc['uyarlama']['net'], sonuc['net_toplam'], sonuc['genel_toplam'],
                sonuc['kampanya']['ad'], eski['sonuc'] if eski else 'BEKLİYOR', eski['red_nedeni'] if eski else '',
                json.dumps(girdi, ensure_ascii=False)))
        self.yedek_al('teklif_kaydi', en_az_dakika=30)
        return no

    def teklifler(self, ara=''):
        sql = 'SELECT no, tarih, musteri, paket_tipi, kullanici, lisans, uyarlama, net, genel_toplam, kampanya, ' \
              'sonuc, red_nedeni FROM teklif'
        par = ()
        if ara:
            sql += ' WHERE musteri LIKE ? OR no LIKE ?'
            par = ('%' + ara + '%', '%' + ara + '%')
        return [dict(r) for r in self.db.execute(sql + ' ORDER BY no DESC', par)]

    def teklif_girdisi(self, no):
        r = self.db.execute('SELECT veri FROM teklif WHERE no=?', (no,)).fetchone()
        return json.loads(r['veri']) if r and r['veri'] else None

    def teklif_sonuc_yaz(self, no, sonuc, red_nedeni=''):
        self.degisti = True
        with self.db:
            self.db.execute('UPDATE teklif SET sonuc=?, red_nedeni=? WHERE no=?', (sonuc, red_nedeni, no))

    def teklif_sil(self, no):
        self.degisti = True
        with self.db:
            self.db.execute('DELETE FROM teklif WHERE no=?', (no,))

    # ------------------------------------------------------------ ice aktarma
    def excelden_ice_aktar(self, yol):
        """RACA_DIA_Teklif_Motoru.xlsx'ten paket listesi, adam/saat, ayarlar, kampanya, rakip, gecmis."""
        from openpyxl import load_workbook
        wb = load_workbook(yol, data_only=True)
        self.yedek_al('ice_aktarma_oncesi')
        rapor = []
        if 'PaketListesi' not in wb.sheetnames:
            raise ValueError("Dosyada 'PaketListesi' sayfasi yok")

        ws = wb['PaketListesi']
        paketler = {}
        satir_kod = {}
        for r in range(2, ws.max_row + 1):
            kod = _metin(ws.cell(r, 2).value)
            if not kod:
                continue
            satir_kod[r] = kod
            paketler[kod] = {
                'kod': kod, 'tur': _metin(ws.cell(r, 1).value), 'ad': _metin(ws.cell(r, 3).value),
                'std_fiyat': _sayi(ws.cell(r, 4).value), 'ent_fiyat': _sayi(ws.cell(r, 5).value),
                'kull_yuzde': _sayi(ws.cell(r, 6).value), 'ybh_yuzde': _sayi(ws.cell(r, 7).value),
                'aciklama': _metin(ws.cell(r, 8).value), 'urunler': _metin(ws.cell(r, 9).value),
                'std_saat': 0.0, 'ent_saat': 0.0, 'aktif': 1}
        rapor.append('%d paket' % len(paketler))

        if 'ADAM SAAT' in wb.sheetnames:
            ws = wb['ADAM SAAT']
            bulunan = 0
            for r in range(5, ws.max_row + 1):
                # Excel'de A kolonu PaketListesi'nin (r-3). satirindaki kodu gosterir
                kod = _metin(ws.cell(r, 1).value) or satir_kod.get(r - 3, '')
                if kod in paketler:
                    std = _sayi(ws.cell(r, 3).value)
                    ent = ws.cell(r, 4).value
                    paketler[kod]['std_saat'] = std
                    paketler[kod]['ent_saat'] = round(std * 1.5) if ent is None else _sayi(ent)
                    bulunan += 1
            rapor.append('%d paketin adam/saati' % bulunan)

        ayar_adlari = {'enterprise masaüstü eşiği': 'ent_masaustu_esik', 'enterprise mobil eşiği': 'ent_mobil_esik',
                       'lisans yüzdesi (%)': 'lisans_yuzdesi', 'saatlik danışman ücreti': 'saatlik_ucret',
                       'minimum uyarlama': 'min_uyarlama', 'maksimum uyarlama': 'max_uyarlama',
                       'kdv oranı (%)': 'kdv'}
        ayarlar, sabitler = {}, {}
        if 'AYARLAR' in wb.sheetnames:
            ws = wb['AYARLAR']
            for r in range(4, ws.max_row + 1):
                ad = _metin(ws.cell(r, 1).value)
                deger = ws.cell(r, 2).value
                if not ad or deger is None:
                    continue
                anahtar = ayar_adlari.get(ad.lower().replace('İ', 'i'))
                if anahtar is None:
                    anahtar = ayar_adlari.get(ad.replace('İ', 'i').lower())
                if anahtar:
                    ayarlar[anahtar] = _sayi(deger)
                elif 'sabit' in _metin(ws.cell(r, 3).value).lower():
                    sabitler[ad] = _sayi(deger)
            rapor.append('%d ayar, %d sabit kalem' % (len(ayarlar), len(sabitler)))

        kampanyalar = []
        if 'KAMPANYALAR' in wb.sheetnames:
            ws = wb['KAMPANYALAR']
            for r in range(4, ws.max_row + 1):
                ad = _metin(ws.cell(r, 1).value)
                if ad:
                    kampanyalar.append({'ad': ad, 'yuzde': _sayi(ws.cell(r, 2).value),
                                        'durum': _metin(ws.cell(r, 3).value) or 'AKTİF',
                                        'aciklama': _metin(ws.cell(r, 4).value)})
            rapor.append('%d kampanya' % len(kampanyalar))

        rakipler = []
        if 'RAKİP' in wb.sheetnames:
            ws = wb['RAKİP']
            for r in range(5, ws.max_row + 1):
                firma = _metin(ws.cell(r, 1).value)
                if firma and 'biz' not in firma.lower() and not firma.endswith(':'):
                    rakipler.append({'firma': firma, 'saat': _sayi(ws.cell(r, 2).value),
                                     'ucret': _sayi(ws.cell(r, 3).value)})
            rapor.append('%d rakip' % len(rakipler))

        gecmis = 0
        with self.db:
            # Bilinen kodlarin saatleri korunur, listede olmayanlar pasif yapilir (eski teklifler bozulmasin)
            self.db.execute('UPDATE paket SET aktif=0')
            for p in paketler.values():
                self.db.execute('INSERT OR REPLACE INTO paket VALUES (:kod,:tur,:ad,:std_fiyat,:ent_fiyat,:kull_yuzde,'
                                ':ybh_yuzde,:aciklama,:urunler,:std_saat,:ent_saat,:aktif)', p)
            for ad, deger in ayarlar.items():
                self.db.execute('INSERT OR REPLACE INTO ayar VALUES (?, ?)', (ad, str(deger)))
            for ad, fiyat in sabitler.items():
                self.db.execute('INSERT OR REPLACE INTO sabit_kalem VALUES (?, ?, 1)', (ad, fiyat))
            if kampanyalar:
                self.db.execute('DELETE FROM kampanya')
                for k in kampanyalar:
                    self.db.execute('INSERT OR REPLACE INTO kampanya VALUES (:ad,:yuzde,:durum,:aciklama)', k)
            if rakipler:
                self.db.execute('DELETE FROM rakip')
                for k in rakipler:
                    self.db.execute('INSERT OR REPLACE INTO rakip VALUES (:firma,:saat,:ucret)', k)
            if 'TEKLİF GEÇMİŞİ' in wb.sheetnames:
                ws = wb['TEKLİF GEÇMİŞİ']
                for r in range(5, ws.max_row + 1):
                    v = [ws.cell(r, c).value for c in range(1, 12)]
                    if not v[0] or not v[2]:
                        continue
                    tarih = v[1].strftime('%Y-%m-%d') if isinstance(v[1], datetime) else _metin(v[1])
                    self.db.execute('INSERT OR IGNORE INTO teklif VALUES (?,?,?,?,?,?,?,?,?,?,?,?,NULL)', (
                        _metin(v[0]), tarih, _metin(v[2]), _metin(v[3]), _metin(v[4]), _sayi(v[5]), _sayi(v[6]),
                        _sayi(v[7]) / (1 + self.ayarlar()['kdv'] / 100), _sayi(v[7]), _metin(v[8]),
                        _metin(v[9]) or 'BEKLİYOR', _metin(v[10])))
                    gecmis += 1
        if gecmis:
            rapor.append('%d geçmiş teklif' % gecmis)
        self.yedek_al('ice_aktarma_sonrasi')
        return rapor

    def paket_listesi_guncelle(self, yol):
        """DIA'nin yeni fiyat Excel'i: basliklar Kodu / Adı / Std. Fiyat / Ent. Fiyat / Kullanıcı % / YBH % /
        Açıklama / Ürünler (Tür: 'Tür' kolonu veya Kodu'nun solundaki kolon). Adam/saatler kod uzerinden korunur."""
        from openpyxl import load_workbook
        wb = load_workbook(yol, data_only=True, read_only=True)
        basliklar = {'kodu': 'kod', 'kod': 'kod', 'paket kodu': 'kod', 'adı': 'ad', 'adi': 'ad', 'paket adı': 'ad',
                     'std. fiyat': 'std_fiyat', 'std fiyat': 'std_fiyat', 'standart fiyat': 'std_fiyat',
                     'ent. fiyat': 'ent_fiyat', 'ent fiyat': 'ent_fiyat', 'enterprise fiyat': 'ent_fiyat',
                     'kullanıcı %': 'kull_yuzde', 'kullanici %': 'kull_yuzde', 'ybh %': 'ybh_yuzde',
                     'açıklama': 'aciklama', 'aciklama': 'aciklama', 'ürünler': 'urunler', 'urunler': 'urunler',
                     'tür': 'tur', 'tur': 'tur', 'grup': 'tur'}
        yeni = {}
        for ws in wb.worksheets:
            satirlar = list(ws.iter_rows(values_only=True))
            for i, satir in enumerate(satirlar[:20]):
                eslesme = {}
                for c, v in enumerate(satir):
                    k = basliklar.get(_metin(v).lower())
                    if k and k not in eslesme:
                        eslesme[k] = c
                if 'kod' in eslesme and 'std_fiyat' in eslesme:
                    if 'tur' not in eslesme and eslesme['kod'] > 0:
                        eslesme['tur'] = eslesme['kod'] - 1
                    for s in satirlar[i + 1:]:
                        kod = _metin(s[eslesme['kod']]) if eslesme['kod'] < len(s) else ''
                        if not kod:
                            continue
                        al = lambda k: s[eslesme[k]] if k in eslesme and eslesme[k] < len(s) else None  # noqa: E731
                        yeni[kod] = {'kod': kod, 'tur': _metin(al('tur')), 'ad': _metin(al('ad')),
                                     'std_fiyat': _sayi(al('std_fiyat')), 'ent_fiyat': _sayi(al('ent_fiyat')),
                                     'kull_yuzde': _sayi(al('kull_yuzde')), 'ybh_yuzde': _sayi(al('ybh_yuzde')),
                                     'aciklama': _metin(al('aciklama')), 'urunler': _metin(al('urunler'))}
                    break
            if yeni:
                break
        wb.close()
        if not yeni:
            raise ValueError("Başlık satırı bulunamadı: 'Kodu' ve 'Std. Fiyat' kolonları olmalı")
        self.yedek_al('paket_guncelleme_oncesi')
        eski = self.paketler()
        degisen = [k for k in yeni if k in eski and (eski[k]['std_fiyat'], eski[k]['ent_fiyat']) !=
                   (yeni[k]['std_fiyat'], yeni[k]['ent_fiyat'])]
        eklenen = [k for k in yeni if k not in eski]
        kalkan = [k for k in eski if k not in yeni and eski[k]['aktif']]
        with self.db:
            self.db.execute('UPDATE paket SET aktif=0')
            for kod, p in yeni.items():
                p['std_saat'] = eski.get(kod, {}).get('std_saat', 0)
                p['ent_saat'] = eski.get(kod, {}).get('ent_saat', 0)
                p['aktif'] = 1
                self.db.execute('INSERT OR REPLACE INTO paket VALUES (:kod,:tur,:ad,:std_fiyat,:ent_fiyat,:kull_yuzde,'
                                ':ybh_yuzde,:aciklama,:urunler,:std_saat,:ent_saat,:aktif)', p)
        return {'toplam': len(yeni), 'fiyati_degisen': degisen, 'yeni': eklenen, 'listeden_kalkan': kalkan}


def _tablo_adi(t):
    if t not in ('paket', 'ayar', 'sabit_kalem', 'kampanya', 'rakip'):
        raise ValueError('Bilinmeyen tablo: ' + str(t))
    return t


def _ayar_dosyasi():
    temel = os.environ.get('APPDATA') or os.path.expanduser('~')
    return os.path.join(temel, 'RACA_Teklif', 'ayar.json')


def varsayilan_klasor():
    """Veri EXE'nin yanina YAZILMAZ: EXE Indirilenler/zip/gecici klasorden calisinca veya yeni surum baska
    klasore konunca kayitlar kayboluyordu. Varsayilan: Belgeler\\RACA_Teklif_Veri (degistirilebilir)."""
    try:
        with open(_ayar_dosyasi(), encoding='utf-8') as f:
            klasor = json.load(f).get('veri_klasoru')
        if klasor:
            return klasor
    except (OSError, ValueError):
        pass
    belgeler = os.path.join(os.path.expanduser('~'), 'Documents')
    if not os.path.isdir(belgeler):
        belgeler = os.path.expanduser('~')
    return os.path.join(belgeler, 'RACA_Teklif_Veri')


def veri_klasoru_degistir(yeni_klasor, tasi=True):
    """Veri klasorunu degistirir (or. ortak ag klasoru). tasi=True ve hedefte teklif.db yoksa mevcut veri kopyalanir."""
    eski = os.path.join(varsayilan_klasor(), 'teklif.db')
    hedef = os.path.join(yeni_klasor, 'teklif.db')
    os.makedirs(yeni_klasor, exist_ok=True)
    if tasi and os.path.exists(eski) and not os.path.exists(hedef):
        shutil.copy2(eski, hedef)
    os.makedirs(os.path.dirname(_ayar_dosyasi()), exist_ok=True)
    with open(_ayar_dosyasi(), 'w', encoding='utf-8') as f:
        json.dump({'veri_klasoru': yeni_klasor}, f, ensure_ascii=False)
    return hedef
