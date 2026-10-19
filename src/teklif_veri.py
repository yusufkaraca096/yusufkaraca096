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
    # --- Teklif kontrolleri (0 = kontrol kapalı)
    ('max_toplam_indirim', 25, 'Toplam indirim (satır + kampanya) liste fiyatına göre bu %\'yi geçerse KRİTİK uyarı'),
    ('min_kar_marji', 0, 'Kâr marjı bu %\'nin altına düşerse KRİTİK uyarı (maliyet girilen tekliflerde)'),
    ('min_saat_ucreti', 0, 'Teklife özel saatlik ücret bu tabanın altına inemez'),
    ('uyarlama_lisans_min', 0, 'Uyarlama / lisans oranı bu %\'nin altındaysa "riskli düşük uyarlama" uyarısı'),
    ('uyarlama_lisans_max', 0, 'Uyarlama / lisans oranı bu %\'nin üstündeyse "müşteri itiraz edebilir" uyarısı'),
    ('yonetici_sifresi', '', 'KRİTİK uyarılı teklifin PDF/Excel çıktısı için yönetici şifresi (boş = sadece uyarı)'),
    ('fiyat_listesi_max_gun', 60, 'Paket fiyat listesi bu günden eskiyse güncelleme uyarısı'),
    ('fiyat_listesi_tarihi', '', 'Son fiyat listesi içe aktarma tarihi (otomatik)'),
    # --- Seyahat
    ('yol_ucreti', 0, 'Şehir dışı: kişi başı gidiş-dönüş yol ücreti (ziyaret başına)'),
    ('konaklama_ucreti', 0, 'Şehir dışı: kişi başı gecelik konaklama'),
    ('gundelik', 0, 'Şehir dışı: kişi başı günlük harcırah'),
    ('ziyaret_basina_gun', 3, 'Bir ziyarette ortalama kaç gün kalınır (yol sayısı = gün / bu)'),
    # --- Bakim / destek anlasmasi
    ('destek_saat_ucreti', 0, 'Bakım anlaşması saat ücreti (0 = saatlik_ucret)'),
    ('destek_taban_saat', 12, 'Bakım: her müşteri için yıllık taban destek saati'),
    ('destek_kullanici_saat', 0.5, 'Bakım: kullanıcı başına yıllık ek saat'),
    ('destek_modul_saat', 0.25, 'Bakım: modül başına yıllık ek saat'),
    ('destek_entegrasyon_saat', 4, 'Bakım: entegrasyon başına yıllık ek saat'),
    ('destek_sube_saat', 2, 'Bakım: ek şube başına yıllık ek saat'),
    ('destek_uyarlama_orani', 10, 'Bakım: uyarlama saatinin bu %\'si yıllık destek saatine eklenir'),
    ('destek_yerinde_gun_ucreti', 0, 'Bakım: yerinde gün ücreti (0 = 8 × destek saat ücreti)'),
    ('destek_esikleri', '40; 100', 'Önerilen bakım paketi: yıllık saat < 1. eşik Temel, < 2. eşik Standart, üstü Premium'),
    # --- 3 yillik maliyet
    ('tco_yil', 3, 'Toplam sahip olma maliyeti tablosunun yıl sayısı'),
    ('ybh_baslangic_yili', 2, 'DIA yıllık bakım (YBH) kaçıncı yıldan başlar'),
    ('kapsam_haric', 'Kapsam dışı talepler (özel rapor / ekran geliştirme, ek eğitim, ek yerinde destek vb.) '
                     'saatlik ücret üzerinden ayrıca ücretlendirilir.', 'PDF\'te uyarlama kapsamının altına yazılır'),
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
CREATE TABLE IF NOT EXISTS hizmet (
    kod TEXT PRIMARY KEY, ad TEXT, hesap TEXT, p1 REAL, p2 REAL, carpan REAL DEFAULT 1, kapsam TEXT,
    sira INTEGER, aktif INTEGER DEFAULT 1);
CREATE TABLE IF NOT EXISTS kesif (
    kod TEXT PRIMARY KEY, soru TEXT, tur TEXT, secenekler TEXT, hedef TEXT, islem TEXT, deger TEXT,
    esik REAL DEFAULT 0, tetik TEXT, sira INTEGER, aktif INTEGER DEFAULT 1);
CREATE TABLE IF NOT EXISTS odeme_secenegi (
    ad TEXT PRIMARY KEY, tur TEXT, taksit INTEGER, indirim REAL, vade_farki REAL, aylik_vade REAL,
    komisyon REAL, musteriye_goster INTEGER, sira INTEGER);
CREATE TABLE IF NOT EXISTS bakim_paketi (
    ad TEXT PRIMARY KEY, saat_carpani REAL, yanit_suresi TEXT, yerinde_gun REAL, indirim REAL,
    aciklama TEXT, sira INTEGER);
CREATE TABLE IF NOT EXISTS gerceklesen (no TEXT, kalem TEXT, tahmin REAL, gercek REAL, PRIMARY KEY (no, kalem));
CREATE TABLE IF NOT EXISTS teklif (
    no TEXT PRIMARY KEY, tarih TEXT, musteri TEXT, paket_tipi TEXT, kullanici TEXT,
    lisans REAL, uyarlama REAL, net REAL, genel_toplam REAL, kampanya TEXT,
    sonuc TEXT DEFAULT 'BEKLİYOR', red_nedeni TEXT, veri TEXT);
"""


PAKET_KOLONLARI = ['kod', 'tur', 'ad', 'std_fiyat', 'ent_fiyat', 'kull_yuzde', 'ybh_yuzde', 'aciklama', 'urunler',
                   'std_saat', 'ent_saat', 'aktif', 'zorunlu']

# ------------------------------------------------------------------ varsayilan kurallar (ornek; Veriler'den degisir)
# hesap: PAKET_SAAT (paket adam/saatleri), KAYIT (kayit / p1 kayit-saat + p2 hazirlik), KULLANICI (tavan(kullanici /
# p1 kisi) x p2 saat), ADET (miktar x p1), SABIT (p1), YUZDE (diger hizmetlerin p1 %'si)
HIZMETLER = [
    ('kurulum', 'Kurulum ve parametre ayarları', 'PAKET_SAAT', 0, 0,
     'Paketlerin kurulumu, firma / dönem / şube tanımları, parametre ve kullanıcı yetki ayarları', 1),
    ('veri_aktarimi', 'Veri aktarımı', 'KAYIT', 2000, 4,
     'Cari, stok ve diğer kartların temizlenip DIA\'ya aktarılması, açılış bakiyeleri', 2),
    ('egitim', 'Kullanıcı eğitimi', 'KULLANICI', 8, 8, 'Kullanıcı grup eğitimleri', 3),
    ('rapor', 'Özel rapor / form tasarımı', 'ADET', 4, 0, 'Müşteriye özel rapor ve form (fatura, irsaliye vb.) tasarımı', 4),
    ('entegrasyon', 'Entegrasyonlar', 'ADET', 8, 0, 'E-ticaret, POS, banka vb. entegrasyonların kurulumu', 5),
    ('canli_gecis', 'Canlıya geçiş desteği', 'SABIT', 8, 0, 'Canlıya geçiş günü ve ilk hafta yoğun destek', 6),
    ('proje', 'Proje yönetimi', 'YUZDE', 10, 0, 'Planlama, koordinasyon, test ve kabul', 7),
]
# islem: EKLE (saat ekler), CARP (katsayi), CARP_YUZDE (1 + cevap/100), MIKTAR (hizmet miktarina ekler), BILGI (etkisiz)
# hedef: hizmet kodu veya * (tum hizmetler)
KESIF_SORULARI = [
    ('gecis_programi', 'Şu an hangi programı kullanıyor?', 'SECIM', 'Yok;Logo;Netsis;ETA;Mikro;Excel;Diğer',
     'veri_aktarimi', 'EKLE', 'Logo=8; Netsis=8; ETA=12; Mikro=8; Excel=4; Diğer=16; Yok=0', 0, '', 1),
    ('yeni_musteri', 'İlk kez DIA mı alıyor (yeni müşteri)?', 'EVET_HAYIR', '', '', 'BILGI', '', 0, 'EVET', 2),
    ('sube_sayisi', 'Şube sayısı (merkez dahil)', 'SAYI', '', 'kurulum', 'EKLE', '4', 1, '', 3),
    ('depo_sayisi', 'Depo sayısı', 'SAYI', '', 'kurulum', 'EKLE', '2', 1, '', 4),
    ('uretim', 'Üretim yapıyor mu?', 'EVET_HAYIR', '', 'kurulum', 'EKLE', '24', 0, 'EVET', 5),
    ('e_donusum', 'e-Fatura / e-Arşiv / e-Defter kurulacak mı?', 'EVET_HAYIR', '', 'kurulum', 'EKLE', '8', 0, 'EVET', 6),
    ('cari_sayisi', 'Aktarılacak cari sayısı', 'SAYI', '', 'veri_aktarimi', 'MIKTAR', '', 0, '', 7),
    ('stok_sayisi', 'Aktarılacak stok sayısı', 'SAYI', '', 'veri_aktarimi', 'MIKTAR', '', 0, '', 8),
    ('diger_kayit', 'Diğer aktarılacak kayıt (banka, personel, demirbaş, hesap planı)', 'SAYI', '', 'veri_aktarimi',
     'MIKTAR', '', 0, '', 9),
    ('veri_hata_orani', 'Verideki hata / mükerrer oranı (%)', 'SAYI', '', 'veri_aktarimi', 'CARP_YUZDE', '', 0, '', 10),
    ('rapor_sayisi', 'Özel rapor / form sayısı', 'SAYI', '', 'rapor', 'MIKTAR', '', 0, '', 11),
    ('entegrasyon_sayisi', 'Entegrasyon sayısı (e-ticaret, POS, banka ...)', 'SAYI', '', 'entegrasyon', 'MIKTAR', '', 0,
     '', 12),
    ('ek_egitim_gun', 'Ek eğitim günü (yönetici / ileri seviye)', 'SAYI', '', 'egitim', 'EKLE', '8', 0, '', 13),
    ('it_personeli_yok', 'Firmada bilgi işlem sorumlusu YOK mu?', 'EVET_HAYIR', '', '*', 'CARP', '1.15', 0, 'EVET', 14),
    ('acil_proje', 'Canlıya geçiş 1 aydan kısa sürede mi isteniyor?', 'EVET_HAYIR', '', '*', 'CARP', '1.1', 0, 'EVET', 15),
]
# tur: PESIN, TAKSIT (havale / cek-senet), KREDI_KARTI. Oranlar 0: kendi bankanizin oranlarini girin.
ODEME_SECENEKLERI = [
    ('Peşin (havale / EFT)', 'PESIN', 1, 5, 0, 0, 0, 1, 1),
    ('Havale 2 taksit', 'TAKSIT', 2, 0, 0, 0, 0, 1, 2),
    ('Havale 3 taksit', 'TAKSIT', 3, 0, 0, 0, 0, 1, 3),
    ('Havale 4 taksit', 'TAKSIT', 4, 0, 0, 0, 0, 0, 4),
    ('Kredi kartı tek çekim', 'KREDI_KARTI', 1, 0, 0, 0, 0, 1, 5),
    ('Kredi kartı 3 taksit', 'KREDI_KARTI', 3, 0, 0, 0, 0, 1, 6),
    ('Kredi kartı 6 taksit', 'KREDI_KARTI', 6, 0, 0, 0, 0, 1, 7),
    ('Kredi kartı 9 taksit', 'KREDI_KARTI', 9, 0, 0, 0, 0, 0, 8),
    ('Kredi kartı 12 taksit', 'KREDI_KARTI', 12, 0, 0, 0, 0, 0, 9),
    ('Çek / senet 3 ay', 'TAKSIT', 3, 0, 0, 0, 0, 0, 10),
]
BAKIM_PAKETLERI = [
    ('Temel', 0.6, '1 iş günü', 0, 0, 'Uzaktan destek (telefon / uzak bağlantı), yıllık saat havuzu', 1),
    ('Standart', 1.0, '4 iş saati', 1, 5, 'Uzaktan destek, yılda 1 gün yerinde ziyaret, sürüm geçiş desteği', 2),
    ('Premium', 1.5, '2 iş saati', 3, 10, 'Öncelikli destek, yılda 3 gün yerinde, periyodik sistem kontrolü, '
                                          'eğitim tazeleme', 3),
]
# Kampanya kosullari (ice aktarilan kampanya adlarina gore). Dil: 'kullanici>=10', 'soru:kod', 'soru:kod=A|B',
# 'tarih:12-15..12-31', 'elle' (otomatik onerilmez); ' ve ' ile birlestirilir.
KAMPANYA_KOSULLARI = {'toplu alım': 'kullanici>=10', 'rakip geçiş': 'soru:gecis_programi=Logo|Netsis|ETA|Mikro',
                      'yeni müşteri': 'soru:yeni_musteri', 'yıl sonu': 'tarih:12-15..12-31', 'özel kampanya': 'elle'}


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
        self._kolon_ekle('paket', 'zorunlu', 'TEXT')
        self._kolon_ekle('kampanya', 'kosul', 'TEXT')
        self._varsayilanlari_doldur()
        for ad, deger, _ in AYARLAR:
            self.db.execute('INSERT OR IGNORE INTO ayar VALUES (?, ?)', (ad, str(deger)))
        if not self.db.execute('SELECT 1 FROM kampanya').fetchone():
            self.db.execute("INSERT INTO kampanya (ad, yuzde, durum, aciklama, kosul) "
                            "VALUES ('YOK', 0, 'AKTİF', 'İndirim uygulanmaz', '')")
        self.db.commit()
        if otomatik_yedek and not self.bugun_yedek_var_mi():
            self.yedek_al('otomatik')

    def _kolon_ekle(self, tablo, kolon, tip):
        if kolon not in [r[1] for r in self.db.execute('PRAGMA table_info(%s)' % tablo)]:
            self.db.execute('ALTER TABLE %s ADD COLUMN %s %s' % (tablo, kolon, tip))

    def _varsayilanlari_doldur(self):
        """Bos kural tablolarini ornek kurallarla doldurur (kullanici sonra Veriler'den degistirir)."""
        def bos(t):
            return not self.db.execute('SELECT 1 FROM %s LIMIT 1' % t).fetchone()
        if bos('hizmet'):
            self.db.executemany('INSERT INTO hizmet (kod, ad, hesap, p1, p2, kapsam, sira, carpan, aktif) '
                                'VALUES (?,?,?,?,?,?,?,1,1)', HIZMETLER)
        if bos('kesif'):
            self.db.executemany('INSERT INTO kesif (kod, soru, tur, secenekler, hedef, islem, deger, esik, tetik, sira, '
                                'aktif) VALUES (?,?,?,?,?,?,?,?,?,?,1)', KESIF_SORULARI)
        if bos('odeme_secenegi'):
            secenekler = list(ODEME_SECENEKLERI)
            eski = self.db.execute("SELECT deger FROM ayar WHERE ad='pesin_indirim'").fetchone()
            if eski:       # ilk surumdeki pesin indirim ayari korunur
                secenekler[0] = secenekler[0][:3] + (_sayi(eski[0], 5),) + secenekler[0][4:]
            self.db.executemany('INSERT INTO odeme_secenegi VALUES (?,?,?,?,?,?,?,?,?)', secenekler)
        if bos('bakim_paketi'):
            self.db.executemany('INSERT INTO bakim_paketi VALUES (?,?,?,?,?,?,?)', BAKIM_PAKETLERI)

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

    def kurallar(self, gecmis=True):
        """Hesap motoruna giden kural tablolari (teklif_hesap.hesapla 'kurallar' parametresi)."""
        k = {
            'hizmetler': [dict(r) for r in self.db.execute('SELECT * FROM hizmet WHERE aktif=1 ORDER BY sira, kod')],
            'sorular': [dict(r) for r in self.db.execute('SELECT * FROM kesif WHERE aktif=1 ORDER BY sira, kod')],
            'odeme': [dict(r) for r in self.db.execute('SELECT * FROM odeme_secenegi ORDER BY sira, ad')],
            'bakim': [dict(r) for r in self.db.execute('SELECT * FROM bakim_paketi ORDER BY sira, ad')],
            'kampanyalar': self.kampanyalar(),
        }
        if gecmis:
            k['gecmis'] = self.gecmis_ozeti()
        return k

    def gecmis_ozeti(self):
        """Satis asistani icin sonuclanmis teklifler: [{no, net, genel_toplam, sonuc, indirim, kullanici, red}]"""
        sonuc = []
        for r in self.db.execute("SELECT no, net, genel_toplam, sonuc, red_nedeni, veri FROM teklif "
                                 "WHERE sonuc IN ('KAZANDI', 'KAYBETTİ')"):
            g = json.loads(r['veri']) if r['veri'] else {}
            sonuc.append({'no': r['no'], 'net': r['net'] or 0, 'genel_toplam': r['genel_toplam'] or 0,
                          'sonuc': r['sonuc'], 'red': r['red_nedeni'] or '',
                          'indirim': float((g.get('_ozet') or {}).get('toplam_indirim_yuzde', g.get('kampanya_yuzde') or 0)),
                          'kullanici': int(g.get('masaustu', 0) or 0) + int(g.get('mobil', 0) or 0)})
        return sonuc

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
        # Gerceklesen saat raporu ve satis asistani icin o anki tahminler saklanir (fiyatlar sonra degisse de)
        girdi['_ozet'] = {'hizmet': {h['kod']: h['saat'] for h in sonuc.get('hizmet', {}).get('kalemler', [])},
                          'toplam_indirim_yuzde': sonuc.get('kontrol', {}).get('toplam_indirim_yuzde', 0),
                          'genel_toplam': sonuc['genel_toplam']}
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

    def _paket_yaz(self, p):
        self.db.execute('INSERT OR REPLACE INTO paket (%s) VALUES (%s)' % (
            ','.join(PAKET_KOLONLARI), ','.join('?' * len(PAKET_KOLONLARI))), [p.get(k) for k in PAKET_KOLONLARI])

    # ------------------------------------------------------------ revizyon / gerceklesen saat
    def revizyon_no(self, no):
        """RT-2026-0001 veya RT-2026-0001-R1 -> bir sonraki revizyon numarasi."""
        ana = re.sub(r'-R\d+$', '', no)
        en_buyuk = 0
        for r in self.db.execute('SELECT no FROM teklif WHERE no LIKE ?', (ana + '-R%',)):
            m = re.search(r'-R(\d+)$', r['no'])
            if m:
                en_buyuk = max(en_buyuk, int(m.group(1)))
        return '%s-R%d' % (ana, en_buyuk + 1)

    def gerceklesen_yaz(self, no, kalemler):
        """kalemler: {kalem: (tahmin, gercek)}"""
        self.degisti = True
        with self.db:
            self.db.execute('DELETE FROM gerceklesen WHERE no=?', (no,))
            for kalem, (tahmin, gercek) in kalemler.items():
                if gercek is not None:
                    self.db.execute('INSERT INTO gerceklesen VALUES (?,?,?,?)', (no, kalem, tahmin, gercek))

    def gerceklesen_oku(self, no):
        return {r['kalem']: (r['tahmin'], r['gercek'])
                for r in self.db.execute('SELECT * FROM gerceklesen WHERE no=?', (no,))}

    def tahmin_gercek_raporu(self):
        """Hizmet kalemi bazinda: proje sayisi, toplam tahmin / gercek saat, sapma %, onerilen carpan."""
        adlar = {r['kod']: (r['ad'], r['carpan'] or 1) for r in self.db.execute('SELECT kod, ad, carpan FROM hizmet')}
        rapor = []
        for r in self.db.execute('SELECT kalem, COUNT(*) n, SUM(tahmin) t, SUM(gercek) g FROM gerceklesen '
                                 'GROUP BY kalem ORDER BY kalem'):
            t, g = r['t'] or 0, r['g'] or 0
            ad, carpan = adlar.get(r['kalem'], (r['kalem'], 1))
            rapor.append({'kalem': r['kalem'], 'ad': ad, 'proje': r['n'], 'tahmin': t, 'gercek': g,
                          'sapma': round((g / t - 1) * 100, 1) if t else None, 'carpan': carpan,
                          'onerilen_carpan': round(carpan * g / t, 2) if t and g else None})
        return rapor

    def carpan_yaz(self, kalem, carpan):
        self.degisti = True
        with self.db:
            self.db.execute('UPDATE hizmet SET carpan=? WHERE kod=?', (carpan, kalem))

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
                'zorunlu': zorunlu_metni(ws.cell(r, 10).value), 'std_saat': 0.0, 'ent_saat': 0.0, 'aktif': 1}
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
                                        'aciklama': _metin(ws.cell(r, 4).value),
                                        'kosul': KAMPANYA_KOSULLARI.get(ad.replace('İ', 'i').lower(), '')})
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
                self._paket_yaz(p)
            ayarlar['fiyat_listesi_tarihi'] = datetime.now().strftime('%Y-%m-%d')
            for ad, deger in ayarlar.items():
                self.db.execute('INSERT OR REPLACE INTO ayar VALUES (?, ?)', (ad, str(deger)))
            for ad, fiyat in sabitler.items():
                self.db.execute('INSERT OR REPLACE INTO sabit_kalem VALUES (?, ?, 1)', (ad, fiyat))
            if kampanyalar:
                self.db.execute('DELETE FROM kampanya')
                for k in kampanyalar:
                    self.db.execute('INSERT OR REPLACE INTO kampanya (ad, yuzde, durum, aciklama, kosul) '
                                    'VALUES (:ad,:yuzde,:durum,:aciklama,:kosul)', k)
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
                     'tür': 'tur', 'tur': 'tur', 'grup': 'tur', 'zorunlu': 'zorunlu'}
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
                                     'aciklama': _metin(al('aciklama')), 'urunler': _metin(al('urunler')),
                                     'zorunlu': zorunlu_metni(al('zorunlu'))}
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
                self._paket_yaz(p)
            self.db.execute('INSERT OR REPLACE INTO ayar VALUES (?, ?)',
                            ('fiyat_listesi_tarihi', datetime.now().strftime('%Y-%m-%d')))
        return {'toplam': len(yeni), 'fiyati_degisen': degisen, 'yeni': eklenen, 'listeden_kalkan': kalkan}


def zorunlu_metni(x):
    """'[scf, kasa]' -> 'scf, kasa'; '[]' -> ''"""
    return ', '.join(p.strip() for p in _metin(x).strip('[]').split(',') if p.strip())


def musteri_veri_ozeti(klasor):
    """Veri temizleme programinin musteri klasorunden kayit sayilari ve hata orani.
    Doner: {'cari': n, 'stok': n, 'diger': n, 'hata_orani': %, 'dosyalar': [...]}"""
    from openpyxl import load_workbook
    sonuc = {'cari': 0, 'stok': 0, 'diger': 0, 'hata_orani': 0.0, 'dosyalar': []}
    toplam, sorunlu = 0, 0
    for modul in ('cari', 'stok', 'banka', 'personel', 'demirbas', 'hesapplan'):
        for ad in (modul + '_dia_hazir.xlsx', modul + '_listesi.xlsx'):
            yol = os.path.join(klasor, ad)
            if not os.path.exists(yol):
                continue
            try:
                wb = load_workbook(yol, read_only=True)
                n = max(wb.active.max_row - 1, 0)
                wb.close()
            except Exception:
                continue
            sonuc['cari' if modul == 'cari' else 'stok' if modul == 'stok' else 'diger'] += n
            sonuc['dosyalar'].append('%s (%d)' % (ad, n))
            toplam += n
            ozet = os.path.join(klasor, modul + '_OZET.xlsx')
            if os.path.exists(ozet):
                try:
                    import pandas as pd
                    detay = pd.read_excel(ozet, sheet_name='DETAY')
                    sorunlu += min(detay['satir'].nunique(), n)
                except Exception:
                    pass
            break
    if toplam:
        sonuc['hata_orani'] = round(sorunlu * 100.0 / toplam, 1)
    return sonuc


def sifre_ozeti(sifre):
    import hashlib
    return hashlib.sha256(('raca-teklif:' + sifre).encode('utf-8')).hexdigest() if sifre else ''


def _tablo_adi(t):
    if t not in ('paket', 'ayar', 'sabit_kalem', 'kampanya', 'rakip', 'hizmet', 'kesif', 'odeme_secenegi',
                 'bakim_paketi'):
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
