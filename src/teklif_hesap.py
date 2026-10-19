"""Teklif hesap motoru (RACA_DIA_Teklif_Motoru.xlsx > HESAP / YBH BAKIM sayfalarinin karsiligi ve fazlasi).

Excel'deki mantik korunmustur; su hatalar duzeltilmistir:
  - K24 hucresindeki bozuk formul (#NAME?) tum toplamlari bozuyordu.
  - Odeme plani ve sablondaki NET TOPLAM, net tutar yerine kampanya INDIRIM tutarindan hesaplaniyordu.
  - Adam/saat toplami SUM(O17:O24) idi; ilk 7 paketin saati uyarlamaya girmiyordu.
  - Satir indirimi birim fiyat uzerinden (F/100*I) hesaplaniyordu; artik satir toplami uzerinden.
  - Ind. Tutar kolonu KDV haric toplami ve KDV'yi etkilemiyordu; artik yuzde + tutar netten duser.
  - Kampanya secimi indirim yuzdesine bagli degildi.
  - Kar marji kampanya indirimi dusulmeden hesaplaniyordu; artik net (indirimli) tutar uzerinden.
  - Rakip karsilastirmasi bizim saatimizi rakibin sabit 8 saatiyle kiyasliyordu; artik ayni saat x rakip
    ucreti (ayar: rakip_ayni_saat = 0 ile Excel'deki gibi).

Eklenenler:
  - Uyarlama hizmet kirilimi (kurulum, veri aktarimi, egitim, rapor, entegrasyon, canliya gecis, proje yonetimi)
    ve kesif formu kurallari (sube, uretim, gecis programi, IT personeli ... saate etkisi, gerekcesiyle).
  - Seyahat / konaklama, odeme secenekleri (vade farki, kredi karti komisyonu, bize kalan), bakim anlasmasi,
    DIA yillik bakim ve 3 yillik toplam maliyet.
  - Kontroller (zorunlu modul, kullanici siniri, cift modul, indirim / kar sinirlari, eksik bilgi) ve
    satis asistani (rakip, kampanya uygunlugu, gecmis kazanma oranlari, en yuksek indirim, tahmini skor).
Tum tutarlar KDV haric; KDV ayrica hesaplanir.
"""
import math
import re
from datetime import date, datetime

_KATLA = str.maketrans('çğıöşüÇĞİÖŞÜ', 'cgiosuCGIOSU')


def _r(x):
    return round(float(x or 0), 2)


def _f(x, varsayilan=0.0):
    try:
        return float(x)
    except (TypeError, ValueError):
        return varsayilan


def sayi_coz(metin, bos=None):
    """Kullanicinin yazdigi sayi: '7.800' / '7 800,00' / '1.250.000' / '20,5' / '%15' / '₺7.800' -> float.

    Nokta ve virgul birlikteyse sondaki ondalik ayiricidir. Yalniz nokta varsa ve 3'erli gruplar halindeyse
    (7.800, 1.250.000) binlik ayiricidir; degilse (20.5) ondaliktir. Yalniz virgul ondaliktir.
    """
    if isinstance(metin, (int, float)):
        return float(metin)
    s = str(metin if metin is not None else '').strip()
    for x in ('%', 'TL', 'tl', '₺', ' ', '\xa0', '\u202f'):
        s = s.replace(x, '')
    if not s:
        return bos
    if ',' in s and '.' in s:
        s = s.replace('.', '').replace(',', '.') if s.rfind(',') > s.rfind('.') else s.replace(',', '')
    elif ',' in s:
        s = s.replace(',', '.')
    elif re.fullmatch(r'[-+]?\d{1,3}(\.\d{3})+', s):
        s = s.replace('.', '')
    try:
        return float(s)
    except ValueError:
        return bos


def katla(s):
    return str(s or '').replace('İ', 'i').replace('I', 'ı').lower().translate(_KATLA)


def para(x):
    s = '{:,.2f}'.format(float(x or 0)).replace(',', 'X').replace('.', ',').replace('X', '.')
    return s + ' TL'


VARSAYILAN_KURALLAR = {
    'hizmetler': [{'kod': 'kurulum', 'ad': 'Kurulum ve parametre ayarları', 'hesap': 'PAKET_SAAT', 'p1': 0, 'p2': 0,
                   'carpan': 1, 'kapsam': ''}],
    'sorular': [],
    'paket_kampanyalari': [],
    'odeme': [{'ad': 'Peşin', 'tur': 'PESIN', 'taksit': 1, 'indirim': 5, 'musteriye_goster': 1},
              {'ad': '2 Taksit', 'tur': 'TAKSIT', 'taksit': 2, 'musteriye_goster': 1},
              {'ad': '3 Taksit', 'tur': 'TAKSIT', 'taksit': 3, 'musteriye_goster': 1},
              {'ad': '4 Taksit', 'tur': 'TAKSIT', 'taksit': 4, 'musteriye_goster': 1}],
    'bakim': [], 'kampanyalar': [], 'gecmis': [],
}


# ======================================================================
# Kopyala-yapistir ile paket kodu girisi
# ======================================================================
def _kod_anahtari(s):
    """Buyuk/kucuk harf, Turkce karakter ve - _ . bosluk farklarini yok sayar."""
    return re.sub(r'[\s\-_.]+', '', katla(s))


def kod_listesi_coz(metin, paketler):
    """Yapistirilan metindeki paket kodlarini bulur.

    Kabul edilenler: alt alta / virgul / noktali virgul / bosluk ile ayrilmis kodlar,
    'KOD 3' (kod + adet) ve Excel'den kopyalanan satirlar (sekme ile ayrilmis hucreler;
    koddan hemen sonraki hucre tam sayiysa adet sayilir).
    Donus: (bulunan [(kod, adet|None)], bulunamayan [metin], pasif [kod])
    """
    harita = {}
    for kod in paketler:
        harita.setdefault(katla(kod), kod)
        harita.setdefault(_kod_anahtari(kod), kod)

    def bul(parca):
        return harita.get(katla(parca)) or harita.get(_kod_anahtari(parca))

    def adet_mi(parca):
        p = parca.strip()
        return p.isdigit() and 0 < int(p) < 100000

    bulunan, bulunamayan, pasif = [], [], []

    def ekle(kod, adet=None):
        if not paketler[kod].get('aktif', 1):
            if kod not in pasif:
                pasif.append(kod)
            return False
        bulunan.append((kod, adet))
        return True

    for satir in str(metin or '').replace('\r', '\n').split('\n'):
        if not satir.strip():
            continue
        if '\t' in satir:
            hucreler = [h.strip() for h in satir.split('\t')]
            for i, h in enumerate(hucreler):
                kod = bul(h) if h else None
                if kod:
                    sonraki = hucreler[i + 1] if i + 1 < len(hucreler) else ''
                    ekle(kod, int(sonraki) if adet_mi(sonraki) else None)
                    break
            else:
                ilk = next((h for h in hucreler if h), '')
                if ilk:
                    bulunamayan.append(ilk)
            continue
        kod = bul(satir.strip())
        if kod:
            ekle(kod)
            continue
        son_eklenen = False
        parcalar = [x for x in re.split(r'[,;\s]+', satir.strip()) if x]
        i = 0
        while i < len(parcalar):
            # bosluklu kodlar icin en uzun eslesme ("DIA URE STD" -> DIA-URE-STD)
            for j in range(len(parcalar), i, -1):
                kod = bul(''.join(parcalar[i:j]))
                if kod:
                    break
            if kod:
                son_eklenen = ekle(kod)
                i = j
                continue
            parca = parcalar[i]
            if adet_mi(parca) and son_eklenen and bulunan[-1][1] is None:
                bulunan[-1] = (bulunan[-1][0], int(parca))
            else:
                bulunamayan.append(parca)
                son_eklenen = False
            i += 1
    return bulunan, bulunamayan, pasif


# ======================================================================
# Lisans satirlari
# ======================================================================
def paket_tipi_bul(masaustu, mobil, ayar, secim='OTOMATIK'):
    if secim in ('STANDART', 'ENTERPRISE'):
        return secim
    if masaustu > ayar['ent_masaustu_esik'] or mobil > ayar['ent_mobil_esik']:
        return 'ENTERPRISE'
    return 'STANDART'


def birim_fiyat(paket, tip):
    """ENTERPRISE'da Ent. fiyat; Ent. fiyati 0 olan pakette Std. fiyat kullanilir (Excel'deki gibi)."""
    if tip == 'ENTERPRISE' and paket.get('ent_fiyat'):
        return float(paket['ent_fiyat'])
    return float(paket.get('std_fiyat') or 0)


def kullanici_sayisi(paket, masaustu, mobil):
    tur = (paket.get('tur') or '').strip().lower()
    if tur.startswith('mas'):
        return masaustu
    if tur.startswith('mob'):
        return mobil
    return 1


def satir_hesapla(paket, tip, adet, ind_yuzde, kdv, saat=None, ind_tutar=0, kamp=None):
    birim = birim_fiyat(paket, tip)
    kull = float(paket.get('kull_yuzde') or 0)
    toplam = 0.0 if adet <= 0 else birim + (adet - 1) * birim * kull / 100
    # 1) DIA paket kampanyasi (Kamp. Ind. % / Tutar) liste toplamindan duser
    kamp_ind = 0.0
    if kamp and toplam:
        kamp_ind = min(toplam * _f(kamp.get('yuzde')) / 100 + _f(kamp.get('tutar')), toplam)
    # 2) Elle indirim = yuzde + sabit tutar, kampanyali tutar uzerinden; KDV ve tum toplamlar net uzerinden
    kalan = toplam - kamp_ind
    elle_ind = min(kalan * float(ind_yuzde or 0) / 100 + float(ind_tutar or 0), kalan)
    indirim = kamp_ind + elle_ind
    net = toplam - indirim
    tablo_saat = float(paket.get('ent_saat' if tip == 'ENTERPRISE' else 'std_saat') or 0)
    saat = tablo_saat if saat in (None, '') else float(saat)
    return {'kod': paket['kod'], 'ad': paket.get('ad', ''), 'tur': paket.get('tur', ''), 'adet': adet,
            'aciklama': paket.get('aciklama') or '', 'urunler': paket.get('urunler') or '',
            'birim': _r(birim), 'kull_yuzde': kull, 'toplam': _r(toplam), 'ind_yuzde': float(ind_yuzde or 0),
            'ind_tutar': float(ind_tutar or 0), 'kamp_ad': (kamp or {}).get('ad', '') if kamp_ind else '',
            'kamp_indirim': _r(kamp_ind), 'elle_indirim': _r(elle_ind),
            'indirim': _r(indirim), 'net': _r(net), 'kdv': _r(net * kdv / 100), 'dahil': _r(net * (1 + kdv / 100)),
            'saat': saat, 'tablo_saat': tablo_saat}


def paket_kampanyasi_bul(kod, kampanyalar, tarih):
    """Koda uyan, aktif ve tarihi gecerli ilk DIA paket kampanyasi. kodlar: 'M550_*, P12_PRKN_PRO' (* joker)."""
    import fnmatch
    gun = tarih.strftime('%Y-%m-%d') if hasattr(tarih, 'strftime') else str(tarih or '')[:10]
    for k in kampanyalar or []:
        if not int(_f(k.get('aktif'), 1)):
            continue
        if k.get('baslangic') and gun and gun < str(k['baslangic'])[:10]:
            continue
        if k.get('bitis') and gun and gun > str(k['bitis'])[:10]:
            continue
        for desen in re.split(r'[,;\s]+', str(k.get('kodlar') or '')):
            if desen and fnmatch.fnmatchcase(katla(kod), katla(desen)):
                return k
    return None


def dia_ozeti(satirlar, paketler, tip, ybh):
    """DIA uye ekraniyla birebir karsilastirma: Uyelik giris tutari, ek +1 kullanici, yillik bulut hizmet bedeli."""
    ek = {'masaustu': 0.0, 'mobil': 0.0}
    for x in satirlar:
        p = paketler.get(x['kod']) or {}
        tur = (p.get('tur') or '').strip().lower()
        alan = 'masaustu' if tur.startswith('mas') else 'mobil' if tur.startswith('mob') else None
        if alan:
            ek[alan] += birim_fiyat(p, tip) * float(p.get('kull_yuzde') or 0) / 100
    return {'uyelik_giris': _r(sum(x['net'] for x in satirlar)),
            'kampanya_indirimi': _r(sum(x['kamp_indirim'] for x in satirlar)),
            'ek_masaustu': _r(ek['masaustu']), 'ek_mobil': _r(ek['mobil']), 'yillik_bulut': _r(ybh['baz'])}


def _kdvli(net, kdv):
    return {'net': _r(net), 'kdv': _r(net * kdv / 100), 'dahil': _r(net * (1 + kdv / 100))}


def urun_kumesi(paket):
    return {u.strip() for u in (paket.get('urunler') or '').split('-') if u.strip()}


def modul_sayisi(paket):
    return len(urun_kumesi(paket))


# ======================================================================
# Kesif formu + hizmet kirilimi
# ======================================================================
def _evet(x):
    if isinstance(x, bool):
        return x
    return katla(x) in ('evet', 'e', 'true', '1', 'var', 'x')


def _secim_haritasi(metin):
    harita = {}
    for parca in str(metin or '').split(';'):
        if '=' in parca:
            a, b = parca.split('=', 1)
            harita[katla(a.strip())] = _f(b.strip().replace(',', '.'))
    return harita


def soru_etkisi(soru, cevap):
    """Doner: (tetiklendi_mi, sayi, aciklama). sayi: EKLE icin saat, CARP icin katsayi, MIKTAR icin adet."""
    tur, islem = soru.get('tur'), soru.get('islem')
    if cevap in (None, '') or islem == 'BILGI':
        return False, 0, ''
    soru_adi = soru.get('soru') or soru['kod']
    if tur == 'EVET_HAYIR':
        tetik = katla(soru.get('tetik') or 'EVET')
        evet = _evet(cevap)
        if evet != (tetik == 'evet'):
            return False, 0, ''
        return True, _f(soru.get('deger'), 1 if islem.startswith('CARP') else 0), soru_adi + ': ' + (
            'Evet' if evet else 'Hayır')
    if tur == 'SECIM':
        deger = _secim_haritasi(soru.get('deger')).get(katla(cevap), 0)
        return bool(deger), deger, '%s: %s' % (soru_adi, cevap)
    v = _f(str(cevap).replace(',', '.'))
    if v <= 0:
        return False, 0, ''
    if islem == 'MIKTAR':
        return True, v, '%s: %g' % (soru_adi, v)
    if islem == 'CARP_YUZDE':
        return True, 1 + v / 100, '%s: %%%g' % (soru_adi, v)
    if islem == 'CARP':
        return True, _f(soru.get('deger'), 1), '%s: %g' % (soru_adi, v)
    fazla = max(0.0, v - _f(soru.get('esik')))
    return fazla > 0, fazla * _f(soru.get('deger')), '%s: %g' % (soru_adi, v)


def cevap_coz(soru, metin):
    """Metin cevabi (standart cevap, profil) soru turune cevirir: EVET_HAYIR -> bool, SAYI -> float, SECIM -> metin."""
    if metin is None or (isinstance(metin, str) and not metin.strip()):
        return None
    tur = soru.get('tur')
    if tur == 'EVET_HAYIR':
        return metin if isinstance(metin, bool) else _evet(metin)
    if tur == 'SECIM':
        secenekler = [x.strip() for x in str(soru.get('secenekler') or '').split(';') if x.strip()]
        return next((x for x in secenekler if katla(x) == katla(metin)), str(metin).strip())
    return sayi_coz(metin, None)


def standart_cevaplar(sorular):
    """Sorularin 'varsayilan' (standart) cevaplari: yeni teklif bu cevaplarla acilir."""
    return {s['kod']: c for s in sorular for c in [cevap_coz(s, s.get('varsayilan'))] if c is not None}


_ISLEM_ADI = {'EKLE': 'saat ekler', 'CARP': 'saati çarpar', 'CARP_YUZDE': 'saati % artırır',
              'MIKTAR': 'miktar (kayıt / adet) verir', 'BILGI': 'sadece bilgi'}


def soru_etki_metni(soru, hizmet_adlari=None):
    """Sorunun ne yaptigini duz cumleyle anlatir (duzenleyici ve form icin)."""
    hedef = soru.get('hedef') or ''
    hedef_adi = 'tüm hizmetler' if hedef == '*' else (hizmet_adlari or {}).get(hedef, hedef) or '?'
    islem, deger, tur = soru.get('islem'), str(soru.get('deger') or ''), soru.get('tur')
    tetik = 'Evet' if katla(soru.get('tetik') or 'EVET') == 'evet' else 'Hayır'
    if islem == 'BILGI' or not islem:
        return 'Sadece bilgi: hesaba girmez (satış asistanı ve kampanya koşulları kullanır).'
    if tur == 'EVET_HAYIR':
        if islem == 'EKLE':
            return '%s ise %s: +%s saat' % (tetik, hedef_adi, deger)
        if islem.startswith('CARP'):
            return '%s ise %s × %s' % (tetik, hedef_adi, deger)
    if tur == 'SECIM' and islem == 'EKLE':
        return '%s: seçime göre +saat (%s)' % (hedef_adi, deger)
    if islem == 'MIKTAR':
        return '%s: girilen sayı kayıt / adet olarak kullanılır' % hedef_adi
    if islem == 'CARP_YUZDE':
        return '%s: girilen %% kadar artar' % hedef_adi
    if islem == 'CARP':
        return '%s × %s (değer girilirse)' % (hedef_adi, deger)
    esik = _f(soru.get('esik'))
    return '%s: %s üzerindeki her 1 için +%s saat' % (hedef_adi, '%g' % esik, deger)


def kesif_ozeti(sorular, cevaplar):
    """Formun dolulugu: {cevaplanan, toplam, yuzde, bos: [soru metni]} (bilgi sorulari dahil)."""
    cevaplar = cevaplar or {}
    bos = [s.get('soru') or s['kod'] for s in sorular
           if cevaplar.get(s['kod']) in (None, '') and s.get('tur') != 'EVET_HAYIR']
    toplam = len([s for s in sorular if s.get('tur') != 'EVET_HAYIR'])
    cevaplanan = toplam - len(bos)
    return {'cevaplanan': cevaplanan, 'toplam': toplam,
            'yuzde': round(cevaplanan * 100.0 / toplam) if toplam else 100, 'bos': bos}


def hizmet_hesapla(satirlar, kullanici, cevaplar, hizmetler, sorular, elle=None, metinler=None):
    """Uyarlama hizmet kalemleri; her kalemin saati ve gerekcesi."""
    elle = elle or {}
    kalemler = []
    for h in hizmetler:
        kalemler.append({'kod': h['kod'], 'ad': h.get('ad') or h['kod'], 'hesap': h.get('hesap', 'SABIT'),
                         'p1': _f(h.get('p1')), 'p2': _f(h.get('p2')), 'carpan': _f(h.get('carpan'), 1) or 1,
                         'kapsam': h.get('kapsam') or '', 'miktar': 0.0, 'ek': 0.0, 'kats': 1.0, 'nedenler': []})
    sozluk = {k['kod']: k for k in kalemler}
    for s in sorular:
        tetik, sayi, aciklama = soru_etkisi(s, (cevaplar or {}).get(s['kod']))
        if not tetik:
            continue
        hedefler = [k for k in kalemler if k['hesap'] != 'YUZDE'] if s.get('hedef') == '*' else \
            [sozluk[s['hedef']]] if s.get('hedef') in sozluk else []
        for k in hedefler:
            if s['islem'] == 'EKLE':
                k['ek'] += sayi
                k['nedenler'].append('%s → +%g saat' % (aciklama, sayi))
            elif s['islem'] in ('CARP', 'CARP_YUZDE'):
                k['kats'] *= sayi
                k['nedenler'].append('%s → ×%g' % (aciklama, round(sayi, 3)))
            elif s['islem'] == 'MIKTAR':
                k['miktar'] += sayi
                k['nedenler'].append(aciklama)
    paket_saat = sum(x['saat'] for x in satirlar)
    for k in kalemler:
        h, p1, p2 = k['hesap'], k['p1'], k['p2']
        if not satirlar:
            taban = 0.0
        elif h == 'PAKET_SAAT':
            taban = paket_saat
            k['nedenler'].insert(0, 'Paket adam/saatleri: %g saat' % paket_saat)
        elif h == 'KAYIT':
            taban = (k['miktar'] / p1 + p2) if k['miktar'] and p1 else 0.0
            if taban:
                k['nedenler'].insert(0, '%g kayıt ÷ %g kayıt/saat + %g saat hazırlık' % (k['miktar'], p1, p2))
        elif h == 'KULLANICI':
            grup = math.ceil(kullanici / p1) if kullanici and p1 else 0
            taban = grup * p2
            if taban:
                k['nedenler'].insert(0, '%d kullanıcı → %d grup × %g saat' % (kullanici, grup, p2))
        elif h == 'ADET':
            taban = k['miktar'] * p1
            if taban:
                k['nedenler'].insert(0, '%g adet × %g saat' % (k['miktar'], p1))
        elif h == 'SABIT':
            taban = p1
        else:
            taban = 0.0
        k['taban'] = taban
        k['saat'] = 0.0 if h == 'YUZDE' or not satirlar else round((taban + k['ek']) * k['kats'] * k['carpan'], 1)
    diger = sum(k['saat'] for k in kalemler if k['hesap'] != 'YUZDE')
    for k in kalemler:
        if k['hesap'] == 'YUZDE' and satirlar:
            k['taban'] = diger * k['p1'] / 100
            k['saat'] = round(k['taban'] * k['kats'] * k['carpan'], 1)
            k['nedenler'].insert(0, 'Diğer hizmetler %g saat × %%%g' % (diger, k['p1']))
    for k in kalemler:
        if k['carpan'] != 1 and k['saat']:
            k['nedenler'].append('Kalibrasyon çarpanı ×%g (gerçekleşen saatlere göre)' % k['carpan'])
        if elle.get(k['kod']) not in (None, ''):
            k['oneri_saat'] = k['saat']
            k['saat'] = _f(elle[k['kod']])
            k['nedenler'].append('Elle girildi (öneri %g saat)' % k['oneri_saat'])
            k['elle'] = True
    for k in kalemler:      # teklife ozel ad / kapsam (PDF'te bu yazilir)
        m = (metinler or {}).get(k['kod']) or {}
        if (m.get('ad') or '').strip():
            k['ad'], k['ozel_ad'] = m['ad'].strip(), True
        if m.get('kapsam') is not None and str(m.get('kapsam')).strip():
            k['kapsam'], k['ozel_kapsam'] = str(m['kapsam']).strip(), True
    return {'kalemler': kalemler, 'saat': round(sum(k['saat'] for k in kalemler), 1)}


# ======================================================================
# Seyahat, odeme, bakim, TCO
# ======================================================================
def seyahat_hesapla(seyahat, ayar):
    seyahat = seyahat or {}
    if not seyahat.get('sehir_disi'):
        return {'tutar': 0.0, 'detay': ''}
    gun = int(_f(seyahat.get('gun')))
    kisi = max(int(_f(seyahat.get('kisi'), 1)), 1)
    if gun <= 0:
        return {'tutar': 0.0, 'detay': 'Şehir dışı ama yerinde gün girilmedi'}
    ziyaret = math.ceil(gun / max(_f(ayar.get('ziyaret_basina_gun'), 3), 1))
    gece = max(gun - ziyaret, 0)
    yol, konak, gunluk = _f(ayar.get('yol_ucreti')), _f(ayar.get('konaklama_ucreti')), _f(ayar.get('gundelik'))
    tutar = kisi * (ziyaret * yol + gece * konak + gun * gunluk)
    return {'tutar': _r(tutar), 'gun': gun, 'kisi': kisi, 'ziyaret': ziyaret, 'gece': gece,
            'detay': '%d kişi × (%d ziyaret × %s yol + %d gece × %s konaklama + %d gün × %s gündelik)' % (
                kisi, ziyaret, para(yol), gece, para(konak), gun, para(gunluk))}


def odeme_secenekleri(net, kdv, secenekler, maliyet=0, secim=None):
    """Her secenek: musteri ne oder (toplam / taksit), bankaya komisyon, bize kalan, kar.
    PESIN: net × (1 - indirim%). TAKSIT: vade farki = sabit % + aylik % × ortalama vade (ay).
    KREDI_KARTI: vade farki musteriye yansitilan %, komisyon bankanin tahsilattan kestigi %."""
    sonuc = []
    for s in secenekler:
        taksit = max(int(_f(s.get('taksit'), 1)), 1)
        indirim = _f(s.get('indirim'))
        vf = _f(s.get('vade_farki'))
        aylik = _f(s.get('aylik_vade'))
        if aylik:
            vf += aylik * (taksit + 1) / 2.0
        kom = _f(s.get('komisyon'))
        kdvh = net * (1 - indirim / 100) * (1 + vf / 100)
        dahil = kdvh * (1 + kdv / 100)
        kom_tutar = dahil * kom / 100
        bize = kdvh - kom_tutar
        k = kom / 100 * (1 + kdv / 100)
        gereken = (1 / (1 - k) - 1) * 100 if kom and k < 1 else 0
        goster = (s['ad'] in secim) if secim is not None else bool(_f(s.get('musteriye_goster'), 1))
        sonuc.append({'ad': s['ad'], 'tur': s.get('tur', 'TAKSIT'), 'taksit': taksit, 'indirim': indirim,
                      'vade_farki': round(vf, 2), 'komisyon': kom, 'kdv_haric': _r(kdvh), 'dahil': _r(dahil),
                      'taksit_tutari': _r(dahil / taksit), 'komisyon_tutari': _r(kom_tutar), 'bize_kalan': _r(bize),
                      'kar': _r(bize - maliyet) if maliyet else None, 'fark': _r(kdvh - net),
                      'gereken_vade_farki': round(gereken, 2), 'musteriye_goster': goster,
                      'komisyon_acigi': _r(max(kom_tutar - (kdvh - net * (1 - indirim / 100)), 0))})
    return sonuc


def bakim_hesapla(girdi, sonuc_parca, paketler, ayar, bakim_paketleri, sorular):
    """Bakim / destek anlasmasi: yillik destek saati neye gore belirlenir (detay listesinde)."""
    cevaplar = girdi.get('cevaplar') or {}
    kull = int(_f(girdi.get('masaustu'))) + int(_f(girdi.get('mobil')))
    modul = sum(modul_sayisi(paketler.get(s['kod'], {})) for s in sonuc_parca['satirlar'])
    entg = _f(cevaplar.get('entegrasyon_sayisi'))
    sube = max(_f(cevaplar.get('sube_sayisi')) - 1, 0)
    uyarlama_saat = sonuc_parca['hizmet']['saat']
    parcalar = [('Taban', _f(ayar.get('destek_taban_saat'))),
                ('%d kullanıcı × %g' % (kull, _f(ayar.get('destek_kullanici_saat'))),
                 kull * _f(ayar.get('destek_kullanici_saat'))),
                ('%d modül × %g' % (modul, _f(ayar.get('destek_modul_saat'))), modul * _f(ayar.get('destek_modul_saat'))),
                ('%g entegrasyon × %g' % (entg, _f(ayar.get('destek_entegrasyon_saat'))),
                 entg * _f(ayar.get('destek_entegrasyon_saat'))),
                ('%g ek şube × %g' % (sube, _f(ayar.get('destek_sube_saat'))), sube * _f(ayar.get('destek_sube_saat'))),
                ('Uyarlama %g saatin %%%g\'si' % (uyarlama_saat, _f(ayar.get('destek_uyarlama_orani'))),
                 uyarlama_saat * _f(ayar.get('destek_uyarlama_orani')) / 100)]
    saat = sum(p[1] for p in parcalar)
    kats = 1.0
    for s in sorular:
        if s.get('hedef') == '*' and s.get('islem') == 'CARP':
            tetik, sayi, aciklama = soru_etkisi(s, cevaplar.get(s['kod']))
            if tetik:
                kats *= sayi
                parcalar.append((aciklama + ' → ×%g' % sayi, None))
    saat = round(saat * kats, 1)
    if not sonuc_parca['satirlar']:
        saat = 0.0
    saat_ucreti = _f(ayar.get('destek_saat_ucreti')) or _f(ayar.get('saatlik_ucret'))
    gun_ucreti = _f(ayar.get('destek_yerinde_gun_ucreti')) or 8 * saat_ucreti
    esikler = [_f(x) for x in str(ayar.get('destek_esikleri') or '').replace(',', ';').split(';') if x.strip()]
    sirali = sorted(bakim_paketleri, key=lambda b: (_f(b.get('sira')), _f(b.get('saat_carpani'))))
    onerilen_no = sum(1 for e in esikler if saat >= e)
    onerilen = sirali[min(onerilen_no, len(sirali) - 1)]['ad'] if sirali else None
    paketler_sonuc = []
    for b in sirali:
        p_saat = round(saat * _f(b.get('saat_carpani'), 1), 1)
        fiyat = (p_saat * saat_ucreti + _f(b.get('yerinde_gun')) * gun_ucreti) * (1 - _f(b.get('indirim')) / 100)
        paketler_sonuc.append({'ad': b['ad'], 'saat': p_saat, 'yillik': _r(fiyat), 'aylik': _r(fiyat / 12),
                               'yanit': b.get('yanit_suresi') or '', 'yerinde_gun': _f(b.get('yerinde_gun')),
                               'indirim': _f(b.get('indirim')), 'aciklama': b.get('aciklama') or '',
                               'onerilen': b['ad'] == onerilen})
    secili = girdi.get('bakim') if girdi.get('bakim') in [b['ad'] for b in sirali] else None
    return {'yillik_saat': saat, 'saat_ucreti': saat_ucreti, 'detay': parcalar, 'paketler': paketler_sonuc,
            'onerilen': onerilen, 'secili': secili, 'sun': bool(girdi.get('bakim_sun', True)) and bool(sirali)}


def tco_hesapla(net_toplam, ybh_yillik, bakim_yillik, ayar, kdv):
    yillar = max(int(_f(ayar.get('tco_yil'), 3)), 1)
    ybh_bas = int(_f(ayar.get('ybh_baslangic_yili'), 2))
    satirlar = []
    for y in range(1, yillar + 1):
        tek = net_toplam if y == 1 else 0.0
        ybh = ybh_yillik if y >= ybh_bas else 0.0
        toplam = tek + ybh + bakim_yillik
        satirlar.append({'yil': y, 'tek_seferlik': _r(tek), 'ybh': _r(ybh), 'bakim': _r(bakim_yillik),
                         'toplam': _r(toplam), 'dahil': _r(toplam * (1 + kdv / 100))})
    return {'yillar': satirlar, 'toplam': _r(sum(s['toplam'] for s in satirlar)),
            'dahil': _r(sum(s['dahil'] for s in satirlar))}


# ======================================================================
# Kampanya kosullari
# ======================================================================
def kosul_saglanir(kosul, baglam):
    """True / False; 'elle' ve bos olmayan gecersiz kosul -> None (otomatik degerlendirilmez)."""
    kosul = (kosul or '').strip()
    if not kosul:
        return True
    if katla(kosul) == 'elle':
        return None
    for parca in re.split(r'\s+ve\s+', kosul):
        parca = parca.strip()
        if parca.startswith('tarih:'):
            try:
                bas, son = parca[6:].split('..')
                bugun = baglam.get('tarih') or date.today()
                ay_gun = (bugun.month, bugun.day)
                b = tuple(int(x) for x in bas.split('-'))
                s = tuple(int(x) for x in son.split('-'))
                if not (b <= ay_gun <= s if b <= s else (ay_gun >= b or ay_gun <= s)):
                    return False
            except ValueError:
                return None
            continue
        if parca.startswith('soru:'):
            ifade = parca[5:]
            if '=' in ifade:
                kod, secenekler = ifade.split('=', 1)
                cevap = (baglam.get('cevaplar') or {}).get(kod.strip())
                if katla(cevap) not in [katla(x) for x in secenekler.split('|')]:
                    return False
            else:
                cevap = (baglam.get('cevaplar') or {}).get(ifade.strip())
                if not (_evet(cevap) or _f(cevap) > 0):
                    return False
            continue
        m = re.match(r'^(\w+)\s*(>=|<=|>|<|=)\s*([\d.,]+)$', parca)
        if not m:
            return None
        ad, op, deger = m.group(1), m.group(2), _f(m.group(3).replace(',', '.'))
        sol = _f(baglam.get(ad, (baglam.get('cevaplar') or {}).get(ad)))
        if not {'>=': sol >= deger, '<=': sol <= deger, '>': sol > deger, '<': sol < deger,
                '=': sol == deger}[op]:
            return False
    return True


# ======================================================================
# Kontroller
# ======================================================================
_MAX_KULLANICI = re.compile(r'max\.?\s*(\d+)\s*kullan', re.IGNORECASE)


def _paket_kontrolleri(satirlar, paketler, tip, kontrol):
    secili = [(s, paketler.get(s['kod'], {})) for s in satirlar]
    for s, p in secili:
        if not p.get('aktif', 1):
            kontrol('UYARI', s['kod'] + ' güncel DIA fiyat listesinde yok (eski fiyat kullanıldı)')
        if tip == 'ENTERPRISE' and not p.get('ent_fiyat'):
            kontrol('BİLGİ', s['kod'] + ': Enterprise fiyatı yok, Standart fiyat kullanıldı')
        m = _MAX_KULLANICI.search(katla(p.get('aciklama')))
        if m and s['adet'] > int(m.group(1)):
            kontrol('KRİTİK', '%s en fazla %s kullanıcı destekler, teklifte %d kullanıcı var. Daha üst paket seçin.' % (
                s['kod'], m.group(1), s['adet']))
        if 'modul eklenemez' in katla(p.get('aciklama')) and len(secili) > 1:
            kontrol('KRİTİK', s['kod'] + ' paketine modül eklenemez; teklifte başka paket/modül var.')
        gerekli = {x.strip() for x in (p.get('zorunlu') or '').split(',') if x.strip()}
        if gerekli:
            mevcut = set()
            for s2, p2 in secili:
                if s2 is not s:
                    mevcut |= urun_kumesi(p2)
            eksik = sorted(gerekli - mevcut)
            if eksik:
                adaylar = sorted((p3 for p3 in paketler.values() if p3.get('aktif', 1) and
                                  set(eksik) <= urun_kumesi(p3) and p3['kod'] != s['kod']),
                                 key=lambda p3: birim_fiyat(p3, tip))[:3]
                kontrol('KRİTİK', '%s için gerekli modüller teklifte yok: %s.%s' % (
                    s['kod'], ', '.join(eksik), (' Bunları içeren paket: ' + ', '.join(
                        '%s (%s)' % (a['kod'], a.get('ad', '')) for a in adaylar)) if adaylar else ''))
    for i in range(len(secili)):
        for j in range(i + 1, len(secili)):
            a, b = urun_kumesi(secili[i][1]), urun_kumesi(secili[j][1])
            kucuk, buyuk, kucuk_kod, buyuk_kod = (a, b, secili[i][0]['kod'], secili[j][0]['kod']) if len(a) <= len(b) \
                else (b, a, secili[j][0]['kod'], secili[i][0]['kod'])
            if len(kucuk) >= 2 and len(kucuk & buyuk) / len(kucuk) >= 0.5:
                kontrol('UYARI', '%s modüllerinin %%%d\'i %s paketinde zaten var (çift ücretlendirme?): %s' % (
                    kucuk_kod, round(len(kucuk & buyuk) * 100 / len(kucuk)), buyuk_kod,
                    ', '.join(sorted(kucuk & buyuk)[:8])))
            elif secili[i][0]['kod'] == secili[j][0]['kod']:
                kontrol('UYARI', secili[i][0]['kod'] + ' teklifte iki kez var')


def kontroller(girdi, s, ayar, paketler, kurallar):
    liste = []

    def kontrol(seviye, mesaj):
        liste.append({'seviye': seviye, 'mesaj': mesaj})

    for u in s['_uyarilar']:
        kontrol('KRİTİK' if 'listesinde yok' in u else 'UYARI', u)
    _paket_kontrolleri(s['satirlar'], paketler, s['paket_tipi'], kontrol)
    sifir = [x['kod'] for x in s['satirlar'] if not x['tablo_saat'] and not x['saat'] and x['toplam']]
    if sifir and any(h.get('hesap') == 'PAKET_SAAT' for h in kurallar.get('hizmetler', [])):
        kontrol('UYARI', 'Adam/saati 0 olan paket(ler): %s — kurulum saati eksik hesaplanır. Veriler > Paketler / '
                         'Adam-Saat\'ten girin veya A/S kolonuna elle yazın.' % ', '.join(sifir[:8]))

    # DIA'nin paket kampanyasi bizim verdigimiz indirim sayilmaz
    liste_fiyati = s['lisans']['toplam'] - s['lisans'].get('kamp_indirim', 0) + s['uyarlama']['ham_liste'] + s['sabit']['net'] + s['seyahat']['tutar']
    toplam_ind = (1 - s['net_toplam'] / liste_fiyati) * 100 if liste_fiyati else 0
    max_ind = _f(ayar.get('max_toplam_indirim'))
    if max_ind and toplam_ind > max_ind + 0.05:
        kontrol('KRİTİK', 'Toplam indirim %%%.1f, izin verilen en fazla %%%g' % (toplam_ind, max_ind))
    min_marj = _f(ayar.get('min_kar_marji'))
    if min_marj and s['kar']['maliyet'] and s['kar']['marj'] < min_marj - 0.05:
        kontrol('KRİTİK', 'Kâr marjı %%%g, alt sınır %%%g' % (s['kar']['marj'], min_marj))
    ucret = s['uyarlama']['saat_ucreti']
    if not ucret and s['satirlar']:
        kontrol('KRİTİK', 'Saatlik danışman ücreti tanımlı değil (Veriler > Ayarlar > saatlik_ucret)')
    elif _f(ayar.get('min_saat_ucreti')) and ucret < _f(ayar.get('min_saat_ucreti')):
        kontrol('KRİTİK', 'Saatlik ücret %s, taban %s' % (para(ucret), para(ayar['min_saat_ucreti'])))
    if s['lisans']['net'] and s['uyarlama']['net']:
        oran = s['uyarlama']['net'] / s['lisans']['net'] * 100
        if _f(ayar.get('uyarlama_lisans_min')) and oran < _f(ayar['uyarlama_lisans_min']):
            kontrol('UYARI', 'Uyarlama lisansın %%%.0f\'i; alt sınır %%%g. Uyarlama düşük kalmış olabilir (proje riski).' % (
                oran, _f(ayar['uyarlama_lisans_min'])))
        if _f(ayar.get('uyarlama_lisans_max')) and oran > _f(ayar['uyarlama_lisans_max']):
            kontrol('UYARI', 'Uyarlama lisansın %%%.0f\'i; üst sınır %%%g. Müşteri itiraz edebilir, kapsamı açıklayın.' % (
                oran, _f(ayar['uyarlama_lisans_max'])))
    if s['uyarlama']['sinir']:
        kontrol('BİLGİ', 'Uyarlamada %s uygulandı' % s['uyarlama']['sinir'])

    # Kontrol listesi (eksik bilgi)
    mb = girdi.get('musteri_bilgi') or {}
    if not (girdi.get('musteri') or '').strip():
        kontrol('KRİTİK', 'Müşteri adı yok')
    if not (mb.get('yetkili') and (mb.get('telefon') or mb.get('eposta'))):
        kontrol('UYARI', 'Müşteri yetkili adı ve telefon/e-posta eksik')
    if not mb.get('vkn'):
        kontrol('BİLGİ', 'Müşteri VKN/TCKN girilmedi')
    cevaplar = {k: v for k, v in (girdi.get('cevaplar') or {}).items() if v not in (None, '', False)}
    if kurallar['sorular'] and not cevaplar and s['satirlar']:
        kontrol('UYARI', 'Keşif formu doldurulmadı: uyarlama sadece paket saatlerine göre hesaplandı')
    gecis = katla(cevaplar.get('gecis_programi'))
    hizmet = {k['kod']: k for k in s['hizmet']['kalemler']}
    if gecis and gecis != 'yok' and 'veri_aktarimi' in hizmet and not hizmet['veri_aktarimi']['miktar']:
        kontrol('UYARI', 'Geçiş yapılıyor ama aktarılacak kayıt sayısı girilmedi (veri aktarımı saati eksik kalır)')
    tarih = str(ayar.get('fiyat_listesi_tarihi') or '')
    if tarih and _f(ayar.get('fiyat_listesi_max_gun')):
        try:
            gun = (date.today() - datetime.strptime(tarih[:10], '%Y-%m-%d').date()).days
            if gun > _f(ayar['fiyat_listesi_max_gun']):
                kontrol('UYARI', 'Paket fiyat listesi %d gün önce yüklendi; DIA\'nın güncel listesini yükleyin' % gun)
        except ValueError:
            pass
    sec = girdi.get('kampanya') or 'YOK'
    for k in kurallar.get('kampanyalar', []):
        if k['ad'] == sec and sec != 'YOK' and kosul_saglanir(k.get('kosul'), s['_baglam']) is False:
            kontrol('UYARI', '"%s" kampanyasının koşulu sağlanmıyor (%s)' % (sec, k.get('kosul')))
    for o in s['odeme']:
        if o['musteriye_goster'] and o['komisyon'] and o['komisyon_acigi'] > 0:
            kontrol('BİLGİ', '%s: banka komisyonunun %s\'lik kısmını siz üstleniyorsunuz (karşılamak için vade farkı '
                             'en az %%%g olmalı)' % (o['ad'], para(o['komisyon_acigi']), o['gereken_vade_farki']))
    sira = {'KRİTİK': 0, 'UYARI': 1, 'BİLGİ': 2}
    liste.sort(key=lambda x: sira.get(x['seviye'], 3))
    return {'liste': liste, 'kritik': any(x['seviye'] == 'KRİTİK' for x in liste),
            'toplam_indirim_yuzde': round(toplam_ind, 1)}


# ======================================================================
# Satis asistani
# ======================================================================
def satis_asistani(girdi, s, ayar, kurallar):
    oneriler = []
    gerekce = []

    def oner(seviye, baslik, mesaj):
        oneriler.append({'seviye': seviye, 'baslik': baslik, 'mesaj': mesaj})

    skor = 50.0
    r = s['rakip']
    ara = s['ara_toplam']
    if r['konum'] is not None:
        etki = max(min(r['konum'], 60), -40) * 0.5
        skor -= etki
        gerekce.append('Rakip konumu %+g%% → %+d puan' % (r['konum'], -round(etki)))
        if r['konum'] > _f(ayar.get('rakip_pahali_esik'), 15):
            fark = r['biz'] - r['ortalama']
            oner('UYARI', 'Rakiplerden pahalısınız', 'Uyarlamanız rakip ortalamasından %%%g yüksek. Ortalamaya inmek '
                 'için uyarlamada %s indirim gerekir (toplam teklifin %%%.1f\'i). Önce kapsamı (hizmet dökümünü) '
                 'anlatın; fiyat yerine değer konuşun.' % (r['konum'], para(fark), fark / ara * 100 if ara else 0))
        elif r['konum'] < -_f(ayar.get('rakip_ucuz_esik'), 10):
            oner('İYİ', 'Rakiplerden ucuzsunuz', 'Uyarlama rakip ortalamasının %%%g altında. Ek indirim vermeden '
                 'kapatmayı deneyin; gerekirse peşin indirimi veya bakım anlaşmasını öne çıkarın.' % abs(r['konum']))
        else:
            oner('BİLGİ', 'Fiyat rakiplerle dengeli', 'Uyarlama rakip ortalamasına yakın (%+g%%).' % r['konum'])
    else:
        oner('BİLGİ', 'Rakip verisi yok', 'Veriler > Rakipler tablosuna saatlik ücretleri girerseniz '
                                          'karşılaştırma ve öneri yapılır.')

    # En yuksek indirim
    kar = s['kar']
    max_ind = _f(ayar.get('max_toplam_indirim'))
    sinirlar = []
    if kar['maliyet'] and _f(ayar.get('min_kar_marji')) < 100:
        net_min = kar['maliyet'] / (1 - _f(ayar.get('min_kar_marji')) / 100)
        sinirlar.append(((1 - net_min / ara) * 100 if ara else 0, 'kâr marjı %%%g altına düşmesin' % _f(
            ayar.get('min_kar_marji'))))
    if max_ind:
        mevcut_satir_ind = s['kontrol']['toplam_indirim_yuzde'] - s['kampanya']['yuzde']
        sinirlar.append((max_ind - max(mevcut_satir_ind, 0), 'toplam indirim %%%g\'yi geçmesin' % max_ind))
    if sinirlar:
        deger, neden = min(sinirlar, key=lambda x: x[0])
        deger = max(deger, 0)
        oner('KRİTİK' if s['kampanya']['yuzde'] > deger + 0.05 else 'BİLGİ', 'Verilebilecek en yüksek kampanya indirimi',
             '%%%.1f (%s). Şu an %%%g.' % (deger, neden, s['kampanya']['yuzde']))
        s['max_kampanya_indirimi'] = round(deger, 1)

    # Kampanya uygunlugu
    baglam = s['_baglam']
    uygun = []
    for k in kurallar.get('kampanyalar', []):
        if k['ad'] == 'YOK' or not _f(k.get('yuzde')):
            continue
        if kosul_saglanir(k.get('kosul'), baglam):
            uygun.append(k)
    uygun.sort(key=lambda k: -_f(k['yuzde']))
    secili = girdi.get('kampanya') or 'YOK'
    if uygun and uygun[0]['ad'] != secili:
        oner('ÖNERİ', 'Uygun kampanya', 'Müşteri şu kampanyalara uygun: %s. Rekabet gerekiyorsa "%s" (%%%g) '
             'kullanılabilir.' % (', '.join('%s (%%%g)' % (k['ad'], _f(k['yuzde'])) for k in uygun), uygun[0]['ad'],
                                   _f(uygun[0]['yuzde'])))
    if secili != 'YOK' and any(k['ad'] == secili for k in uygun):
        skor += 5
        gerekce.append('Uygun kampanya uygulandı → +5')

    # Gecmis teklifler
    tum_gecmis = kurallar.get('gecmis') or []
    gecmis = [g for g in tum_gecmis if g['sonuc'] in ('KAZANDI', 'KAYBETTİ') and g['no'] != girdi.get('teklif_no')]
    if len(gecmis) >= 3:
        kaz = [g for g in gecmis if g['sonuc'] == 'KAZANDI']
        oran = len(kaz) * 100.0 / len(gecmis)
        benzer = [g for g in gecmis if g['net'] and 0.6 <= s['net_toplam'] / g['net'] <= 1.4] if s['net_toplam'] else []
        b_kaz = [g for g in benzer if g['sonuc'] == 'KAZANDI']
        mesaj = 'Sonuçlanan %d teklifin %d\'i kazanıldı (%%%d).' % (len(gecmis), len(kaz), round(oran))
        if len(benzer) >= 3:
            b_oran = len(b_kaz) * 100.0 / len(benzer)
            mesaj += ' Benzer büyüklükteki %d teklifte kazanma %%%d.' % (len(benzer), round(b_oran))
            skor = skor * 0.5 + b_oran * 0.5
            gerekce.append('Benzer büyüklükte kazanma oranı %%%d (ağırlık %%50)' % round(b_oran))
        elif len(gecmis) >= 5:
            skor = skor * 0.7 + oran * 0.3
            gerekce.append('Genel kazanma oranı %%%d (ağırlık %%30)' % round(oran))
        ort_kaz = sum(g['indirim'] for g in kaz) / len(kaz) if kaz else None
        kay = [g for g in gecmis if g['sonuc'] == 'KAYBETTİ']
        if ort_kaz is not None:
            mesaj += ' Kazanılanlarda ortalama indirim %%%.1f' % ort_kaz
            if kay:
                mesaj += ', kaybedilenlerde %%%.1f.' % (sum(g['indirim'] for g in kay) / len(kay))
            if s['kontrol']['toplam_indirim_yuzde'] + 3 < ort_kaz:
                mesaj += ' Bu teklifteki indirim (%%%.1f) kazanılanların ortalamasının altında.' % \
                         s['kontrol']['toplam_indirim_yuzde']
        fiyat_red = sum(1 for g in kay if 'fiyat' in katla(g['red']) or 'pahal' in katla(g['red']))
        if kay and fiyat_red * 2 >= len(kay):
            mesaj += ' Kaybedilenlerin çoğunun nedeni fiyat.'
        oner('BİLGİ', 'Geçmiş tekliflerden', mesaj)

    # Odeme
    od = [o for o in s['odeme'] if o['musteriye_goster']]
    pesin = next((o for o in od if o['tur'] == 'PESIN' and o['indirim']), None)
    if pesin:
        oner('ÖNERİ', 'Peşin ödeme', 'Peşin ödemede müşteri %s kazanır (%%%g). Nakit akışı için önce bunu önerin.' % (
            para(s['genel_toplam'] - pesin['dahil']), pesin['indirim']))
    acik = [o for o in s['odeme'] if o['komisyon'] and o['komisyon_acigi'] > 0]
    if acik:
        o = max(acik, key=lambda x: x['komisyon_acigi'])
        oner('UYARI', 'Kredi kartı komisyonu', '%s seçeneğinde komisyonun %s\'si size kalıyor; vade farkını en az '
             '%%%g yapın ya da bu seçeneği müşteriye göstermeyin.' % (o['ad'], para(o['komisyon_acigi']),
                                                                       o['gereken_vade_farki']))

    # Uyarlama orani
    if s['lisans']['net'] and s['uyarlama']['net'] / s['lisans']['net'] < 0.15 and s['satirlar']:
        oner('UYARI', 'Uyarlama düşük olabilir', 'Uyarlama lisansın %%%d\'i. Keşif formunu doldurduğunuzdan emin '
             'olun; eksik kalan hizmet projede zarar ettirir.' % round(s['uyarlama']['net'] / s['lisans']['net'] * 100))

    # Bakim
    b = s['bakim']
    if b['onerilen'] and not b['secili']:
        p = next(x for x in b['paketler'] if x['ad'] == b['onerilen'])
        oner('ÖNERİ', 'Bakım anlaşması', 'Önerilen paket: %s (yıllık %g saat, %s/yıl, aylık %s). Teklife eklemek '
             'yıllık düzenli gelir sağlar ve müşteriye güven verir.' % (p['ad'], p['saat'], para(p['yillik']),
                                                                          para(p['aylik'])))

    # Ihtiyac - paket uyumu (kesif cevaplari ile secilen paketler)
    cevaplar = girdi.get('cevaplar') or {}
    ihtiyac = _ihtiyac_paket_uyumu(s['satirlar'], cevaplar, baglam)
    for baslik, mesaj in ihtiyac:
        oner('UYARI', baslik, mesaj)
    if ihtiyac:
        skor -= 5
        gerekce.append('İhtiyaç-paket uyumsuzluğu → -5')

    # Kesif dolulugu ve musteri bilgisi
    ko = kesif_ozeti(kurallar.get('sorular') or [], cevaplar)
    if kurallar.get('sorular'):
        if ko['yuzde'] < 60:
            skor -= 8
            gerekce.append('Keşif formu %%%d dolu → -8 (tahmin belirsiz)' % ko['yuzde'])
            oner('UYARI', 'Keşif eksik', 'Form %%%d dolu. Müşteriye sorun: %s.' % (
                ko['yuzde'], ', '.join(ko['bos'][:4])))
        elif ko['yuzde'] >= 90:
            skor += 3
            gerekce.append('Keşif formu eksiksiz → +3')
    mb = girdi.get('musteri_bilgi') or {}
    if not (mb.get('yetkili') and (mb.get('telefon') or mb.get('eposta'))):
        skor -= 3
        gerekce.append('Karar verici / iletişim bilgisi yok → -3')
    if (girdi.get('bakim') or ''):
        skor += 3
        gerekce.append('Bakım anlaşması teklifte → +3')

    if s['kontrol']['kritik']:
        skor -= 10
        gerekce.append('Kritik kontrol hatası → -10')
    skor = int(round(max(5, min(95, skor))))
    seviye = 'Yüksek' if skor >= 65 else 'Orta' if skor >= 40 else 'Düşük'
    sira = {'KRİTİK': 0, 'UYARI': 1, 'ÖNERİ': 2, 'İYİ': 3, 'BİLGİ': 4}
    oneriler.sort(key=lambda x: sira.get(x['seviye'], 5))
    pazarlik = _pazarlik_araligi(s, ayar)
    senaryolar = kurallar.get('senaryolar') or []
    yer = _yer_tutucular(s, ayar, pazarlik, girdi)
    asama = girdi.get('asama') or ('Teklif verildi' if girdi.get('teklif_no') else 'Fırsat')
    return {'skor': skor, 'seviye': seviye, 'gerekce': gerekce, 'oneriler': oneriler,
            'pazarlik': pazarlik, 'taviz': _taviz_sirasi(s, ayar, pazarlik, senaryolar),
            'itirazlar': _itiraz_cevaplari(s, ayar, senaryolar, yer), 'takip': _takip_plani(girdi, ayar),
            'hazirlik': _hazirlik_listesi(girdi, s, ko, kurallar), 'kesif': ko,
            'asama': asama, 'sonraki_adim': _sonraki_adim(asama, senaryolar, yer, girdi, s, ko),
            'ogrenilen': gecmisten_ogren(tum_gecmis, girdi, s)}


ASAMA_ADIMLARI = {   # senaryo tablosu bossa
    'Fırsat': 'Karar vericiyi, bütçeyi ve karar tarihini öğrenin.\nKeşif görüşmesi / demo tarihi alın.',
    'Keşif': 'Keşif formunu %80+ doldurun.\nDemo yapın ve teklif tarihini netleştirin.',
    'Teklif verildi': '2 gün içinde arayın; sorularını alın; karar tarihini öğrenin.',
    'Müzakere': 'Önce kapsam ve değer; taviz sırasını izleyin. Hedef {hedef}, taban {taban}.',
    'Sözlü onay': 'Sözleşme ve ödeme planını gönderin; proje başlangıç tarihini belirleyin.',
    'Kazandı': 'Proje başlangıcı; 1 ay sonra memnuniyet görüşmesi; bakım anlaşması ve referans.',
    'Kaybetti': 'Kayıp nedenini kaydedin; 3-6 ay sonra tekrar arayın.',
}


def _yer_tutucular(s, ayar, pazarlik, girdi):
    from datetime import timedelta
    u, r = s['uyarlama'], s['rakip']
    gun = float(ayar.get('is_gunu_saat') or 8) or 8
    kalemler = sorted((k for k in s['hizmet']['kalemler'] if k['saat']), key=lambda k: -k['saat'])[:3]
    b = s.get('bakim') or {}
    secili = next((p for p in b.get('paketler', []) if p['ad'] == (b.get('secili') or b.get('onerilen'))), None)
    return {'dokum': ', '.join('%s %g saat' % (k['ad'], k['saat']) for k in kalemler) or '-',
            'saat': ('%g' % u['saat']).replace('.', ','), 'gun': ('%g' % round(u['saat'] / gun, 1)).replace('.', ','),
            'rakip_ort': para(r['ortalama']) if r.get('ortalama') else '-', 'biz': para(r.get('biz') or 0),
            'gecerlilik': (_tarih(girdi.get('tarih')) + timedelta(days=int(_f(ayar.get('gecerlilik_gun'), 15)))
                           ).strftime('%d.%m.%Y'),
            'taban': para(pazarlik['taban']) if pazarlik.get('taban') else 'tanımlı değil',
            'hedef': para(pazarlik['hedef']),
            'bakim': ('%s, aylık %s' % (secili['ad'], para(secili['aylik']))) if secili else 'bakım paketi'}


def _doldur(metin, yer):
    return re.sub(r'\{(\w+)\}', lambda m: str(yer.get(m.group(1), m.group(0))), metin or '')


def _sonraki_adim(asama, senaryolar, yer, girdi, s, ko):
    """Asamaya gore yapilacaklar (Satis Senaryolari tablosundan) + teklifin durumuna gore ekler."""
    metin = next((x['metin'] for x in senaryolar if x.get('tur') == 'ASAMA' and x['baslik'] == asama), None)
    adimlar = [x.strip() for x in _doldur(metin if metin is not None else ASAMA_ADIMLARI.get(asama, ''),
                                          yer).split('\n') if x.strip()]
    if asama in ('Fırsat', 'Keşif', 'Teklif verildi', 'Müzakere'):
        mb = girdi.get('musteri_bilgi') or {}
        if not mb.get('yetkili'):
            adimlar.insert(0, 'Karar verici (yetkili) bilgisi yok: kimin onaylayacağını öğrenin.')
        if ko['toplam'] and ko['yuzde'] < 80:
            adimlar.append('Keşif formu %%%d dolu: boş sorular → %s' % (ko['yuzde'], ', '.join(ko['bos'][:3])))
        if s['kontrol']['kritik']:
            adimlar.insert(0, 'Teklifte kritik hata var; göndermeden önce düzeltin (Kontroller).')
    return adimlar


def _oran(liste):
    kaz = sum(1 for g in liste if g['sonuc'] == 'KAZANDI')
    return (kaz * 100.0 / len(liste)) if liste else None


def gecmisten_ogren(gecmis, girdi, s):
    """Gecmis tekliflerden bu teklif icin cikarimlar (en az 3 ornek olan gruplar)."""
    from datetime import datetime as _dt
    sonuc = []
    biten = [g for g in gecmis if g['sonuc'] in ('KAZANDI', 'KAYBETTİ') and g['no'] != girdi.get('teklif_no')]
    # Ayni musteri
    musteri = katla(girdi.get('musteri') or '').strip()
    if musteri:
        once = [g for g in gecmis if katla(g.get('musteri') or '').strip() == musteri and g['no'] != girdi.get('teklif_no')]
        if once:
            son = max(once, key=lambda g: g.get('tarih') or '')
            sonuc.append('Bu müşteriye daha önce %d teklif verilmiş; sonuncusu %s (%s, %s%s).' % (
                len(once), son['no'], son.get('tarih') or '-', son['sonuc'],
                (': ' + son['red']) if son.get('red') else ''))
    if len(biten) < 3:
        if not sonuc:
            sonuc.append('Henüz yeterli sonuçlanmış teklif yok (en az 3). Teklifleri Geçmiş\'te KAZANDI / KAYBETTİ '
                         'olarak işaretledikçe asistan sizin verinizden öğrenir.')
        return sonuc
    # Indirim bandi
    bantlar = [(0, 5), (5, 10), (10, 15), (15, 101)]
    sonuclar = []
    for alt, ust in bantlar:
        grup = [g for g in biten if alt <= g['indirim'] < ust]
        if len(grup) >= 3:
            sonuclar.append((alt, ust, _oran(grup), len(grup)))
    if len(sonuclar) >= 2:
        en_iyi = max(sonuclar, key=lambda x: x[2])
        mevcut = s['kontrol']['toplam_indirim_yuzde']
        metin = 'Kazanma oranı en yüksek indirim aralığı %%%d-%s (%%%d, %d teklif); bu teklifte %%%s.' % (
            en_iyi[0], '%%%d' % en_iyi[1] if en_iyi[1] <= 100 else 'üstü', round(en_iyi[2]), en_iyi[3],
            ('%g' % round(mevcut, 1)).replace('.', ','))
        if mevcut > en_iyi[1]:
            metin += ' Daha fazla indirim kazanma şansını artırmıyor görünüyor.'
        sonuc.append(metin)
    # Bakim dahil olanlar
    b_var = [g for g in biten if g.get('bakim')]
    b_yok = [g for g in biten if not g.get('bakim')]
    if len(b_var) >= 3 and len(b_yok) >= 3:
        sonuc.append('Bakım anlaşması eklenen tekliflerde kazanma %%%d, eklenmeyenlerde %%%d.' % (
            round(_oran(b_var)), round(_oran(b_yok))))
    # Paket tipi
    tip = [g for g in biten if g.get('paket_tipi') == s['paket_tipi']]
    if len(tip) >= 3:
        sonuc.append('%s tekliflerde kazanma oranı %%%d (%d teklif).' % (s['paket_tipi'], round(_oran(tip)), len(tip)))
    # Karar suresi
    sureler = []
    for g in biten:
        if g['sonuc'] == 'KAZANDI' and g.get('tarih') and g.get('sonuc_tarihi'):
            try:
                sureler.append((_dt.strptime(g['sonuc_tarihi'][:10], '%Y-%m-%d') -
                                _dt.strptime(g['tarih'][:10], '%Y-%m-%d')).days)
            except ValueError:
                pass
    if len(sureler) >= 3:
        ort = sum(sureler) / len(sureler)
        metin = 'Kazanılan teklifler ortalama %d günde kapandı.' % round(ort)
        if girdi.get('tarih'):
            gecen = (_dt.now() - _dt.strptime(girdi['tarih'][:10], '%Y-%m-%d')).days
            if gecen > ort * 1.5:
                metin += ' Bu teklif %d gündür açık: karar sürecini yeniden sorun, gerekirse yöneticinizi dahil edin.' \
                         % gecen
        sonuc.append(metin)
    # Kayip nedenleri
    nedenler = {}
    for g in biten:
        if g['sonuc'] == 'KAYBETTİ' and g.get('red'):
            n = katla(g['red']).split(' ')[0]
            nedenler[n] = nedenler.get(n, 0) + 1
    if nedenler:
        n, sayi = max(nedenler.items(), key=lambda x: x[1])
        sonuc.append('En sık kayıp nedeni: "%s" (%d kez). Bu itiraza önceden hazırlıklı olun.' % (n, sayi))
    return sonuc


# Kesif cevabi 'evet' ise teklifte bu urun / kod parcalarindan biri olmali
IHTIYAC_PAKET = [
    ('e_donusum', ('efatura', 'earsiv', 'edefter', 'e_fatura', 'gib'), 'e-Dönüşüm paketi yok',
     'Keşifte e-Fatura / e-Arşiv / e-Defter kurulacak denmiş ama teklifte e-dönüşüm paketi yok.'),
    ('uretim', ('ure', 'uretim', 'mrp'), 'Üretim paketi yok',
     'Keşifte üretim yapıyor denmiş ama teklifte üretim modülü yok.'),
]


def _ihtiyac_paket_uyumu(satirlar, cevaplar, baglam):
    sonuc = []
    parcalar = set()
    for x in satirlar:
        parcalar.update(p for p in re.split(r'[-_\s,]+', katla(x['kod'] + ' ' + (x.get('urunler') or ''))) if p)
    metin = ' '.join(katla(x['kod'] + ' ' + (x.get('urunler') or '')) for x in satirlar)
    for kod, aranan, baslik, mesaj in IHTIYAC_PAKET:
        if _evet(cevaplar.get(kod)) and satirlar and not any(a in parcalar or (len(a) >= 5 and a in metin) for a in aranan):
            sonuc.append((baslik, mesaj + ' Paketi ekleyin ya da müşteriye ayrıca teklif edileceğini yazın.'))
    if baglam.get('mobil') and satirlar and not any(katla(x.get('tur') or '').startswith('mob') for x in satirlar):
        sonuc.append(('Mobil paket yok', '%d mobil kullanıcı girilmiş ama teklifte mobil paket yok.' % baglam['mobil']))
    return sonuc


def _pazarlik_araligi(s, ayar):
    """Liste (indirimsiz) -> teklif -> hedef -> taban (inilebilecek en dusuk net). Tutarlar KDV haric."""
    liste = s['lisans']['toplam'] - s['lisans'].get('kamp_indirim', 0) + s['uyarlama']['ham_liste'] + \
        s['sabit']['net'] + s['seyahat']['tutar']
    teklif = s['net_toplam']
    adaylar, neden = [], []
    max_ind = _f(ayar.get('max_toplam_indirim'))
    if max_ind:
        adaylar.append(liste * (1 - max_ind / 100))
        neden.append('en fazla %%%g indirim' % max_ind)
    if s['kar']['maliyet'] and _f(ayar.get('min_kar_marji')) < 100:
        adaylar.append(s['kar']['maliyet'] / (1 - _f(ayar.get('min_kar_marji')) / 100))
        neden.append('kâr marjı en az %%%g' % _f(ayar.get('min_kar_marji')))
    taban = max(adaylar) if adaylar else None
    hedef = (teklif + taban) / 2 if taban is not None and taban < teklif else teklif
    return {'liste': _r(liste), 'teklif': _r(teklif), 'hedef': _r(hedef), 'taban': _r(taban) if taban else None,
            'pay': _r(teklif - taban) if taban is not None else None, 'taban_nedeni': ', '.join(neden)}


def _taviz_sirasi(s, ayar, pazarlik, senaryolar=None):
    """Fiyat indirmeden once verilebilecek tavizler (Satis Senaryolari > TAVIZ), en son fiyat."""
    adimlar = [x['metin'] for x in (senaryolar or []) if x.get('tur') == 'TAVIZ']
    if not adimlar:
        ucret = s['uyarlama']['saat_ucreti']
        if ucret:
            adimlar.append('Ek eğitim saati hediye (ör. 4 saat ≈ %s maliyet, fiyat düşmeden değer artar)' %
                           para(4 * ucret))
        if (s.get('bakim') or {}).get('onerilen'):
            adimlar.append('Bakım anlaşmasının ilk 3 ayı ücretsiz veya 1. yıl %10 indirimli')
        od = [o for o in s['odeme'] if o['tur'] != 'PESIN' and not o['vade_farki'] and o['taksit'] > 1]
        if od:
            adimlar.append('Vade farksız taksit (%s) — fiyat aynı kalır' % ', '.join(o['ad'] for o in od[:2]))
    if pazarlik.get('pay'):
        adimlar.append('Son adım fiyat: en fazla %s (taban %s; %s)' % (
            para(pazarlik['pay']), para(pazarlik['taban']), pazarlik['taban_nedeni']))
    else:
        adimlar.append('Son adım fiyat: Ayarlar\'da en fazla indirim / maliyet girilirse inilebilecek taban hesaplanır')
    return adimlar


def _itiraz_cevaplari(s, ayar, senaryolar=None, yer=None):
    itirazlar = [x for x in (senaryolar or []) if x.get('tur') == 'ITIRAZ']
    if itirazlar:
        return [(x['baslik'], _doldur(x['metin'], yer or {})) for x in itirazlar
                if '{rakip_ort}' not in x['metin'] or s['rakip'].get('ortalama')]
    u = s['uyarlama']
    r = s['rakip']
    gun = float(ayar.get('is_gunu_saat') or 8) or 8
    kalemler = sorted((k for k in s['hizmet']['kalemler'] if k['saat']), key=lambda k: -k['saat'])[:3]
    dokum = ', '.join('%s %g saat' % (k['ad'], k['saat']) for k in kalemler)
    sonuc = [('"Fiyat yüksek"', 'Hizmet dökümünü gösterin: %s (toplam %g saat ≈ %g iş günü). Rakip teklifinde '
                                'kaç saat kurulum, eğitim ve canlıya geçiş desteği olduğunu sorun.' % (
                                    dokum or '-', u['saat'], round(u['saat'] / gun, 1)))]
    if r.get('ortalama'):
        sonuc.append(('"Rakip daha ucuz"', 'Aynı saat için rakip ortalaması %s, bizimki %s. Fark saat ücretinde değil '
                                           'kapsamda olabilir; rakibin saat ve kapsamını yazılı isteyin.' % (
                                               para(r['ortalama']), para(r['biz']))))
    return sonuc


def _takip_plani(girdi, ayar):
    t = _tarih(girdi.get('tarih'))
    from datetime import timedelta
    gecerlilik = int(_f(ayar.get('gecerlilik_gun'), 15))
    plan = [(t + timedelta(days=2), 'İlk takip: teklif ulaştı mı, sorular?'),
            (t + timedelta(days=7), 'İkinci takip: karar vericiyle görüşme / demo')]
    if gecerlilik > 3:
        plan.append((t + timedelta(days=gecerlilik - 2), 'Son hatırlatma: geçerlilik bitmeden 2 gün önce'))
    return [(d.strftime('%d.%m.%Y'), m) for d, m in plan]


def _hazirlik_listesi(girdi, s, ko, kurallar):
    mb = girdi.get('musteri_bilgi') or {}
    od = [o for o in s['odeme'] if o['musteriye_goster']]
    return [('Müşteri adı', bool(girdi.get('musteri'))),
            ('Karar verici (yetkili)', bool(mb.get('yetkili'))),
            ('Telefon veya e-posta', bool(mb.get('telefon') or mb.get('eposta'))),
            ('Keşif formu %80+ dolu', ko['yuzde'] >= 80 if kurallar.get('sorular') else True),
            ('Kritik kontrol yok', not s['kontrol']['kritik']),
            ('Ödeme seçeneği sunuldu', bool(od)),
            ('Bakım anlaşması önerildi', bool((s.get('bakim') or {}).get('sun'))),
            ('Rakip verisi var', s['rakip'].get('ortalama') is not None)]


def satis_brifingi(girdi, s):
    """Satis asistaninin duz metin ozeti (panoya kopyalanip CRM / mesaj olarak paylasilir)."""
    a = s['asistan']
    p = a['pazarlik']
    satirlar = ['SATIŞ BRİFİNGİ - %s (%s)' % (girdi.get('musteri') or '-', girdi.get('teklif_no') or 'kaydedilmedi'),
                'Kazanma ihtimali: %%%d (%s) | Aşama: %s' % (a['skor'], a['seviye'], a.get('asama', '-')),
                'Net teklif: %s + KDV | Genel toplam: %s' % (para(s['net_toplam']), para(s['genel_toplam'])),
                'Pazarlık: liste %s, teklif %s, hedef %s, taban %s' % (
                    para(p['liste']), para(p['teklif']), para(p['hedef']), para(p['taban']) if p['taban'] else '-'),
                '', 'Sonraki adımlar:'] + ['- ' + x for x in a.get('sonraki_adim', [])] + ['', 'Öneriler:']
    satirlar += ['- [%s] %s: %s' % (o['seviye'], o['baslik'], o['mesaj']) for o in a['oneriler']]
    satirlar += ['', 'Taviz sırası:'] + ['%d. %s' % (i, x) for i, x in enumerate(a['taviz'], 1)]
    satirlar += ['', 'İtirazlar:'] + ['- %s → %s' % x for x in a['itirazlar']]
    satirlar += ['', 'Takip:'] + ['- %s: %s' % x for x in a['takip']]
    if a.get('ogrenilen'):
        satirlar += ['', 'Geçmişten öğrenilenler:'] + ['- ' + x for x in a['ogrenilen']]
    eksik = [m for m, tamam in a['hazirlik'] if not tamam]
    if eksik:
        satirlar += ['', 'Eksikler: ' + ', '.join(eksik)]
    return '\n'.join(satirlar)


# ======================================================================
# Ana hesap
# ======================================================================
def hesapla(girdi, paketler, ayar, sabit_fiyatlar=None, rakipler=None, kurallar=None):
    """girdi: {musteri, musteri_bilgi{yetkili, telefon, eposta, vkn, vergi_dairesi, adres}, masaustu, mobil,
               paket_tipi: OTOMATIK|STANDART|ENTERPRISE,
               kalemler: [{kod, adet (None = otomatik), ind_yuzde, ind_tutar, saat (None = ADAM SAAT tablosu)}],
               cevaplar: {kesif_kodu: cevap}, hizmet_elle: {hizmet_kodu: saat}, saatlik_ucret (None = ayar),
               seyahat: {sehir_disi, gun, kisi}, sabitler: [ad], kampanya, kampanya_yuzde, maliyet,
               odeme_secimi: [ad] (None = tablodaki 'musteriye_goster'), bakim: ad, bakim_sun: bool}
    kurallar: TeklifVeri.kurallar() (None = sadece paket saatleri, eski davranis)"""
    kurallar = dict(VARSAYILAN_KURALLAR, **(kurallar or {}))
    sabit_fiyatlar = sabit_fiyatlar or {}
    kdv = float(ayar['kdv'])
    masaustu = int(_f(girdi.get('masaustu')))
    mobil = int(_f(girdi.get('mobil')))
    tip = paket_tipi_bul(masaustu, mobil, ayar, girdi.get('paket_tipi', 'OTOMATIK'))
    uyarilar = []

    satirlar = []
    for sira, k in enumerate(girdi.get('kalemler', [])):
        p = paketler.get(k.get('kod'))
        if not p:
            uyarilar.append('Paket listesinde yok: ' + str(k.get('kod')))
            continue
        adet = k.get('adet')
        adet = kullanici_sayisi(p, masaustu, mobil) if adet in (None, '') else int(adet)
        kamp = None if k.get('kamp_kapali') else paket_kampanyasi_bul(p['kod'], kurallar.get('paket_kampanyalari'),
                                                                       _tarih(girdi.get('tarih')))
        satir = satir_hesapla(p, tip, adet, k.get('ind_yuzde', 0), kdv, k.get('saat'), k.get('ind_tutar', 0), kamp)
        satir['sira'] = sira     # formdaki kalem sirasi (bilinmeyen kod atlaninca kaymasin)
        satirlar.append(satir)

    lisans_toplam = sum(s['toplam'] for s in satirlar)
    lisans_indirim = sum(s['indirim'] for s in satirlar)
    lisans = dict(_kdvli(sum(s['net'] for s in satirlar), kdv), toplam=_r(lisans_toplam), indirim=_r(lisans_indirim),
                  kamp_indirim=_r(sum(s['kamp_indirim'] for s in satirlar)))

    hizmet = hizmet_hesapla(satirlar, masaustu + mobil, girdi.get('cevaplar') or {}, kurallar['hizmetler'],
                            kurallar['sorular'], girdi.get('hizmet_elle'), girdi.get('hizmet_metin'))
    saat = hizmet['saat']
    ucret = _f(girdi.get('saatlik_ucret')) if girdi.get('saatlik_ucret') not in (None, '') else \
        float(ayar['saatlik_ucret'])
    lisans_payi = lisans['net'] * float(ayar['lisans_yuzdesi']) / 100
    saat_bedeli = saat * ucret
    ham = lisans_payi + saat_bedeli
    uyarlama_net, sinir = ham, ''
    if satirlar and ham < float(ayar['min_uyarlama']):
        uyarlama_net, sinir = float(ayar['min_uyarlama']), 'ALT SINIR'
    if float(ayar['max_uyarlama'] or 0) > 0 and uyarlama_net > float(ayar['max_uyarlama']):
        uyarlama_net, sinir = float(ayar['max_uyarlama']), 'ÜST SINIR'
    for k in hizmet['kalemler']:
        k['tutar'] = _r(k['saat'] * ucret)
    uyarlama = dict(_kdvli(uyarlama_net, kdv), saat=saat, lisans_payi=_r(lisans_payi), saat_ucreti=ucret,
                    saat_bedeli=_r(saat_bedeli), ham=_r(ham), sinir=sinir,
                    ham_liste=_r(lisans_payi + saat * float(ayar['saatlik_ucret'])))

    sabit_kalemler = []
    for ad in girdi.get('sabitler', []):
        if ad in sabit_fiyatlar:
            sabit_kalemler.append({'ad': ad, 'net': _r(sabit_fiyatlar[ad])})
        else:
            uyarilar.append('Sabit kalem tanımlı değil: ' + ad)
    sabit = dict(_kdvli(sum(s['net'] for s in sabit_kalemler), kdv), kalemler=sabit_kalemler)
    seyahat = seyahat_hesapla(girdi.get('seyahat'), ayar)

    ara = lisans['net'] + uyarlama['net'] + sabit['net'] + seyahat['tutar']
    kampanya_yuzde = float(girdi.get('kampanya_yuzde') or 0)
    kampanya_tutar = ara * kampanya_yuzde / 100
    net_toplam = ara - kampanya_tutar
    genel = _kdvli(net_toplam, kdv)

    maliyet = float(girdi.get('maliyet') or 0)
    brut = net_toplam - maliyet
    kar = {'maliyet': _r(maliyet), 'brut': _r(brut), 'marj': round(brut / net_toplam * 100, 1) if net_toplam else 0}
    odeme = odeme_secenekleri(net_toplam, kdv, kurallar['odeme'], maliyet, girdi.get('odeme_secimi'))

    s = {'paket_tipi': tip, 'satirlar': satirlar, 'lisans': lisans, 'hizmet': hizmet, 'uyarlama': uyarlama,
         'sabit': sabit, 'seyahat': seyahat, 'ara_toplam': _r(ara),
         'kampanya': {'ad': girdi.get('kampanya') or 'YOK', 'yuzde': kampanya_yuzde, 'tutar': _r(kampanya_tutar)},
         'net_toplam': genel['net'], 'kdv': genel['kdv'], 'genel_toplam': genel['dahil'], 'kdv_orani': kdv,
         'odeme': odeme, 'kar': kar, 'rakip': rakip_karsilastir(saat, ayar, rakipler or [], ucret),
         'uyarilar': uyarilar, '_uyarilar': list(uyarilar),
         '_baglam': {'kullanici': masaustu + mobil, 'masaustu': masaustu, 'mobil': mobil,
                     'cevaplar': girdi.get('cevaplar') or {}, 'tarih': _tarih(girdi.get('tarih'))}}
    s['bakim'] = bakim_hesapla(girdi, s, paketler, ayar, kurallar['bakim'], kurallar['sorular'])
    s['ybh'] = ybh_hesapla(tip, [{'kod': x['kod'], 'kullanici': x['adet']} for x in satirlar], paketler, ayar)
    s['dia'] = dia_ozeti(satirlar, paketler, tip, s['ybh'])
    secili_bakim = next((b for b in s['bakim']['paketler'] if b['ad'] == s['bakim']['secili']), None)
    s['tco'] = tco_hesapla(net_toplam, s['ybh']['yillik']['net'], secili_bakim['yillik'] if secili_bakim else 0,
                           ayar, kdv)
    s['kontrol'] = kontroller(girdi, s, ayar, paketler, kurallar)
    s['asistan'] = satis_asistani(girdi, s, ayar, kurallar)
    # eski arayuz uyumu: 'uyarilar' kritik + uyari mesajlari
    s['uyarilar'] = [k['mesaj'] for k in s['kontrol']['liste'] if k['seviye'] != 'BİLGİ']
    return s


def _tarih(t):
    try:
        return datetime.strptime(t, '%Y-%m-%d').date() if t else date.today()
    except (TypeError, ValueError):
        return date.today()


def rakip_karsilastir(saat, ayar, rakipler, ucret=None):
    """Uyarlama bedeli karsilastirmasi (saat x saatlik ucret)."""
    ucret = float(ayar['saatlik_ucret']) if ucret is None else ucret
    biz = saat * ucret
    satirlar = []
    ayni_saat = bool(float(ayar.get('rakip_ayni_saat', 1) or 0))
    for r in rakipler:
        r_saat = saat if ayni_saat or not r.get('saat') else float(r['saat'])
        toplam = r_saat * float(r.get('ucret') or 0)
        if toplam:
            satirlar.append({'firma': r['firma'], 'saat': r_saat, 'ucret': r['ucret'], 'toplam': _r(toplam),
                             'fark': _r(toplam - biz), 'fark_yuzde': round((toplam / biz - 1) * 100, 1) if biz else 0})
    ort = sum(s['toplam'] for s in satirlar) / len(satirlar) if satirlar else 0
    if not ort or not biz:
        oneri, konum = 'Rakip verisi girin' if not ort else 'Uyarlama saati yok', None
    else:
        konum = round((biz / ort - 1) * 100, 1)
        if biz < ort * (1 - float(ayar['rakip_ucuz_esik']) / 100):
            oneri = 'Rekabetçisiniz - alım ihtimali yüksek'
        elif biz > ort * (1 + float(ayar['rakip_pahali_esik']) / 100):
            oneri = 'Pahalısınız - indirim düşünün'
        else:
            oneri = 'Dengeli - normal aralık'
    return {'biz': _r(biz), 'satirlar': satirlar, 'ortalama': _r(ort), 'konum': konum, 'oneri': oneri}


# ======================================================================
# Yillik bakim (YBH)
# ======================================================================
def kademe_katsayi(deger, kademe_metni):
    """'5:1; 10:1.1; *:1.65' -> deger<=5 ise 1, <=10 ise 1.1, aksi halde 1.65"""
    son = 1.0
    for parca in str(kademe_metni).split(';'):
        if ':' not in parca:
            continue
        sinir, kats = parca.split(':', 1)
        sinir, kats = sinir.strip(), float(kats.strip().replace(',', '.'))
        if sinir == '*':
            son = kats
            continue
        if deger <= float(sinir):
            return kats
    return son


def ybh_hesapla(tip, kalemler, paketler, ayar):
    """kalemler: [{kod, kullanici}]. Lisans bedeli = birim + (kullanici-1) * birim * kull% ; baz = lisans * YBH%."""
    kdv = float(ayar['kdv'])
    satirlar, uyarilar = [], []
    for k in kalemler:
        p = paketler.get(k.get('kod'))
        if not p:
            uyarilar.append('Paket listesinde yok: ' + str(k.get('kod')))
            continue
        kull = max(int(k.get('kullanici') or 1), 1)
        birim = birim_fiyat(p, tip)
        lisans = birim + (kull - 1) * birim * float(p.get('kull_yuzde') or 0) / 100
        ybh = float(p.get('ybh_yuzde') or 0)
        satirlar.append({'kod': p['kod'], 'ad': p.get('ad', ''), 'kullanici': kull, 'modul': modul_sayisi(p),
                         'lisans': _r(lisans), 'ybh_yuzde': ybh, 'baz': _r(lisans * ybh / 100)})
    baz = sum(s['baz'] for s in satirlar)
    top_kull = sum(s['kullanici'] for s in satirlar)
    top_modul = sum(s['modul'] for s in satirlar)
    k_modul = kademe_katsayi(top_modul, ayar['ybh_modul_kademe'])
    k_kull = kademe_katsayi(top_kull, ayar['ybh_kullanici_kademe'])
    return {'satirlar': satirlar, 'baz': _r(baz), 'kullanici': top_kull, 'modul': top_modul,
            'modul_katsayi': k_modul, 'kullanici_katsayi': k_kull, 'katsayi': round(k_modul * k_kull, 4),
            'yillik': _kdvli(baz * k_modul * k_kull, kdv), 'uyarilar': uyarilar}
