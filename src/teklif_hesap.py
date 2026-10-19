"""Teklif hesap motoru (RACA_DIA_Teklif_Motoru.xlsx > HESAP / YBH BAKIM sayfalarinin karsiligi).

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
Tum tutarlar KDV haric; KDV ayrica hesaplanir.
"""


def _r(x):
    return round(float(x or 0), 2)


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


def satir_hesapla(paket, tip, adet, ind_yuzde, kdv, saat=None, ind_tutar=0):
    birim = birim_fiyat(paket, tip)
    kull = float(paket.get('kull_yuzde') or 0)
    toplam = 0.0 if adet <= 0 else birim + (adet - 1) * birim * kull / 100
    # Indirim = yuzde + sabit tutar; ikisi de satir NET'inden duser, KDV ve tum toplamlar net uzerinden
    indirim = min(toplam * float(ind_yuzde or 0) / 100 + float(ind_tutar or 0), toplam)
    net = toplam - indirim
    tablo_saat = float(paket.get('ent_saat' if tip == 'ENTERPRISE' else 'std_saat') or 0)
    saat = tablo_saat if saat in (None, '') else float(saat)
    return {'kod': paket['kod'], 'ad': paket.get('ad', ''), 'tur': paket.get('tur', ''), 'adet': adet,
            'birim': _r(birim), 'kull_yuzde': kull, 'toplam': _r(toplam), 'ind_yuzde': float(ind_yuzde or 0),
            'ind_tutar': float(ind_tutar or 0),
            'indirim': _r(indirim), 'net': _r(net), 'kdv': _r(net * kdv / 100), 'dahil': _r(net * (1 + kdv / 100)),
            'saat': saat, 'tablo_saat': tablo_saat}


def _kdvli(net, kdv):
    return {'net': _r(net), 'kdv': _r(net * kdv / 100), 'dahil': _r(net * (1 + kdv / 100))}


def hesapla(girdi, paketler, ayar, sabit_fiyatlar=None, rakipler=None):
    """girdi: {musteri, masaustu, mobil, paket_tipi: OTOMATIK|STANDART|ENTERPRISE,
               kalemler: [{kod, adet (None = otomatik), ind_yuzde, ind_tutar, saat (None = ADAM SAAT tablosu)}],
               sabitler: [ad, ...],
               kampanya, kampanya_yuzde, maliyet}"""
    sabit_fiyatlar = sabit_fiyatlar or {}
    kdv = float(ayar['kdv'])
    masaustu = int(girdi.get('masaustu') or 0)
    mobil = int(girdi.get('mobil') or 0)
    tip = paket_tipi_bul(masaustu, mobil, ayar, girdi.get('paket_tipi', 'OTOMATIK'))
    uyarilar = []

    satirlar = []
    for sira, k in enumerate(girdi.get('kalemler', [])):
        p = paketler.get(k.get('kod'))
        if not p:
            uyarilar.append('Paket listesinde yok: ' + str(k.get('kod')))
            continue
        if not p.get('aktif', 1):
            uyarilar.append(p['kod'] + ' güncel DIA fiyat listesinde yok (eski fiyat kullanıldı)')
        if tip == 'ENTERPRISE' and not p.get('ent_fiyat'):
            uyarilar.append(p['kod'] + ': Enterprise fiyatı yok, Standart fiyat kullanıldı')
        adet = k.get('adet')
        adet = kullanici_sayisi(p, masaustu, mobil) if adet in (None, '') else int(adet)
        satir = satir_hesapla(p, tip, adet, k.get('ind_yuzde', 0), kdv, k.get('saat'), k.get('ind_tutar', 0))
        satir['sira'] = sira     # formdaki kalem sirasi (bilinmeyen kod atlaninca kaymasin)
        satirlar.append(satir)

    lisans_toplam = sum(s['toplam'] for s in satirlar)
    lisans_indirim = sum(s['indirim'] for s in satirlar)
    lisans = dict(_kdvli(sum(s['net'] for s in satirlar), kdv), toplam=_r(lisans_toplam), indirim=_r(lisans_indirim))

    saat = sum(s['saat'] for s in satirlar)
    lisans_payi = lisans['net'] * float(ayar['lisans_yuzdesi']) / 100
    saat_bedeli = saat * float(ayar['saatlik_ucret'])
    ham = lisans_payi + saat_bedeli
    uyarlama_net, sinir = ham, ''
    if satirlar and ham < float(ayar['min_uyarlama']):
        uyarlama_net, sinir = float(ayar['min_uyarlama']), 'ALT SINIR'
    if float(ayar['max_uyarlama'] or 0) > 0 and uyarlama_net > float(ayar['max_uyarlama']):
        uyarlama_net, sinir = float(ayar['max_uyarlama']), 'ÜST SINIR'
    uyarlama = dict(_kdvli(uyarlama_net, kdv), saat=saat, lisans_payi=_r(lisans_payi),
                    saat_bedeli=_r(saat_bedeli), ham=_r(ham), sinir=sinir)

    sabit_kalemler = []
    for ad in girdi.get('sabitler', []):
        if ad in sabit_fiyatlar:
            sabit_kalemler.append({'ad': ad, 'net': _r(sabit_fiyatlar[ad])})
        else:
            uyarilar.append('Sabit kalem tanımlı değil: ' + ad)
    sabit = dict(_kdvli(sum(s['net'] for s in sabit_kalemler), kdv), kalemler=sabit_kalemler)

    ara = lisans['net'] + uyarlama['net'] + sabit['net']
    kampanya_yuzde = float(girdi.get('kampanya_yuzde') or 0)
    kampanya_tutar = ara * kampanya_yuzde / 100
    net_toplam = ara - kampanya_tutar
    genel = _kdvli(net_toplam, kdv)

    odeme = []
    pesin = net_toplam * (1 - float(ayar['pesin_indirim']) / 100)
    odeme.append(dict(_kdvli(pesin, kdv), plan='Peşin (%%%g indirim)' % float(ayar['pesin_indirim']), adet=1))
    for n in str(ayar.get('taksitler') or '').split(','):
        n = n.strip()
        if n.isdigit() and int(n) > 1:
            t = dict(_kdvli(net_toplam / int(n), kdv), plan=n + ' Taksit', adet=int(n))
            odeme.append(t)

    maliyet = float(girdi.get('maliyet') or 0)
    brut = net_toplam - maliyet
    kar = {'maliyet': _r(maliyet), 'brut': _r(brut), 'marj': round(brut / net_toplam * 100, 1) if net_toplam else 0}

    return {'paket_tipi': tip, 'satirlar': satirlar, 'lisans': lisans, 'uyarlama': uyarlama, 'sabit': sabit,
            'ara_toplam': _r(ara), 'kampanya': {'ad': girdi.get('kampanya') or 'YOK', 'yuzde': kampanya_yuzde,
                                                 'tutar': _r(kampanya_tutar)},
            'net_toplam': genel['net'], 'kdv': genel['kdv'], 'genel_toplam': genel['dahil'], 'kdv_orani': kdv,
            'odeme': odeme, 'kar': kar, 'rakip': rakip_karsilastir(saat, ayar, rakipler or []),
            'uyarilar': uyarilar}


def rakip_karsilastir(saat, ayar, rakipler):
    """Uyarlama bedeli karsilastirmasi (saat x saatlik ucret)."""
    biz = saat * float(ayar['saatlik_ucret'])
    satirlar = []
    ayni_saat = bool(float(ayar.get('rakip_ayni_saat', 1) or 0))
    for r in rakipler:
        r_saat = saat if ayni_saat or not r.get('saat') else float(r['saat'])
        toplam = r_saat * float(r.get('ucret') or 0)
        if toplam:
            satirlar.append({'firma': r['firma'], 'saat': r_saat, 'ucret': r['ucret'], 'toplam': _r(toplam),
                             'fark': _r(toplam - biz), 'fark_yuzde': round((toplam / biz - 1) * 100, 1) if biz else 0})
    ort = sum(s['toplam'] for s in satirlar) / len(satirlar) if satirlar else 0
    if not ort:
        oneri, konum = 'Rakip verisi girin', None
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


def modul_sayisi(paket):
    urunler = (paket.get('urunler') or '').strip()
    return len([u for u in urunler.split('-') if u.strip()]) if urunler else 0


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
