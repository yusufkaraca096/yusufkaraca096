"""Il / ilce ayristirma ve duzeltme.

Musterilerin sik yaptigi hatalar:
  - il ve ilce ayni hucreye yazilmis      : 'Kadıköy/İstanbul', 'İstanbul - Kadıköy'
  - il kolonunda ilce, ilce kolonunda il  : sehir='Kadıköy', ilce='İstanbul'
  - il bos, sadece ilce yazilmis          : sehir='', ilce='Kadıköy'
  - yazim farklari                         : 'istanbul', 'ISTANBUL', 'Kadikoy', 'Afyon', 'Urfa'
"""
import re
import pandas as pd

from il_ilce_veri import IL_ILCE

_ASCII = str.maketrans('ÇĞİÖŞÜÂÎÛ', 'CGIOSUAIU')

# Eski / kisa / halk arasinda kullanilan adlar
IL_TAKMA_ADLAR = {
    'AFYON': 'AFYONKARAHİSAR',
    'MARAS': 'KAHRAMANMARAŞ',
    'KMARAS': 'KAHRAMANMARAŞ',
    'URFA': 'ŞANLIURFA',
    'SURFA': 'ŞANLIURFA',
    'ANTEP': 'GAZİANTEP',
    'GANTEP': 'GAZİANTEP',
    'ICEL': 'MERSİN',
}
ILCE_TAKMA_ADLAR = {
    'EYUP': 'EYÜPSULTAN',
    'ONDOKUZMAYIS': '19 MAYIS',
}


def tr_buyuk(s):
    """Turkce kurallara gore buyuk harf: i -> İ, ı -> I."""
    return str(s).replace('i', 'İ').replace('ı', 'I').upper()


def tr_baslik(s):
    """'KADIKÖY' -> 'Kadıköy', '19 MAYIS' -> '19 Mayıs'."""
    kelimeler = []
    for k in str(s).split(' '):
        if not k:
            continue
        kucuk = k[1:].replace('I', 'ı').replace('İ', 'i').lower()
        kelimeler.append(tr_buyuk(k[0]) + kucuk)
    return ' '.join(kelimeler)


def anahtar(s):
    """Karsilastirma anahtari: buyuk harf, Turkce karakter ve bosluk/nokta yok."""
    s = tr_buyuk(str(s).strip()).translate(_ASCII)
    return re.sub(r'[^A-Z0-9]', '', s)


# Arama tablolari
_IL = {}          # anahtar -> resmi il adi
_ILCE = {}        # anahtar -> {il, ...}
_ILCE_AD = {}     # anahtar -> resmi ilce adi
for _il, _ilceler in IL_ILCE.items():
    _IL[anahtar(_il)] = _il
    for _ilce in _ilceler:
        _k = anahtar(_ilce)
        _ILCE.setdefault(_k, set()).add(_il)
        _ILCE_AD[_k] = _ilce
for _k, _il in IL_TAKMA_ADLAR.items():
    _IL.setdefault(_k, _il)
for _k, _ilce in ILCE_TAKMA_ADLAR.items():
    _ka = anahtar(_ilce)
    _ILCE[_k] = _ILCE[_ka]
    _ILCE_AD[_k] = _ILCE_AD[_ka]

_AYRAC = re.compile(r'\s*[/\\,;|\-–()]+\s*')


def _bos_mu(x):
    return x is None or (not isinstance(x, str) and pd.isna(x)) or str(x).strip() in ('', 'nan', 'None')


def _parcala(metin):
    """Hucreyi il/ilce adaylarina boler. Her parca: (orijinal, il_adi|None, ilce_anahtari|None)."""
    if _bos_mu(metin):
        return []
    metin = str(metin).strip()
    k = anahtar(metin)
    if k in _IL or k in _ILCE:
        parcalar = [metin]
    else:
        parcalar = [p for p in _AYRAC.split(metin) if p.strip()]
        if len(parcalar) == 1:
            kelimeler = metin.split()
            parcalar = []
            i = 0
            while i < len(kelimeler):
                # iki kelimelik adlar ('19 Mayıs') once denenir
                if i + 1 < len(kelimeler):
                    ikili = kelimeler[i] + ' ' + kelimeler[i + 1]
                    if anahtar(ikili) in _IL or anahtar(ikili) in _ILCE:
                        parcalar.append(ikili)
                        i += 2
                        continue
                parcalar.append(kelimeler[i])
                i += 1
    sonuc = []
    for p in parcalar:
        pk = anahtar(p)
        sonuc.append((p.strip(), _IL.get(pk), pk if pk in _ILCE else None))
    return sonuc


def il_ilce_coz(il_deger, ilce_deger, il_bos_ise_doldur=True):
    """(il, ilce) hucrelerini cozer.

    Doner: (il, ilce, durum). durum 'OK', 'BOŞ', 'DÜZELTİLDİ (...)' veya 'UYARI: ...'.
    Il/ilce resmi adlariyla BUYUK harf doner; taninmayan metin oldugu gibi kalir.
    """
    if _bos_mu(il_deger) and _bos_mu(ilce_deger):
        return '', '', 'BOŞ'

    il_p = _parcala(il_deger)
    ilce_p = _parcala(ilce_deger)

    # Bulunan iller: once il kolonundakiler, sonra ilce kolonundakiler
    iller = []
    for _, il, _ in il_p + ilce_p:
        if il and il not in iller:
            iller.append(il)

    uyarilar = []
    il = iller[0] if iller else None
    if len(iller) > 1:
        uyarilar.append('BIRDEN FAZLA IL (' + ', '.join(iller) + ')')

    # Ilce adaylari: once ilce kolonundakiler. Il adi olarak kullanilan parca ilce sayilmaz.
    ilce_adaylari = [(o, ik) for o, i, ik in ilce_p + il_p if ik and not i]
    ilce = None
    for o, ik in ilce_adaylari:
        if il is None or il in _ILCE[ik]:
            ilce = ik
            break

    if il is None and ilce is not None:
        olasi = sorted(_ILCE[ilce])
        if len(olasi) == 1 and il_bos_ise_doldur:
            il = olasi[0]
        elif len(olasi) > 1:
            uyarilar.append(_ILCE_AD[ilce] + ' BIRDEN FAZLA ILDE VAR (' + ', '.join(olasi) + ')')

    if ilce is None and ilce_adaylari:
        ik = ilce_adaylari[0][1]
        uyarilar.append(_ILCE_AD[ik] + ' ILCESI ' + str(il) + ' ILINDE YOK')
        ilce = ik

    # Taninmayan parcalar
    taninan = set()
    if il:
        taninan.add(anahtar(il))
    if ilce:
        taninan.add(ilce)
    taninmayan = [o for o, i, ik in il_p + ilce_p
                  if anahtar(o) not in taninan and not (i and i == il)]

    yeni_il = il if il else ('' if _bos_mu(il_deger) else str(il_deger).strip())
    if ilce:
        yeni_ilce = _ILCE_AD[ilce]
    elif not _bos_mu(ilce_deger) and anahtar(ilce_deger) not in _IL:
        yeni_ilce = str(ilce_deger).strip()
    else:
        yeni_ilce = ''
    if il is None and yeni_il and anahtar(yeni_il) in _ILCE:
        # il kolonunda sadece (belirsiz) ilce vardi; ilceye tasindi
        yeni_il = ''

    if il is None and not _bos_mu(il_deger) and not uyarilar:
        uyarilar.append('IL TANINMADI (' + str(il_deger).strip() + ')')
    elif taninmayan:
        if ilce is None:
            uyarilar.append('ILCE TANINMADI (' + ', '.join(taninmayan) + ')')
        else:
            uyarilar.append('ATILAN METIN (' + ', '.join(taninmayan) + ')')

    eski_il = '' if _bos_mu(il_deger) else str(il_deger).strip()
    eski_ilce = '' if _bos_mu(ilce_deger) else str(ilce_deger).strip()
    if uyarilar:
        durum = 'UYARI: ' + ' | '.join(uyarilar)
    elif anahtar(eski_il) != anahtar(yeni_il) or anahtar(eski_ilce) != anahtar(yeni_ilce):
        durum = ('DÜZELTİLDİ (' + eski_il + ' / ' + eski_ilce + ' -> '
                 + yeni_il + ' / ' + yeni_ilce + ')')
    else:
        durum = 'OK'
    return yeni_il, yeni_ilce, durum


def il_ilce_duzelt(df, il_sutun='sehir', ilce_sutun='ilce', il_bos_ise_doldur=True,
                   yazim='BUYUK'):
    """DataFrame'deki il/ilce kolonlarini duzeltir.

    Doner: (df, rapor). rapor: [{'satir', 'alan', 'deger', 'durum'}, ...]
    yazim: 'BUYUK' -> İSTANBUL / KADIKÖY,  'BASLIK' -> İstanbul / Kadıköy
    """
    rapor = []
    if il_sutun not in df.columns and ilce_sutun not in df.columns:
        return df, rapor
    if il_sutun not in df.columns:
        df[il_sutun] = ''
    if ilce_sutun not in df.columns:
        df[ilce_sutun] = ''

    yeni_iller = []
    yeni_ilceler = []
    for i, (il_d, ilce_d) in enumerate(zip(df[il_sutun].tolist(), df[ilce_sutun].tolist())):
        il, ilce, durum = il_ilce_coz(il_d, ilce_d, il_bos_ise_doldur)
        if yazim == 'BASLIK':
            if anahtar(il) in _IL:
                il = tr_baslik(il)
            if anahtar(ilce) in _ILCE:
                ilce = tr_baslik(ilce)
        yeni_iller.append(il)
        yeni_ilceler.append(ilce)
        if durum not in ('OK', 'BOŞ'):
            rapor.append({'satir': i + 2, 'alan': il_sutun + '/' + ilce_sutun,
                          'deger': str(il) + ' / ' + str(ilce), 'durum': durum})

    df[il_sutun] = yeni_iller
    df[ilce_sutun] = yeni_ilceler
    duzeltilen = sum(1 for r in rapor if r['durum'].startswith('DÜZELTİLDİ'))
    uyari = len(rapor) - duzeltilen
    if rapor:
        print('  Il/ilce: ' + str(duzeltilen) + ' satir duzeltildi, ' + str(uyari) + ' uyari')
    return df, rapor
