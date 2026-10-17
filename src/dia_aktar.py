"""Temizlenmis *_dia_hazir.xlsx dosyalarini DIA'ya web servisle aktarir.

- Mevcut kartlar kod alanina gore listelenir; kod DIA'da yoksa EKLE, varsa GUNCELLE (veya ATLA).
- Hangi Excel kolonunun hangi DIA alanina gidecegi dia_eslesme.json dosyasindadir;
  musteri klasorunde yoksa varsayilan hali olusturulur ve elle duzenlenebilir.
- kuru_calistir=True iken hicbir sey gonderilmez, sadece gonderilecek istekler raporlanir.
"""
import copy
import json
import os

import pandas as pd

from dia_ws import DiaHatasi, DiaKontorHatasi
from il_ilce import anahtar

ESLESME_DOSYASI = 'dia_eslesme.json'

# Deger yazimi:
#   "dia_alani": "excel_kolonu"                      -> kart["dia_alani"] = satir["excel_kolonu"]
#   "_key_xxx": {"arama_alani": "excel_kolonu"}      -> kart["_key_xxx"] = {"arama_alani": deger}  (DIA key'i bulur)
# '?' ile baslayan anahtarlar devre disidir (dogrulanmamis alanlar); basindaki '?' silinince aktif olur.
VARSAYILAN_ESLESME = {
    '_aciklama': ("Excel kolonu -> DIA alani eslemesi. '?' ile baslayan satirlar dogrulanmamis "
                  "alanlardir ve gonderilmez; DIA dokumanindan (ScfCarikartModel, ScfStokkartModel) "
                  "kontrol edip basindaki '?' isaretini silin."),
    'cari': {
        'modul': 'scf',
        'servis_listele': 'scf_carikart_listele',
        'servis_ekle': 'scf_carikart_ekle',
        'servis_guncelle': 'scf_carikart_guncelle',
        'kod_alani': 'carikartkodu',
        'excel_kod_kolonu': 'carikartkodu',
        'mevcut_ise': 'guncelle',
        'kart': {
            'carikartkodu': 'carikartkodu',
            'unvan': 'unvan',
            'verginumarasi': 'verginumarasi',
            'tckimlikno': 'tckimlikno2',
            '?carikayitturu': 'cari kayitturu',
            '?eposta': 'eposta',
            '?webadresi': 'web adresi',
            '?ozelkod1': 'ozelkod1',
            '?_key_sis_vergidairesi': {'?vergidairesiadi': 'vergidairesi'},
        },
        'sabit': {'carikarttipi': 'AL'},
        'adres': {
            'aktif': True,
            'servis_ekle': 'scf_carikart_adresleri_ekle',
            'bag_alani': '_key_scf_carikart',
            'kart': {
                'adres1': 'adres1',
                'adres2': 'adres2',
                'ilce': 'ilce',
                'postakodu': 'postakodu',
                'telefon1': 'telefon1',
                'telefon2': 'telefon2',
                'ceptel': 'ceptel',
                'fax': 'fax',
                '?eposta': 'eposta',
            },
            'sabit': {'?adresadi': 'MERKEZ'},
            'sehir_kolonu': 'sehir',
            'sehir_key_alani': '_key_sis_sehirler',
            'sehir_servis': 'sis_sehirler_listele',
            'sehir_ad_alani': 'sehiradi',
        },
    },
    'stok': {
        'modul': 'scf',
        'servis_listele': 'scf_stokkart_listele',
        'servis_ekle': 'scf_stokkart_ekle',
        'servis_guncelle': 'scf_stokkart_guncelle',
        'kod_alani': 'stokkartkodu',
        'excel_kod_kolonu': 'stok kart kodu',
        'mevcut_ise': 'guncelle',
        'kart': {
            'stokkartkodu': 'stok kart kodu',
            'aciklama': 'stok adı',
            '?ozelkod1': 'Özel Kod 1',
            '?marka': 'Marka',
        },
        'sabit': {'stokkartturu': 'TCR'},
        'birim': {
            'aktif': False,
            'alt_liste': 'm_birimler',
            'kolon': 'Ana Birim',
            'kart': {'?_key_scf_birim': {'?birimkod': '{deger}'}},
            'barkod_kolonu': 'Barkod No',
            'barkod_alt_liste': 'm_barkodlar',
            'barkod_alani': '?barkod',
        },
    },
    'liste_parametreleri': {'?selectedcolumns': ['_key', '{kod_alani}']},
}


def eslesme_yukle(klasor):
    """Musteri klasorundeki dia_eslesme.json'u okur; yoksa varsayilani yazar."""
    yol = os.path.join(klasor, ESLESME_DOSYASI)
    if not os.path.exists(yol):
        with open(yol, 'w', encoding='utf-8') as f:
            json.dump(VARSAYILAN_ESLESME, f, ensure_ascii=False, indent=2)
        return copy.deepcopy(VARSAYILAN_ESLESME)
    with open(yol, encoding='utf-8') as f:
        return json.load(f)


def _aktif(sozluk):
    return {k: v for k, v in sozluk.items() if not str(k).startswith('?') and not str(k).startswith('_aciklama')}


def _bos(deger):
    return deger is None or (not isinstance(deger, str) and pd.isna(deger)) or str(deger).strip() in ('', 'nan')


def _deger(deger):
    if isinstance(deger, float) and deger.is_integer():
        return int(deger)
    if hasattr(deger, 'item'):        # numpy tipleri
        return deger.item()
    return str(deger).strip() if isinstance(deger, str) else deger


def kart_olustur(satir, alanlar, sabit=None):
    """Excel satirindan DIA kart sozlugu uretir. Bos hucreler gonderilmez (guncellemede silmesin)."""
    kart = {}
    for dia_alani, kaynak in _aktif(alanlar).items():
        if isinstance(kaynak, dict):
            arama = {}
            for arama_alani, kolon in _aktif(kaynak).items():
                if kolon in satir and not _bos(satir[kolon]):
                    arama[arama_alani] = _deger(satir[kolon])
            if arama:
                kart[dia_alani] = arama
        elif kaynak in satir and not _bos(satir[kaynak]):
            kart[dia_alani] = _deger(satir[kaynak])
    for alan, deger in _aktif(sabit or {}).items():
        kart.setdefault(alan, deger)
    return kart


def _liste_parametreleri(eslesme, kod_alani):
    params = {}
    for k, v in _aktif(eslesme.get('liste_parametreleri', {})).items():
        if isinstance(v, list):
            v = [kod_alani if x == '{kod_alani}' else x for x in v]
        params[k] = v
    return params


def mevcut_kartlar(istemci, ayar, eslesme):
    """DIA'daki kartlarin {kod: _key} sozlugu."""
    kod_alani = ayar['kod_alani']
    satirlar = istemci.listele(ayar['modul'], ayar['servis_listele'],
                               params=_liste_parametreleri(eslesme, kod_alani))
    return {str(s.get(kod_alani)).strip(): s.get('_key') for s in satirlar if s.get(kod_alani) is not None}


def sehir_keyleri(istemci, adres_ayar):
    """{il anahtari: _key} - il adlari il_ilce.anahtar() ile normalize edilir."""
    satirlar = istemci.listele('sis', adres_ayar['sehir_servis'])
    ad_alani = adres_ayar['sehir_ad_alani']
    return {anahtar(s.get(ad_alani, '')): s.get('_key') for s in satirlar if s.get(ad_alani)}


def _birim_ekle(kart, satir, birim_ayar):
    kolon = birim_ayar.get('kolon')
    if not kolon or kolon not in satir or _bos(satir[kolon]):
        return
    birim = {}
    for alan, kaynak in _aktif(birim_ayar.get('kart', {})).items():
        if isinstance(kaynak, dict):
            birim[alan] = {a: _deger(satir[kolon]) if v == '{deger}' else v for a, v in _aktif(kaynak).items()}
        else:
            birim[alan] = _deger(satir[kolon]) if kaynak == '{deger}' else kaynak
    barkod_alani = birim_ayar.get('barkod_alani', '')
    b_kolon = birim_ayar.get('barkod_kolonu')
    if barkod_alani and not barkod_alani.startswith('?') and b_kolon in satir and not _bos(satir[b_kolon]):
        birim[birim_ayar['barkod_alt_liste']] = [{barkod_alani: _deger(satir[b_kolon])}]
    if birim:
        kart[birim_ayar['alt_liste']] = [birim]


def aktar(istemci, df, modul_adi, eslesme, kuru_calistir=False, log=print):
    """df satirlarini DIA'ya aktarir. Doner: rapor DataFrame."""
    ayar = eslesme[modul_adi]
    kod_kolonu = ayar['excel_kod_kolonu']
    if kod_kolonu not in df.columns:
        raise ValueError("'" + kod_kolonu + "' kolonu yok. Once temizleme calistirilmali.")

    log('DIA mevcut ' + modul_adi + ' kartlari okunuyor...')
    mevcut = mevcut_kartlar(istemci, ayar, eslesme)
    log('  DIA\'da ' + str(len(mevcut)) + ' kart var.')

    adres_ayar = ayar.get('adres') if ayar.get('adres', {}).get('aktif') else None
    sehirler = {}
    if adres_ayar and adres_ayar.get('sehir_key_alani') and not adres_ayar['sehir_key_alani'].startswith('?'):
        try:
            sehirler = sehir_keyleri(istemci, adres_ayar)
            log('  ' + str(len(sehirler)) + ' sehir okundu.')
        except DiaHatasi as e:
            log('  UYARI: sehir listesi okunamadi, adreste il gonderilmeyecek: ' + str(e))
    birim_ayar = ayar.get('birim') if ayar.get('birim', {}).get('aktif') else None

    rapor = []
    toplam = len(df)
    for i, (_, satir) in enumerate(df.iterrows(), start=1):
        kod = '' if _bos(satir[kod_kolonu]) else str(_deger(satir[kod_kolonu])).strip()
        kayit = {'satir': i + 1, 'kod': kod, 'islem': '', 'sonuc': '', '_key': '', 'mesaj': '', 'istek': ''}
        if not kod:
            kayit.update(islem='ATLA', sonuc='HATA', mesaj='Kod bos')
            rapor.append(kayit)
            continue

        kart = kart_olustur(satir, ayar['kart'], ayar.get('sabit'))
        if birim_ayar:
            _birim_ekle(kart, satir, birim_ayar)
        key = mevcut.get(kod)
        if key is not None:
            if ayar.get('mevcut_ise', 'guncelle') != 'guncelle':
                kayit.update(islem='ATLA', sonuc='OK', _key=key, mesaj='DIA\'da var')
                rapor.append(kayit)
                continue
            kart = {k: v for k, v in kart.items() if k not in _aktif(ayar.get('sabit', {}))}
            kart['_key'] = key
            islem, servis = 'GUNCELLE', ayar['servis_guncelle']
        else:
            islem, servis = 'EKLE', ayar['servis_ekle']
        kayit['islem'] = islem
        kayit['istek'] = json.dumps(kart, ensure_ascii=False, default=str)

        adres_kart = None
        if adres_ayar and islem == 'EKLE':
            adres_kart = kart_olustur(satir, adres_ayar['kart'], adres_ayar.get('sabit'))
            il = satir.get(adres_ayar.get('sehir_kolonu', ''), '')
            if sehirler and not _bos(il):
                sehir_key = sehirler.get(anahtar(il))
                if sehir_key:
                    adres_kart[adres_ayar['sehir_key_alani']] = sehir_key
                else:
                    kayit['mesaj'] = 'Il DIA sehir listesinde yok: ' + str(il)
            if len(adres_kart) <= len(_aktif(adres_ayar.get('sabit', {}))):
                adres_kart = None

        if kuru_calistir:
            kayit['sonuc'] = 'KURU'
            if adres_kart:
                kayit['istek'] += ' | ADRES: ' + json.dumps(adres_kart, ensure_ascii=False, default=str)
            rapor.append(kayit)
            continue

        try:
            if islem == 'EKLE':
                yeni = istemci.ekle(ayar['modul'], servis, kart)
                kayit['_key'] = yeni if yeni is not None else ''
                mevcut[kod] = yeni
                if adres_kart and yeni is not None:
                    adres_kart[adres_ayar['bag_alani']] = yeni
                    try:
                        istemci.ekle(ayar['modul'], adres_ayar['servis_ekle'], adres_kart)
                    except DiaKontorHatasi:
                        raise
                    except DiaHatasi as e:
                        kayit['mesaj'] = (kayit['mesaj'] + ' | ' if kayit['mesaj'] else '') + 'Adres: ' + e.msg
            else:
                istemci.guncelle(ayar['modul'], servis, kart)
                kayit['_key'] = key
            kayit['sonuc'] = 'OK'
        except DiaKontorHatasi as e:
            kayit.update(sonuc='HATA', mesaj='KONTOR YETERSIZ: ' + e.msg)
            rapor.append(kayit)
            log('  DURDURULDU: kontor yetersiz.')
            break
        except DiaHatasi as e:
            kayit.update(sonuc='HATA', mesaj=e.code + ' - ' + e.msg)
        rapor.append(kayit)
        if i % 50 == 0 or i == toplam:
            log('  ' + str(i) + '/' + str(toplam) + ' islendi')

    rdf = pd.DataFrame(rapor, columns=['satir', 'kod', 'islem', 'sonuc', '_key', 'mesaj', 'istek'])
    ozet = rdf.groupby(['islem', 'sonuc']).size().to_dict() if len(rdf) else {}
    log('  Ozet: ' + ', '.join(k[0] + '/' + k[1] + '=' + str(v) for k, v in ozet.items()))
    return rdf


def eksik_dosya_mesaji(klasor, modul):
    """*_dia_hazir.xlsx yoksa sebebini ve ne yapilacagini anlatir."""
    hazir = modul['cikti_dosya']
    inceleme = hazir.replace('.xlsx', '_INCELEME.xlsx')
    girdi = modul.get('girdi_dosya', '')
    satirlar = [hazir + ' bulunamadi.', 'Klasor: ' + klasor]
    if os.path.exists(os.path.join(klasor, inceleme)):
        satirlar.append("Sebep: CALISTIR 'Inceleme' modunda calismis; bu mod sadece " + inceleme +
                        " uretir. Ana ekranda 'Otomatik Sil' (veya E isaretledikten sonra 'Uygula') "
                        "secip CALISTIR'a basin.")
    elif girdi and os.path.exists(os.path.join(klasor, girdi)):
        satirlar.append('Sebep: ' + girdi + " var ama henuz temizlenmemis. Ana ekranda modulu ve "
                        "'Otomatik Sil' secip CALISTIR'a basin.")
    else:
        satirlar.append('Sebep: bu klasorde ' + (girdi or 'girdi dosyasi') + " de yok. Once 'Logo'dan Cek' "
                        "ile cekin veya dosyayi bu klasore koyun; ana ekrandaki Musteri Klasoru dogru mu?")
    return '\n'.join(satirlar)


def dosyadan_aktar(istemci, klasor, modul_adi, kurallar, kuru_calistir=False, log=print):
    """<cikti_dosya> (ornek cari_dia_hazir.xlsx) dosyasini okuyup aktarir, raporu klasore yazar."""
    from dia_isle import METIN_KOLONLAR
    modul = kurallar[modul_adi]
    hazir = os.path.join(klasor, modul['cikti_dosya'])
    if not os.path.exists(hazir):
        raise FileNotFoundError(eksik_dosya_mesaji(klasor, modul))
    log('Okunuyor: ' + modul['cikti_dosya'])
    df = pd.read_excel(hazir, dtype={k: str for k in METIN_KOLONLAR})
    if 'SIL_ONAY' in df.columns:
        df = df.drop(columns=['SIL_ONAY'])
    eslesme = eslesme_yukle(klasor)
    rapor = aktar(istemci, df, modul_adi, eslesme, kuru_calistir, log)
    rapor_adi = modul_adi + ('_DIA_KURU_RAPOR.xlsx' if kuru_calistir else '_DIA_AKTARIM_RAPORU.xlsx')
    rapor.to_excel(os.path.join(klasor, rapor_adi), index=False)
    log('Rapor: ' + rapor_adi)
    return rapor
