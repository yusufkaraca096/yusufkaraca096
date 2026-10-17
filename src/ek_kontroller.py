"""Ek temizlik ve kontroller (ayarlari ek_ayarlar.py'de).

Akista uc noktada cagrilir (dia_isle.veri_temizle):
  hucre_on_temizlik : her seyden once (hata degerleri, gorunmez karakter, bozuk Turkce)
  on_islem          : il/ilce duzeltmesinden ve hiyerarsi kontrolunden once (deger duzeltmeleri)
  son_kontroller    : en sonda (sadece uyarilar)
Ayrica dia_isle.modul_isle: capraz_kontrol, degisiklik raporu, esleme hafizasi.

Rapor satiri: {'satir': konum + 2, 'alan', 'deger', 'durum'} (kontrol_uyari_ekle ile KONTROL_UYARI'ya yazilir).
"""
import difflib
import functools
import json
import os
import re
from datetime import datetime
from decimal import Decimal, InvalidOperation

import pandas as pd

from _ortak import ayar, sayisal_metin, kod_isle, KAYNAK_SATIR
from il_ilce import anahtar, tr_buyuk, tr_baslik, _IL, _ILCE, _ILCE_AD

_KATLA = str.maketrans('ÇĞİÖŞÜÂÎÛ', 'CGIOSUAIU')
_HARF = '0-9A-ZÇĞİÖŞÜÂÎÛ'


# ============================================================
# Yardimcilar
# ============================================================
def _bos(x):
    return x is None or (not isinstance(x, str) and pd.isna(x)) or str(x).strip() in ('', 'nan', 'None')


def _metin(x):
    return '' if _bos(x) else sayisal_metin(x).strip()


def katla(s):
    """Karsilastirma icin: Turkce buyuk harf, Turkce karakterler ASCII. Uzunluk korunur."""
    return tr_buyuk(s).translate(_KATLA)


def _ekle(rapor, i, alan, deger, durum):
    rapor.append({'satir': i + 2, 'alan': alan, 'deger': str(deger), 'durum': durum})


def liste_ayari(deger):
    """'a, b' -> ['a', 'b'];  'YOK' -> []"""
    if _bos(deger) or str(deger).strip().upper() in ('YOK', '-', 'NONE'):
        return []
    return [p.strip() for p in str(deger).split(',') if p.strip()]


def cift_ayari(deger):
    """'A=B; C=D' -> [('A', 'B'), ('C', 'D')]"""
    if _bos(deger) or str(deger).strip().upper() in ('YOK', '-', 'NONE'):
        return []
    sonuc = []
    for parca in re.split(r'[;\n]', str(deger)):
        if '=' in parca:
            a, b = parca.split('=', 1)
            if a.strip():
                sonuc.append((a.strip(), b.strip()))
    return sonuc


def _sayi(x):
    if _bos(x):
        return None
    if isinstance(x, (int, float)):
        return float(x)
    try:
        return float(str(x).replace(' ', '').replace(',', '.'))
    except ValueError:
        return None


def _tarih(x):
    if _bos(x):
        return None
    if isinstance(x, datetime):
        return x
    t = pd.to_datetime(str(x), dayfirst=True, errors='coerce')
    return None if pd.isna(t) else t.to_pydatetime()


@functools.lru_cache(maxsize=512)
def _ifade_deseni(aranan):
    """'LTD STI' -> L.T.D ve S.T.I harflerini nokta/bosluk esnek arayan desen. Sonda $ -> sadece metin sonu."""
    sonda = aranan.endswith('$')
    tokenlar = [t for t in re.split(r'[\s.]+', katla(aranan.rstrip('$')).strip()) if t]
    if not tokenlar:
        return None
    govde = r'[\s.]+'.join(r'\.?'.join(re.escape(c) for c in t) for t in tokenlar)
    son = r'\.?\s*$' if sonda else r'\.?(?![' + _HARF + '])'
    return re.compile(r'(?<![' + _HARF + '])' + govde + son)


def kisaltma_uygula(metin, ciftler):
    """Ifade ciftlerini (aranan=yazilacak) Turkce harf/nokta/buyuk-kucuk duyarsiz uygular."""
    if _bos(metin):
        return metin
    s = str(metin)
    for aranan, yazilacak in sorted(ciftler, key=lambda c: -len(c[0])):
        desen = _ifade_deseni(aranan)
        if desen is None:
            continue
        k = katla(s)
        if len(k) != len(s):
            s = tr_buyuk(s)
            k = katla(s)
        for m in reversed(list(desen.finditer(k))):
            bitis = m.end()
            sonra = s[bitis:bitis + 1]
            ek = ' ' if sonra and sonra not in ' .,;)' else ''
            s = s[:m.start()] + yazilacak + ek + s[bitis:]
    return re.sub(r'\s+', ' ', s).strip()


# ============================================================
# 1) Hucre on temizligi (tum moduller)
# ============================================================
_GORUNMEZ = [('_x000D_', ' '), ('_x000A_', ' '), ('_x0009_', ' '), ('\r', ' '), ('\n', ' '), ('\t', ' '),
             ('\xa0', ' '), ('​', ''), ('‌', ''), ('‍', ''), ('﻿', ''), ('­', '')]
_MOJIBAKE = {'Ã‡': 'Ç', 'Ã§': 'ç', 'Ä°': 'İ', 'Ä±': 'ı', 'Ã–': 'Ö', 'Ã¶': 'ö', 'Ãœ': 'Ü', 'Ã¼': 'ü',
             'Åž': 'Ş', 'ÅŸ': 'ş', 'Äž': 'Ğ', 'ÄŸ': 'ğ', 'Ã‚': 'Â', 'Ã¢': 'â'}
_LATIN5 = str.maketrans({'Ý': 'İ', 'ý': 'ı', 'Þ': 'Ş', 'þ': 'ş', 'Ð': 'Ğ', 'ð': 'ğ'})


def bozuk_turkce_onar(s):
    if not isinstance(s, str):
        return s
    if any(c in s for c in 'ÃÄÅ'):
        try:
            return s.encode('cp1252').decode('utf-8')
        except (UnicodeEncodeError, UnicodeDecodeError):
            for a, b in _MOJIBAKE.items():
                s = s.replace(a, b)
    if any(c in s for c in 'ÝýÞþÐð'):
        s = s.translate(_LATIN5)
    return s


def hucre_on_temizlik(df, modul, modul_adi):
    hata_aktif = ayar(modul, 'hata_degerleri_temizle', modul_adi)
    hatalar = {katla(h) for h in liste_ayari(ayar(modul, 'hata_degerleri', modul_adi))} if hata_aktif else set()
    gorunmez = ayar(modul, 'gorunmez_karakter_temizle', modul_adi)
    bozuk = ayar(modul, 'bozuk_turkce_duzelt', modul_adi)
    sayac = {'hata': 0, 'gorunmez': 0, 'bozuk': 0}

    def temizle(x):
        if not isinstance(x, str):
            return x
        s = x
        if gorunmez:
            for a, b in _GORUNMEZ:
                s = s.replace(a, b)
            if s != x:
                sayac['gorunmez'] += 1
        if bozuk:
            y = bozuk_turkce_onar(s)
            if y != s:
                sayac['bozuk'] += 1
                s = y
        if hatalar and katla(s.strip()) in hatalar:
            sayac['hata'] += 1
            return None
        return s

    for col in df.columns:
        if col != KAYNAK_SATIR and (df[col].dtype == object or pd.api.types.is_string_dtype(df[col])):
            df[col] = df[col].map(temizle)
    if sayac['hata']:
        print('  Hata/anlamsiz deger bosaltildi: ' + str(sayac['hata']) + ' hucre')
    if sayac['gorunmez']:
        print('  Gorunmez karakter temizlendi: ' + str(sayac['gorunmez']) + ' hucre')
    if sayac['bozuk']:
        print('  Bozuk Turkce karakter duzeltildi: ' + str(sayac['bozuk']) + ' hucre')
    return df


_BILIMSEL = re.compile(r'^\s*(\d+)(?:[.,](\d+))?\s*[eE]\s*\+?\s*(\d+)\s*$')


def bilimsel_isle(deger):
    """'1,23457E+11' -> ('123457000000', 'GEÇERSİZ ...') ; '8.69123456789E+12' -> ('8691234567890', 'DÜZELTİLDİ')"""
    if not isinstance(deger, str):
        return deger, 'OK'
    m = _BILIMSEL.match(deger)
    if not m:
        return deger, 'OK'
    tam, kesir, us = m.group(1), m.group(2) or '', int(m.group(3))
    try:
        sayi = Decimal(tam + '.' + kesir if kesir else tam) * (Decimal(10) ** us)
    except InvalidOperation:
        return deger, 'OK'
    yeni = str(int(sayi))
    anlamli = len((tam + kesir).lstrip('0'))
    if anlamli < len(yeni) and yeni.endswith('0' * (len(yeni) - anlamli)):
        return deger, 'GEÇERSİZ: BİLİMSEL GÖSTERİM, HANE KAYBI VAR (' + deger.strip() + \
            "). Kaynak Excel'de hucreyi Metin yapip tekrar alin"
    return yeni, 'DÜZELTİLDİ (' + deger.strip() + ' -> ' + yeni + ')'


def bilimsel_kontrol(df, kolonlar, rapor):
    for alan in kolonlar:
        if alan not in df.columns:
            continue
        yeni = []
        for i, v in enumerate(df[alan].tolist()):
            d, durum = bilimsel_isle(v)
            yeni.append(d)
            if durum != 'OK':
                _ekle(rapor, i, alan, d, durum)
        df[alan] = yeni
    return df


# ============================================================
# 2) Cari
# ============================================================
_SIRKET = re.compile(r'(?<![A-Z0-9])(LTD|LIMITED|STI|SIRKETI|A\.?\s?S|ANONIM|KOOP|KOOPERATIFI?|HOLDING|TIC|SAN)(?![A-Z0-9])')
_VD_EK = re.compile(r'(?:[\s.,\-]*(?<![A-Z0-9])(?:VERGI\s*DAIRESI(?:\s*MUD(?:URLUGU)?)?|V\s*\.?\s*D\.?|MUD(?:URLUGU)?\.?|MD\.?))+[\s.,\-]*$')


def sirket_mi(unvan):
    return bool(_SIRKET.search(katla(_metin(unvan)).replace('.', ' ').replace('  ', ' '))) or \
        bool(re.search(r'(?<![A-Z0-9])A\.S\.?(?![A-Z0-9])', katla(_metin(unvan))))


def vergi_dairesi_isle(deger, ozel):
    if _bos(deger):
        return '', 'BOŞ'
    orijinal = str(deger).strip()
    s = orijinal
    k = katla(s)
    if len(k) == len(s):
        m = _VD_EK.search(k)
        if m and m.start() > 0:
            s = s[:m.start()]
    s = re.sub(r'\s+', ' ', s).strip(' .,-')
    a = anahtar(s)
    ozel_harita = {anahtar(x): y for x, y in ozel}
    if a in ozel_harita:
        s = ozel_harita[a]
    elif a in _IL:
        s = _IL[a]
    elif a in _ILCE_AD:
        s = _ILCE_AD[a]
    s = tr_buyuk(s)
    if s != orijinal:
        return s, 'DÜZELTİLDİ (' + orijinal + ' -> ' + s + ')'
    return s, 'OK'


_ADRES_YOL = re.compile(r'(MAH|MH|CAD|CD|SOK|SK|SOKAK|SOKAGI|CADDE|CADDESI|MAHALLE|MAHALLESI|BUL|BLV|BULVAR|BULVARI|'
                        r'YOLU|YOL|MEYDAN|MEYDANI|APT|APARTMANI|SITESI|SIT|PLAZA|IS|HANI|PASAJI|KOYU|BELDESI)')


def adresten_il_ilce(adres):
    """Adreste gecen (il, ilce). Cadde/sokak/mahalle adi olarak gecenler atlanir. Bulamazsa ('', '')."""
    kelimeler = re.findall(r'[0-9A-Za-zÇĞİÖŞÜçğıöşüÂâÎîÛû]+', _metin(adres))
    adaylar = []   # (konum, uzunluk, anahtar)
    for i in range(len(kelimeler)):
        for n in (2, 1):
            if i + n > len(kelimeler):
                continue
            sonraki = katla(kelimeler[i + n]) if i + n < len(kelimeler) else ''
            if sonraki and _ADRES_YOL.fullmatch(sonraki):
                continue
            adaylar.append((i, n, anahtar(' '.join(kelimeler[i:i + n]))))
    il = ''
    il_konum = -1
    for i, n, a in adaylar:
        if a in _IL and i >= il_konum:
            il, il_konum = _IL[a], i
    if not il:
        return '', ''
    ilce = ''
    for i, n, a in adaylar:
        if a in _ILCE and il in _ILCE[a] and anahtar(_ILCE_AD[a]) != anahtar(il) and i != il_konum:
            ilce = _ILCE_AD[a]
    return il, ilce


def _telefon_cekirdek(x):
    r = re.sub(r'\D', '', _metin(x))
    if r.startswith('90') and len(r) == 12:
        r = r[2:]
    if r.startswith('0') and len(r) == 11:
        r = r[1:]
    return r


def cari_on_islem(df, modul, rapor):
    a = lambda ad: ayar(modul, ad, 'cari')  # noqa: E731

    if a('unvan_standart') and 'unvan' in df.columns:
        ciftler = cift_ayari(a('unvan_kisaltmalari'))
        yeni = []
        for i, v in enumerate(df['unvan'].tolist()):
            y = kisaltma_uygula(v, ciftler) if ciftler else v
            if not _bos(v) and y != str(v).strip():
                _ekle(rapor, i, 'unvan', y, 'DÜZELTİLDİ (' + str(v).strip() + ' -> ' + y + ')')
            yeni.append(y)
        df['unvan'] = yeni

    if a('vergi_dairesi_duzelt') and 'vergidairesi' in df.columns:
        ozel = cift_ayari(a('vergi_dairesi_eslesme'))
        yeni = []
        for i, v in enumerate(df['vergidairesi'].tolist()):
            y, durum = vergi_dairesi_isle(v, ozel)
            yeni.append(y)
            if durum.startswith('DÜZELTİLDİ'):
                _ekle(rapor, i, 'vergidairesi', y, durum)
        df['vergidairesi'] = yeni

    adres_kolonlari = [c for c in ('adres1', 'adres2') if c in df.columns]
    if a('adres_kisaltma'):
        ciftler = cift_ayari(a('adres_kisaltmalari'))
        for c in adres_kolonlari:
            df[c] = [kisaltma_uygula(v, ciftler) if not _bos(v) else v for v in df[c].tolist()]

    il_s, ilce_s = ayar(modul, 'il_sutunu', 'cari'), ayar(modul, 'ilce_sutunu', 'cari')
    if a('adresten_il_ilce') and adres_kolonlari:
        iller = df[il_s].tolist() if il_s in df.columns else [''] * len(df)
        ilceler = df[ilce_s].tolist() if ilce_s in df.columns else [''] * len(df)
        adresler = [df[c].tolist() for c in reversed(adres_kolonlari)]
        doldurulan = 0
        for i in range(len(df)):
            il_v, ilce_v = iller[i], ilceler[i]
            if not (_bos(il_v) or _bos(ilce_v)):
                continue
            il = ilce = ''
            for liste in adresler:
                il, ilce = adresten_il_ilce(liste[i])
                if il:
                    break
            if not il:
                continue
            if _bos(il_v) and _bos(ilce_v):
                iller[i], ilceler[i] = il, ilce
            elif _bos(il_v):
                iller[i] = il      # ilce yazili: uyusmazligi il_ilce_duzelt yakalar
            elif ilce and anahtar(il_v) == anahtar(il):
                ilceler[i] = ilce
            else:
                continue
            doldurulan += 1
            _ekle(rapor, i, il_s + '/' + ilce_s, str(iller[i]) + ' / ' + str(ilceler[i]),
                  'DÜZELTİLDİ (ADRESTEN DOLDURULDU, kontrol edin)')
        df[il_s], df[ilce_s] = iller, ilceler
        if doldurulan:
            print('  Adresten il/ilce dolduruldu: ' + str(doldurulan) + ' satir')

    cep = a('cep_sutunu')
    if a('cep_ayir') and cep:
        kaynaklar = [c for c in modul.get('telefon_alanlari', []) if c in df.columns and c != cep
                     and 'fax' not in c.lower() and 'faks' not in c.lower()]
        if kaynaklar:
            cepler = df[cep].tolist() if cep in df.columns else [''] * len(df)
            telefonlar = {c: df[c].tolist() for c in kaynaklar}
            for i in range(len(df)):
                if not _bos(cepler[i]):
                    continue
                for c in kaynaklar:
                    v = telefonlar[c][i]
                    r = _telefon_cekirdek(v)
                    if len(r) == 10 and r.startswith('5'):
                        cepler[i] = v
                        telefonlar[c][i] = ''
                        _ekle(rapor, i, cep, v, 'DÜZELTİLDİ (' + c + ' cep numarasi -> ' + cep + ')')
                        break
            df[cep] = cepler
            for c in kaynaklar:
                df[c] = telefonlar[c]
    return df


IL_PLAKA = {anahtar(il): no for no, il in enumerate([
    'ADANA', 'ADIYAMAN', 'AFYONKARAHİSAR', 'AĞRI', 'AMASYA', 'ANKARA', 'ANTALYA', 'ARTVİN', 'AYDIN',
    'BALIKESİR', 'BİLECİK', 'BİNGÖL', 'BİTLİS', 'BOLU', 'BURDUR', 'BURSA', 'ÇANAKKALE', 'ÇANKIRI', 'ÇORUM',
    'DENİZLİ', 'DİYARBAKIR', 'EDİRNE', 'ELAZIĞ', 'ERZİNCAN', 'ERZURUM', 'ESKİŞEHİR', 'GAZİANTEP', 'GİRESUN',
    'GÜMÜŞHANE', 'HAKKARİ', 'HATAY', 'ISPARTA', 'MERSİN', 'İSTANBUL', 'İZMİR', 'KARS', 'KASTAMONU', 'KAYSERİ',
    'KIRKLARELİ', 'KIRŞEHİR', 'KOCAELİ', 'KONYA', 'KÜTAHYA', 'MALATYA', 'MANİSA', 'KAHRAMANMARAŞ', 'MARDİN',
    'MUĞLA', 'MUŞ', 'NEVŞEHİR', 'NİĞDE', 'ORDU', 'RİZE', 'SAKARYA', 'SAMSUN', 'SİİRT', 'SİNOP', 'SİVAS',
    'TEKİRDAĞ', 'TOKAT', 'TRABZON', 'TUNCELİ', 'ŞANLIURFA', 'UŞAK', 'VAN', 'YOZGAT', 'ZONGULDAK', 'AKSARAY',
    'BAYBURT', 'KARAMAN', 'KIRIKKALE', 'BATMAN', 'ŞIRNAK', 'BARTIN', 'ARDAHAN', 'IĞDIR', 'YALOVA', 'KARABÜK',
    'KİLİS', 'OSMANİYE', 'DÜZCE'], start=1)}


def cari_son_kontrol(df, modul, rapor):
    a = lambda ad: ayar(modul, ad, 'cari')  # noqa: E731
    if a('vkn_tc_tutarlilik') and 'unvan' in df.columns:
        for i in range(len(df)):
            vkn = re.sub(r'\D', '', _metin(df.at[i, 'verginumarasi'])) if 'verginumarasi' in df.columns else ''
            tc = re.sub(r'\D', '', _metin(df.at[i, 'tckimlikno2'])) if 'tckimlikno2' in df.columns else ''
            if not vkn and len(tc) == 11 and sirket_mi(df.at[i, 'unvan']):
                _ekle(rapor, i, 'verginumarasi', tc, 'UYARI: ŞİRKET ÜNVANI AMA VKN YOK, SADECE TC VAR')

    if a('ad_soyad_ayir') and 'unvan' in df.columns:
        ad_s, soyad_s = a('ad_sutunu'), a('soyad_sutunu')
        for c in (ad_s, soyad_s):
            if c not in df.columns:
                df[c] = ''
        tur = df['cari kayitturu'].tolist() if 'cari kayitturu' in df.columns else ['SHS'] * len(df)
        for i in range(len(df)):
            u = _metin(df.at[i, 'unvan'])
            kelimeler = u.split()
            if tur[i] != 'SHS' or sirket_mi(u) or not 2 <= len(kelimeler) <= 4 or \
                    not all(re.fullmatch(r'[A-Za-zÇĞİÖŞÜçğıöşüÂâÎîÛû.]+', k) for k in kelimeler):
                continue
            if _bos(df.at[i, ad_s]) and _bos(df.at[i, soyad_s]):
                df.at[i, ad_s] = ' '.join(kelimeler[:-1])
                df.at[i, soyad_s] = kelimeler[-1]

    il_s = ayar(modul, 'il_sutunu', 'cari')
    posta_kodu_kontrol(df, il_s, 'postakodu', a('posta_kodu_il_kontrol'), rapor)
    return df


def posta_kodu_kontrol(df, il_s, pk_s, aktif, rapor):
    if not aktif or il_s not in df.columns or pk_s not in df.columns:
        return
    for i in range(len(df)):
        pk = re.sub(r'\D', '', _metin(df.at[i, pk_s]))
        plaka = IL_PLAKA.get(anahtar(_metin(df.at[i, il_s])))
        if len(pk) == 5 and plaka and int(pk[:2]) != plaka:
            _ekle(rapor, i, pk_s, pk, 'UYARI: POSTA KODU ' + pk + ' ' + _metin(df.at[i, il_s]) +
                  ' ILINE AIT DEGIL (plaka ' + str(plaka).zfill(2) + ')')


# ============================================================
# 3) Stok
# ============================================================
def gtin_gecerli(kod):
    if not kod.isdigit() or len(kod) not in (8, 12, 13, 14):
        return None     # standart disi: kontrol edilmez
    toplam = sum(int(d) * (3 if i % 2 == 0 else 1) for i, d in enumerate(reversed(kod[:-1])))
    return (10 - toplam % 10) % 10 == int(kod[-1])


_DOVIZ_ESLESME = [('TL', r'TL|TRY|₺|YTL'), ('USD', r'USD|\$|DOLAR'), ('EUR', r'EUR|EURO|€|AVRO'),
                  ('GBP', r'GBP|£|STERLIN')]


def fiyat_isle(deger):
    """'1.250,50 TL' -> (1250.5, 'TL', durum).  Sayi (float) ise dokunulmaz."""
    if _bos(deger) or not isinstance(deger, str):
        return deger, '', 'OK'
    s = katla(deger.strip())
    doviz = ''
    for kod, desen in _DOVIZ_ESLESME:
        if re.search(r'(?<![A-Z])(' + desen + r')(?![A-Z])', s):
            doviz = kod
            s = re.sub(r'(?<![A-Z])(' + desen + r')(?![A-Z])', '', s)
            break
    s = s.replace(' ', '')
    if not re.fullmatch(r'-?[\d.,]+', s or 'x'):
        return deger, doviz, 'GEÇERSİZ FİYAT (' + deger + ')'
    if '.' in s and ',' in s:
        ondalik = '.' if s.rfind('.') > s.rfind(',') else ','
        binlik = ',' if ondalik == '.' else '.'
        s = s.replace(binlik, '').replace(ondalik, '.')
    elif ',' in s:
        s = s.replace('.', '').replace(',', '.') if s.count(',') == 1 else s.replace(',', '')
    elif s.count('.') > 1 or re.fullmatch(r'-?\d{1,3}\.\d{3}', s):
        s = s.replace('.', '')      # 1.250 -> 1250 (Turkce binlik)
    try:
        f = float(s)
    except ValueError:
        return deger, doviz, 'GEÇERSİZ FİYAT (' + deger + ')'
    return f, doviz, 'DÜZELTİLDİ (' + deger + ' -> ' + str(f) + (' ' + doviz if doviz else '') + ')'


_OLCU = {'GR': 'GR', 'G': 'GR', 'GRAM': 'GR', 'KG': 'KG', 'ML': 'ML', 'CL': 'CL', 'LT': 'LT', 'L': 'LT',
         'LITRE': 'LT', 'CM': 'CM', 'MM': 'MM', 'MT': 'MT', 'M': 'M', 'MG': 'MG', 'CC': 'CC'}
_OLCU_DESEN = re.compile(r'(?<![0-9A-Za-zÇĞİÖŞÜçğıöşü])(\d+(?:[.,]\d+)?)\s*(gram|litre|gr|kg|ml|cl|lt|mg|cc|cm|mm|mt|g|l|m)'
                         r'(?![0-9A-Za-zÇĞİÖŞÜçğıöşü])', re.IGNORECASE)


def stok_adi_olcu_duzelt(ad):
    if _bos(ad):
        return ad
    return _OLCU_DESEN.sub(lambda m: m.group(1) + ' ' + _OLCU[katla(m.group(2))], str(ad))


def stok_on_islem(df, modul, rapor):
    a = lambda ad: ayar(modul, ad, 'stok')  # noqa: E731

    if a('birim_standart'):
        harita = {anahtar(x): y for x, y in cift_ayari(a('birim_eslesme'))}
        for c in liste_ayari(a('birim_sutunu')):
            if c not in df.columns:
                continue
            yeni = []
            for i, v in enumerate(df[c].tolist()):
                if _bos(v):
                    yeni.append(v)
                    continue
                y = harita.get(anahtar(v))
                if y is None:
                    _ekle(rapor, i, c, v, 'UYARI: TANIMSIZ BİRİM (' + str(v) + '), birim_eslesme ayarina ekleyin')
                    yeni.append(v)
                else:
                    if y != str(v).strip():
                        _ekle(rapor, i, c, y, 'DÜZELTİLDİ (' + str(v) + ' -> ' + y + ')')
                    yeni.append(y)
            df[c] = yeni

    if a('fiyat_temizle'):
        for n in range(1, 11):
            c = 'Fiyat ' + str(n)
            if c not in df.columns:
                continue
            dc = 'Döviz ' + str(n)
            yeni = []
            for i, v in enumerate(df[c].tolist()):
                f, doviz, durum = fiyat_isle(v)
                yeni.append(f)
                if durum != 'OK':
                    _ekle(rapor, i, c, f, durum)
                if doviz and doviz != 'TL':
                    if dc not in df.columns:
                        df[dc] = ''
                    if _bos(df.at[i, dc]):
                        df.at[i, dc] = doviz
                    elif str(df.at[i, dc]).strip().upper() != doviz:
                        _ekle(rapor, i, dc, df.at[i, dc], 'UYARI: FİYATTA ' + doviz + ' YAZIYOR, ' + dc + ' FARKLI')
            df[c] = yeni

    if 'stok adı' in df.columns:
        yazim = str(a('stok_adi_yazim')).upper()
        yeni = []
        for i, v in enumerate(df['stok adı'].tolist()):
            y = v
            if a('stok_adi_birim_duzelt'):
                y = stok_adi_olcu_duzelt(y)
            if not _bos(y) and yazim == 'BUYUK':
                y = tr_buyuk(y)
            elif not _bos(y) and yazim == 'BASLIK':
                y = tr_baslik(y)
            if not _bos(v) and y != v:
                _ekle(rapor, i, 'stok adı', y, 'DÜZELTİLDİ (' + str(v) + ' -> ' + str(y) + ')')
            yeni.append(y)
        df['stok adı'] = yeni
    return df


def stok_son_kontrol(df, modul, rapor):
    a = lambda ad: ayar(modul, ad, 'stok')  # noqa: E731
    bk = a('barkod_sutunu')
    if a('barkod_kontrol') and bk in df.columns:
        for i, v in enumerate(df[bk].tolist()):
            for kod in re.split(r'[,;/\s]+', _metin(v)):
                if kod and gtin_gecerli(kod) is False:
                    _ekle(rapor, i, bk, kod, 'GEÇERSİZ BARKOD (kontrol hanesi hatali: ' + kod + ')')

    eski = set()
    for o in liste_ayari(a('kdv_eski_oranlar')):
        try:
            eski.add(int(float(o)))
        except ValueError:
            pass
    if eski:
        for c in modul.get('kdv_alanlari', []):
            if c not in df.columns:
                continue
            for i, v in enumerate(df[c].tolist()):
                n = _sayi(v)
                if n is not None and int(n) in eski:
                    _ekle(rapor, i, c, v, 'UYARI: ESKİ KDV ORANI (%' + str(int(n)) + ')')

    s_c, a_c = a('satis_fiyat_sutunu'), a('alis_fiyat_sutunu')
    if a('alis_satis_kontrol') and s_c in df.columns and a_c in df.columns:
        for i in range(len(df)):
            s, al = _sayi(df.at[i, s_c]), _sayi(df.at[i, a_c])
            if s and al and al > s:
                _ekle(rapor, i, a_c, al, 'UYARI: ALIŞ (' + str(al) + ') > SATIŞ (' + str(s) + ')')
    return df


# ============================================================
# 4) Banka / Personel
# ============================================================
BANKALAR = {
    '00010': ('T.C. ZİRAAT BANKASI', ['ZIRAATBANK', 'TCZIRAAT', 'ZIRAAT']),
    '00012': ('HALKBANK', ['HALKBANK', 'TURKIYEHALK', 'HALK']),
    '00015': ('VAKIFBANK', ['VAKIFBANK', 'VAKIFLARBANK', 'VAKIF']),
    '00032': ('TEB', ['TURKEKONOMI', 'TEB']),
    '00046': ('AKBANK', ['AKBANK']),
    '00059': ('ŞEKERBANK', ['SEKERBANK']),
    '00062': ('GARANTİ BBVA', ['GARANTI']),
    '00064': ('İŞ BANKASI', ['TURKIYEISBANK', 'ISBANK', 'ISBANKASI']),
    '00067': ('YAPI KREDİ', ['YAPIKREDI', 'YAPIVEKREDI', 'YKB']),
    '00099': ('ING BANK', ['INGBANK', 'ING']),
    '00103': ('FİBABANKA', ['FIBABANK', 'FIBA']),
    '00111': ('QNB', ['QNB', 'FINANSBANK']),
    '00123': ('HSBC', ['HSBC']),
    '00124': ('ALTERNATİFBANK', ['ALTERNATIF']),
    '00125': ('BURGAN BANK', ['BURGAN']),
    '00134': ('DENİZBANK', ['DENIZBANK']),
    '00135': ('ANADOLUBANK', ['ANADOLUBANK']),
    '00146': ('ODEABANK', ['ODEA']),
    '00203': ('ALBARAKA TÜRK', ['ALBARAKA']),
    '00205': ('KUVEYT TÜRK', ['KUVEYT']),
    '00206': ('TÜRKİYE FİNANS', ['TURKIYEFINANS']),
    '00209': ('ZİRAAT KATILIM', ['ZIRAATKATILIM']),
    '00210': ('VAKIF KATILIM', ['VAKIFKATILIM']),
    '00211': ('EMLAK KATILIM', ['EMLAKKATILIM']),
}
_ALIAS = sorted(((al, kod) for kod, (_, als) in BANKALAR.items() for al in als), key=lambda x: -len(x[0]))
_KISA_ALIAS = {'ING', 'TEB', 'YKB', 'QNB', 'HSBC', 'HALK', 'VAKIF', 'FIBA', 'ODEA'}


def banka_kodu_adindan(ad):
    kelimeler = [anahtar(k) for k in re.split(r'[\s.,\-]+', _metin(ad)) if anahtar(k)]
    # alias bir kelimenin basindan baslamali: 'HOLDING BANKASI' icinde 'INGBANK' aranmaz
    kuyruklar = [''.join(kelimeler[k:]) for k in range(len(kelimeler))]
    for al, kod in _ALIAS:
        if al in _KISA_ALIAS:
            if al in kelimeler:
                return kod
        elif any(t.startswith(al) for t in kuyruklar):
            return kod
    return None


def banka_kodu_ibandan(iban):
    s = re.sub(r'\s', '', _metin(iban)).upper()
    return s[4:9] if s.startswith('TR') and len(s) == 26 else None


def banka_on_islem(df, modul, modul_adi, rapor):
    a = lambda ad: ayar(modul, ad, modul_adi)  # noqa: E731
    iban_c, ad_c = a('iban_sutunu'), a('banka_adi_sutunu')
    var_iban = iban_c in df.columns
    if (a('iban_banka_adi') and var_iban) or (a('banka_adi_standart') and ad_c in df.columns):
        if ad_c not in df.columns:
            df[ad_c] = ''
        for i in range(len(df)):
            mevcut = df.at[i, ad_c]
            iban_kod = banka_kodu_ibandan(df.at[i, iban_c]) if var_iban and a('iban_banka_adi') else None
            ad_kod = banka_kodu_adindan(mevcut) if not _bos(mevcut) else None
            if _bos(mevcut):
                if iban_kod in BANKALAR:
                    df.at[i, ad_c] = BANKALAR[iban_kod][0]
                    _ekle(rapor, i, ad_c, df.at[i, ad_c], "DÜZELTİLDİ (IBAN'dan banka adi dolduruldu)")
                continue
            if iban_kod in BANKALAR and ad_kod and ad_kod != iban_kod:
                _ekle(rapor, i, ad_c, mevcut, 'UYARI: BANKA ADI IBAN İLE UYUŞMUYOR (IBAN: ' +
                      BANKALAR[iban_kod][0] + ')')
            elif a('banka_adi_standart') and ad_kod and BANKALAR[ad_kod][0] != str(mevcut).strip():
                df.at[i, ad_c] = BANKALAR[ad_kod][0]
                _ekle(rapor, i, ad_c, df.at[i, ad_c], 'DÜZELTİLDİ (' + str(mevcut) + ' -> ' + df.at[i, ad_c] + ')')

    no_c = a('hesap_no_sutunu')
    if modul_adi == 'banka' and a('iban_hesap_no_doldur') and var_iban and no_c:
        if no_c not in df.columns:
            df[no_c] = ''
        for i in range(len(df)):
            s = re.sub(r'\s', '', _metin(df.at[i, iban_c])).upper()
            if _bos(df.at[i, no_c]) and s.startswith('TR') and len(s) == 26 and s[10:].isdigit():
                df.at[i, no_c] = s[10:].lstrip('0')
                _ekle(rapor, i, no_c, df.at[i, no_c], "DÜZELTİLDİ (IBAN'dan hesap no dolduruldu)")
    return df


_HESAP_DOVIZ = [('USD', r'USD|DOLAR|\$'), ('EUR', r'EUR|EURO|AVRO|€'), ('GBP', r'GBP|STERLIN|£'), ('TL', r'TL|TRY')]


def banka_son_kontrol(df, modul, rapor):
    if not ayar(modul, 'doviz_hesap_adi_kontrol', 'banka') or 'Hesap Adı' not in df.columns \
            or 'Döviz Türü' not in df.columns:
        return df
    for i in range(len(df)):
        ad = katla(_metin(df.at[i, 'Hesap Adı']))
        doviz = _metin(df.at[i, 'Döviz Türü']).upper() or 'TL'
        for kod, desen in _HESAP_DOVIZ:
            if re.search(r'(?<![A-Z])(' + desen + r')(?![A-Z])', ad):
                if kod != doviz:
                    _ekle(rapor, i, 'Döviz Türü', doviz, 'UYARI: HESAP ADINDA ' + kod + ' VAR, DÖVİZ TÜRÜ ' + doviz)
                break
    return df


_CINSIYET = {'E': 'ERKEK', 'ERKEK': 'ERKEK', 'BAY': 'ERKEK', 'M': 'ERKEK', 'MALE': 'ERKEK',
             'K': 'KADIN', 'KADIN': 'KADIN', 'BAYAN': 'KADIN', 'F': 'KADIN', 'FEMALE': 'KADIN', 'KIZ': 'KADIN'}
_MEDENI = {'EVLI': 'EVLİ', 'E': 'EVLİ', 'MARRIED': 'EVLİ', 'BEKAR': 'BEKAR', 'B': 'BEKAR', 'SINGLE': 'BEKAR',
           'DUL': 'DUL', 'BOSANMIS': 'BOŞANMIŞ'}


def kan_grubu_isle(v):
    s = katla(_metin(v)).replace('RH', ' ').replace('(', ' ').replace(')', ' ')
    m = re.match(r'^\s*(AB|A|B|0|O)\b\s*(.*)$', s)
    if not m:
        return v
    grup = '0' if m.group(1) in ('0', 'O') else m.group(1)
    kalan = m.group(2)
    if re.search(r'\+|POZ|POS', kalan):
        return grup + ' RH+'
    if re.search(r'-|NEG', kalan):
        return grup + ' RH-'
    return grup


def personel_on_islem(df, modul, rapor):
    if not ayar(modul, 'personel_standart', 'personel'):
        return df
    for adaylar, cevir in ((('Cinsiyet', 'Cinsiyeti'), lambda v: _CINSIYET.get(anahtar(v), v)),
                           (('Medeni Hal', 'Medeni Hali', 'Medeni Durum', 'Medeni Durumu'),
                            lambda v: _MEDENI.get(anahtar(v), v)),
                           (('Kan Grubu', 'Kan Grubu ', 'Kan'), kan_grubu_isle)):
        for c in adaylar:
            if c in df.columns:
                df[c] = [v if _bos(v) else cevir(v) for v in df[c].tolist()]
    return df


def personel_son_kontrol(df, modul, rapor):
    a = lambda ad: ayar(modul, ad, 'personel')  # noqa: E731
    bugun = datetime.now()
    yas_min, yas_max = _sayi(a('yas_min')) or 0, _sayi(a('yas_max')) or 200
    for i in range(len(df)):
        dogum = _tarih(df.at[i, 'Doğum Tarihi']) if 'Doğum Tarihi' in df.columns else None
        giris = _tarih(df.at[i, 'İşe Başlama Tarihi']) if 'İşe Başlama Tarihi' in df.columns else None
        if dogum and dogum > bugun:
            _ekle(rapor, i, 'Doğum Tarihi', dogum.strftime('%d.%m.%Y'), 'GEÇERSİZ: DOĞUM TARİHİ GELECEKTE')
            continue
        if dogum and (bugun - dogum).days / 365.25 > yas_max:
            _ekle(rapor, i, 'Doğum Tarihi', dogum.strftime('%d.%m.%Y'),
                  'UYARI: YAŞ ' + str(int((bugun - dogum).days / 365.25)) + ' (> ' + str(int(yas_max)) + ')')
        if dogum and giris and (giris - dogum).days / 365.25 < yas_min:
            _ekle(rapor, i, 'İşe Başlama Tarihi', giris.strftime('%d.%m.%Y'),
                  'UYARI: İŞE GİRİŞTE YAŞ ' + str(int((giris - dogum).days / 365.25)) + ' (< ' + str(int(yas_min)) + ')')
    il_s = ayar(modul, 'il_sutunu', 'personel')
    posta_kodu_kontrol(df, il_s, 'postakodu', a('posta_kodu_il_kontrol'), rapor)
    return df


# ============================================================
# 5) Demirbas
# ============================================================
def demirbas_son_kontrol(df, modul, rapor):
    if not ayar(modul, 'demirbas_kontrol', 'demirbas'):
        return df
    bugun = datetime.now()
    bir_c = ayar(modul, 'birikmis_amortisman_sutunu', 'demirbas')
    for i in range(len(df)):
        t = _tarih(df.at[i, 'TARİH']) if 'TARİH' in df.columns else None
        if t and t > bugun:
            _ekle(rapor, i, 'TARİH', t.strftime('%d.%m.%Y'), 'UYARI: ALIŞ TARİHİ GELECEKTE')
        tutar = _sayi(df.at[i, 'TUTAR']) if 'TUTAR' in df.columns else None
        if tutar is not None and tutar <= 0:
            _ekle(rapor, i, 'TUTAR', tutar, 'UYARI: TUTAR 0 VEYA NEGATİF')
        yil = _sayi(df.at[i, 'AMORTİSMAN YIL']) if 'AMORTİSMAN YIL' in df.columns else None
        oran = _sayi(df.at[i, 'AMORTİSMAN ORANI']) if 'AMORTİSMAN ORANI' in df.columns else None
        if yil and oran:
            normal, azalan = 100 / yil, min(200 / yil, 50)
            if abs(oran - normal) > 0.5 and abs(oran - azalan) > 0.5:
                _ekle(rapor, i, 'AMORTİSMAN ORANI', oran, 'UYARI: ORAN %' + str(oran) + ' SÜREYLE (' + str(int(yil)) +
                      ' yıl) UYUMSUZ: normal %' + str(round(normal, 2)) + ', azalan %' + str(round(azalan, 2)))
        if bir_c in df.columns and tutar is not None:
            bir = _sayi(df.at[i, bir_c])
            if bir is not None and bir > tutar:
                _ekle(rapor, i, bir_c, bir, 'GEÇERSİZ: BİRİKMİŞ AMORTİSMAN TUTARDAN BÜYÜK')
    return df


# ============================================================
# 6) Hesap plani
# ============================================================
TDHP = set('''
100 101 102 103 108 110 111 112 118 119 120 121 122 124 126 127 128 129 131 132 133 135 136 137 138 139
150 151 152 153 157 158 159 170 171 172 173 174 175 176 177 178 179 180 181 190 191 192 193 195 196 197 198 199
220 221 222 224 226 229 231 232 233 235 236 237 239 240 241 242 243 244 245 246 247 248 249
250 251 252 253 254 255 256 257 258 259 260 261 262 263 264 267 268 269 271 272 277 278 279 280 281
291 292 293 294 295 297 298 299
300 301 302 303 304 305 306 308 309 320 321 322 326 329 331 332 333 335 336 337 340 349 350 358
360 361 368 369 370 371 372 373 379 380 381 391 392 393 397 399
400 401 402 405 407 408 409 420 421 422 426 429 431 432 433 436 437 438 440 449 472 479 480 481 492 493 499
500 501 502 503 520 521 522 523 524 529 540 541 542 548 549 570 580 590 591
600 601 602 610 611 612 620 621 622 623 630 631 632 640 641 642 643 644 645 646 647 648 649
653 654 655 656 657 658 659 660 661 671 679 680 681 689 690 691 692 697 698
700 701 710 711 712 713 720 721 722 723 730 731 732 733 740 741 742 743 750 751 752 753
760 761 762 763 770 771 772 773 780 781 782 790 791 792 793 794 795 796 797 798 799
'''.split())


def hesap_kodu_kolonu(df):
    for aday in ('Hesap Kodu', 'hesap kodu', 'Kodu'):
        if aday in df.columns:
            return aday
    return None


def hesap_kodu_isle(kod, ayrac, seviyeler):
    s = _metin(kod)
    if not s:
        return s, 'BOŞ'
    yeni = re.sub(r'[\s\-_/,.]+', ayrac, s).strip(ayrac)
    if seviyeler and yeni.isdigit():
        toplam, parcalar, konum = 0, [], 0
        for n in seviyeler:
            if konum + n > len(yeni):
                break
            parcalar.append(yeni[konum:konum + n])
            konum += n
        if konum == len(yeni) and len(parcalar) > 1:
            yeni = ayrac.join(parcalar)
    if seviyeler:
        for n, (p, beklenen) in enumerate(zip(yeni.split(ayrac), seviyeler)):
            if len(p) != beklenen:
                return yeni, 'UYARI: ' + str(n + 1) + '. SEVİYE ' + str(len(p)) + ' HANE (' + str(beklenen) + \
                    ' OLMALI)' + (" - Excel sayi olarak okuyup sondaki 0'i silmis olabilir" if len(p) < beklenen else '')
    if yeni != s:
        return yeni, 'DÜZELTİLDİ (' + s + ' -> ' + yeni + ')'
    return yeni, 'OK'


def hesapplan_on_islem(df, modul, rapor):
    a = lambda ad: ayar(modul, ad, 'hesapplan')  # noqa: E731
    c = hesap_kodu_kolonu(df)
    if not c:
        return df
    ayrac = str(a('hesap_ayrac') or '.')[:1]
    seviyeler = [int(x) for x in liste_ayari(a('hesap_seviye_uzunluklari')) if str(x).isdigit()]
    yeni = []
    for i, v in enumerate(df[c].tolist()):
        y, durum = hesap_kodu_isle(v, ayrac, seviyeler)
        yeni.append(y)
        if durum not in ('OK', 'BOŞ'):
            _ekle(rapor, i, c, y, durum)
    df[c] = yeni

    if a('eksik_ust_hesap_ekle'):
        mevcut = set(k for k in yeni if k)
        eklenecek = []
        for k in yeni:
            parcalar = k.split(ayrac) if k else []
            for n in range(1, len(parcalar)):
                ust = ayrac.join(parcalar[:n])
                if ust not in mevcut:
                    mevcut.add(ust)
                    eklenecek.append(ust)
        if eklenecek:
            ad_c = 'Hesap Adı' if 'Hesap Adı' in df.columns else None
            satirlar = []
            for ust in sorted(eklenecek):
                satir = {c: ust}
                if ad_c:
                    satir[ad_c] = '(OTOMATIK) ' + ust
                satirlar.append(satir)
            basla = len(df)
            df = pd.concat([df, pd.DataFrame(satirlar)], ignore_index=True)
            for n, ust in enumerate(sorted(eklenecek)):
                _ekle(rapor, basla + n, c, ust, 'DÜZELTİLDİ (EKSİK ÜST HESAP EKLENDİ, adını girin)')
            print('  Eksik ust hesap eklendi: ' + str(len(eklenecek)))
    return df


def hesapplan_son_kontrol(df, modul, rapor):
    if not ayar(modul, 'tdhp_kontrol', 'hesapplan'):
        return df
    c = hesap_kodu_kolonu(df)
    if not c:
        return df
    gecerli = TDHP | set(liste_ayari(ayar(modul, 'ek_ana_hesaplar', 'hesapplan')))
    for i, v in enumerate(df[c].tolist()):
        ana = _metin(v)[:3]
        if not ana:
            continue
        if not ana.isdigit():
            _ekle(rapor, i, c, v, 'GEÇERSİZ: ANA HESAP SAYISAL DEĞİL')
        elif not ana.startswith('9') and ana not in gecerli:
            _ekle(rapor, i, c, v, "UYARI: " + ana + " TEK DÜZEN HESAP PLANINDA YOK (kullaniyorsaniz ek_ana_hesaplar'a ekleyin)")
    return df


# ============================================================
# 7) Grup agaci (urunagaci modulu)
# ============================================================
def grup_agaci_kontrol(df, modul, rapor):
    if not ayar(modul, 'grup_agaci_tutarlilik', 'urunagaci'):
        return df
    seviyeler = [c for c in ['Grup1', 'GRUP2', 'GRUP3', 'GRUP4', 'GRUP5', 'GRUP6'] if c in df.columns]
    for n in range(1, len(seviyeler)):
        yollar = {}
        for i in range(len(df)):
            kod = _metin(df.at[i, seviyeler[n]])
            if kod:
                yollar.setdefault(kod, set()).add(tuple(_metin(df.at[i, s]) for s in seviyeler[:n]))
        for i in range(len(df)):
            kod = _metin(df.at[i, seviyeler[n]])
            if kod and len(yollar[kod]) > 1:
                _ekle(rapor, i, seviyeler[n], kod, 'UYARI: ' + seviyeler[n] + ' ' + kod + ' FARKLI ÜST GRUPLARDA (' +
                      ', '.join(sorted('/'.join(y) for y in yollar[kod])[:3]) + ')')
    return df


# ============================================================
# 8) Ortak son kontroller: zorunlu alan, uzunluk, benzer kayit
# ============================================================
def zorunlu_uzunluk_kontrol(df, modul, modul_adi, rapor):
    for c in liste_ayari(ayar(modul, 'zorunlu_alanlar', modul_adi)):
        if c not in df.columns:
            print("  UYARI: Zorunlu kolon '" + c + "' dosyada yok")
            continue
        for i, v in enumerate(df[c].tolist()):
            if _bos(v):
                _ekle(rapor, i, c, '', 'GEÇERSİZ: ZORUNLU ALAN BOŞ')
    for c, n in cift_ayari(ayar(modul, 'alan_uzunluklari', modul_adi)):
        if c not in df.columns or not n.isdigit():
            continue
        for i, v in enumerate(df[c].tolist()):
            s = _metin(v)
            if len(s) > int(n):
                _ekle(rapor, i, c, s, 'GEÇERSİZ: UZUNLUK ' + str(len(s)) + ' > ' + n + ' KARAKTER')


_BENZER_ATILAN = {'LTD', 'STI', 'SIRKETI', 'LIMITED', 'AS', 'A', 'S', 'ANONIM', 'TIC', 'TICARET', 'SAN', 'SANAYI',
                  'VE', 'INS', 'INSAAT', 'PAZ', 'PAZARLAMA', 'ITH', 'IHR', 'ITHALAT', 'IHRACAT', 'DIS', 'LTDSTI'}


def benzer_anahtar(metin):
    kelimeler = re.findall(r'[A-Z0-9]+', katla(metin))
    return ' '.join(k for k in kelimeler if k not in _BENZER_ATILAN)


def benzer_kayitlar(degerler, esik=88, blok_siniri=400):
    """[(i, j, puan)]: yazimi farkli (birebir ayni olmayan) ama benzer degerler."""
    anahtarlar = [benzer_anahtar(d) if not _bos(d) else '' for d in degerler]
    ham = [_metin(d) for d in degerler]
    bloklar = {}
    for i, k in enumerate(anahtarlar):
        if len(k.replace(' ', '')) >= 3:
            bloklar.setdefault(k.replace(' ', '')[:3], []).append(i)
    sonuc = []
    esik = esik / 100.0
    for idx in bloklar.values():
        if len(idx) < 2:
            continue
        for x in range(len(idx)):
            i = idx[x]
            for y in range(x + 1, len(idx)):
                j = idx[y]
                if ham[i] == ham[j]:
                    continue    # birebir ayni: mukerrer kontrolu zaten isaretliyor
                a, b = anahtarlar[i], anahtarlar[j]
                if a == b:
                    sonuc.append((i, j, 100))
                    continue
                if len(idx) > blok_siniri:
                    continue
                sm = difflib.SequenceMatcher(None, a, b)
                if sm.real_quick_ratio() < esik or sm.quick_ratio() < esik:
                    continue
                r = sm.ratio()
                if r >= esik:
                    sonuc.append((i, j, int(round(r * 100))))
    return sonuc


def benzer_kayit_kontrol(df, modul, modul_adi, rapor):
    if not ayar(modul, 'benzer_kayit_kontrol', modul_adi):
        return
    kolonlar = [c for c in liste_ayari(ayar(modul, 'benzer_kayit_sutunu', modul_adi)) if c in df.columns]
    if not kolonlar:
        return
    degerler = [' '.join(_metin(df.at[i, c]) for c in kolonlar).strip() for i in range(len(df))]
    esik = _sayi(ayar(modul, 'benzer_kayit_esik', modul_adi)) or 88
    ciftler = benzer_kayitlar(degerler, esik)
    en_iyi = {}
    for i, j, p in ciftler:
        for a, b in ((i, j), (j, i)):
            if a not in en_iyi or p > en_iyi[a][1]:
                en_iyi[a] = (b, p)
    for i, (j, p) in sorted(en_iyi.items()):
        _ekle(rapor, i, '+'.join(kolonlar), degerler[i],
              'UYARI: OLASI MÜKERRER (satir ' + str(j + 2) + ', %' + str(p) + ': ' + degerler[j][:60] + ')')
    if en_iyi:
        print('  Benzer kayit (olasi mukerrer): ' + str(len(en_iyi)) + ' satir')


# ============================================================
# Akis noktalari
# ============================================================
METIN_TURU_KOLONLAR = ['verginumarasi', 'tckimlikno2', 'TC Kimlik No', 'Hesap NO', 'IBAN No', 'IBAN',
                       'Barkod No', 'GTIP Kodu', 'telefon1', 'telefon2', 'ceptel', 'fax', 'telefon',
                       'Telefon 1', 'Telefon 2', 'Fax', 'postakodu']


def on_islem(df, modul, modul_adi, rapor):
    if ayar(modul, 'bilimsel_gosterim_kontrol', modul_adi):
        df = bilimsel_kontrol(df, METIN_TURU_KOLONLAR + list(modul.get('kod_alanlari', [])), rapor)
    if modul_adi == 'cari':
        df = cari_on_islem(df, modul, rapor)
    elif modul_adi == 'stok':
        df = stok_on_islem(df, modul, rapor)
    elif modul_adi in ('banka', 'personel'):
        df = banka_on_islem(df, modul, modul_adi, rapor)
        if modul_adi == 'personel':
            df = personel_on_islem(df, modul, rapor)
    elif modul_adi == 'hesapplan':
        df = hesapplan_on_islem(df, modul, rapor)
    return df


def son_kontroller(df, modul, modul_adi, rapor):
    zorunlu_uzunluk_kontrol(df, modul, modul_adi, rapor)
    benzer_kayit_kontrol(df, modul, modul_adi, rapor)
    if modul_adi == 'cari':
        df = cari_son_kontrol(df, modul, rapor)
    elif modul_adi == 'stok':
        df = stok_son_kontrol(df, modul, rapor)
    elif modul_adi == 'banka':
        df = banka_son_kontrol(df, modul, rapor)
    elif modul_adi == 'personel':
        df = personel_son_kontrol(df, modul, rapor)
    elif modul_adi == 'demirbas':
        df = demirbas_son_kontrol(df, modul, rapor)
    elif modul_adi == 'hesapplan':
        df = hesapplan_son_kontrol(df, modul, rapor)
    elif modul_adi == 'urunagaci':
        df = grup_agaci_kontrol(df, modul, rapor)
    return df


# ============================================================
# Moduller arasi capraz kontrol
# ============================================================
def _modul_kodlari(kurallar, hedef, kolonlar, klasor):
    m = kurallar.get(hedef) or {}
    for dosya in (m.get('cikti_dosya'), m.get('girdi_dosya')):
        if not dosya:
            continue
        yol = os.path.join(klasor, dosya)
        if not os.path.exists(yol):
            continue
        try:
            d = pd.read_excel(yol, dtype=str)
        except Exception as e:
            print('  UYARI: ' + dosya + ' okunamadi: ' + str(e))
            continue
        d.columns = [str(c).strip() for c in d.columns]
        d = d.rename(columns=m.get('eslesme', {}))
        kodlar = set()
        for c in kolonlar:
            if c in d.columns:
                kodlar |= {kod_isle(v)[0] for v in d[c].tolist() if not _bos(v)}
        return kodlar, dosya
    return None, None


def capraz_kontrol(df, modul, modul_adi, kurallar, klasor):
    rapor = []
    for kaynak, hedef_tanim in cift_ayari(str(ayar(modul, 'capraz_kontrol', modul_adi) or '').replace('>', '=')):
        if kaynak not in df.columns or ':' not in hedef_tanim:
            continue
        hedef, kolonlar = hedef_tanim.split(':', 1)
        kodlar, dosya = _modul_kodlari(kurallar, hedef.strip(), liste_ayari(kolonlar), klasor)
        if kodlar is None:
            print('  Capraz kontrol atlandi: ' + hedef.strip() + ' dosyasi yok')
            continue
        bulunamayan = 0
        for i, v in enumerate(df[kaynak].tolist()):
            k = kod_isle(v)[0] if not _bos(v) else ''
            if k and k not in kodlar:
                bulunamayan += 1
                _ekle(rapor, i, kaynak, v, 'UYARI: ' + str(v) + ' ' + dosya + " DOSYASINDA YOK")
        print('  Capraz kontrol ' + kaynak + ' -> ' + dosya + ': ' + str(bulunamayan) + ' kod bulunamadi')
    return rapor


# ============================================================
# Degisiklik raporu (oncesi / sonrasi)
# ============================================================
def _karsilastir_deger(x):
    if _bos(x):
        return ''
    if isinstance(x, float) and x.is_integer():
        return str(int(x))
    return str(x)


def degisiklikler(once, sonra):
    """Okunan (once) ve temizlenmis (sonra) veriyi KAYNAK_SATIR uzerinden karsilastirir."""
    kolonlar = ['Excel Satır', 'Temiz Satır', 'Kolon', 'Eski Değer', 'Yeni Değer']
    atla = {KAYNAK_SATIR, 'KONTROL_UYARI', 'KONTROL_MUKERRER', 'SIL_ONAY'}
    ortak = [c for c in sonra.columns if c in once.columns and c not in atla]
    kaynak = sonra[KAYNAK_SATIR].tolist()
    eslesen = [j for j, r in enumerate(kaynak) if not _bos(r)]
    satirlar = [int(kaynak[j]) for j in eslesen]
    eski_df = once.drop_duplicates(KAYNAK_SATIR).set_index(KAYNAK_SATIR).reindex(satirlar)
    parcalar = []
    for c in ortak:
        eski = [_karsilastir_deger(x) for x in eski_df[c].tolist()]
        sutun = sonra[c].tolist()
        yeni = [_karsilastir_deger(sutun[j]) for j in eslesen]
        for n, (e, y) in enumerate(zip(eski, yeni)):
            if e != y:
                parcalar.append((satirlar[n], eslesen[n] + 2, c, e, y))
    parcalar.sort(key=lambda x: (x[1], ortak.index(x[2])))
    for j, r in enumerate(kaynak):
        if _bos(r):
            parcalar.append(('', j + 2, '(satır)', '', 'YENİ SATIR EKLENDİ'))
    gorulen = set(satirlar)
    for r in once[KAYNAK_SATIR].tolist():
        if int(r) not in gorulen:
            parcalar.append((int(r), '', '(satır)', '', 'SATIR SİLİNDİ (boş/özet satır)'))
    return pd.DataFrame(parcalar, columns=kolonlar)


def degisiklik_raporu_yaz(fark, sonra, yol, not_siniri=2000, renk_siniri=100000):
    """DEGISIKLIKLER: eski -> yeni listesi.  TEMIZ_VERI: degisen hucreler turuncu, eski deger not olarak."""
    from openpyxl import Workbook
    from openpyxl.comments import Comment
    from openpyxl.styles import PatternFill
    veri = sonra.drop(columns=[KAYNAK_SATIR], errors='ignore')
    wb = Workbook()
    ws_f = wb.active
    ws_f.title = 'DEGISIKLIKLER'
    ws_f.append(list(fark.columns))
    for satir in fark.itertuples(index=False):
        ws_f.append(list(satir))
    for kol, gen in zip('ABCDE', (12, 12, 22, 40, 40)):
        ws_f.column_dimensions[kol].width = gen
    ws = wb.create_sheet('TEMIZ_VERI')
    ws.append([str(c) for c in veri.columns])
    for satir in veri.itertuples(index=False):
        ws.append([None if _bos(v) else v for v in satir])
    turuncu = PatternFill(start_color='FFC000', end_color='FFC000', fill_type='solid')
    kolon_no = {c: n + 1 for n, c in enumerate(veri.columns)}
    notlar = 0
    for n, (temiz, kolon, eski) in enumerate(zip(fark['Temiz Satır'], fark['Kolon'], fark['Eski Değer'])):
        if n >= renk_siniri:
            break
        if kolon not in kolon_no or temiz == '':
            continue
        hucre = ws.cell(row=int(temiz), column=kolon_no[kolon])
        hucre.fill = turuncu
        if notlar < not_siniri:
            hucre.comment = Comment('Eski: ' + str(eski)[:200], 'DIA Temizleme')
            notlar += 1
    wb.save(yol)


# ============================================================
# Esleme hafizasi (musteriler arasi)
# ============================================================
def esleme_hafizasi_yolu(program_klasoru):
    return os.path.join(program_klasoru, 'esleme_hafizasi.json')


def esleme_hafizasi_yukle(program_klasoru):
    try:
        with open(esleme_hafizasi_yolu(program_klasoru), encoding='utf-8') as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


def esleme_hafizasi_kaydet(program_klasoru, modul_adi, eslesme):
    """Musterinin onayli eslemelerini hafizaya ekler. Yeni eklenen sayisini dondurur."""
    hafiza = esleme_hafizasi_yukle(program_klasoru)
    m = hafiza.setdefault(modul_adi, {})
    yeni = 0
    for baslik, alan in eslesme.items():
        if baslik != alan and m.get(baslik) != alan:
            m[baslik] = alan
            yeni += 1
    if yeni:
        try:
            with open(esleme_hafizasi_yolu(program_klasoru), 'w', encoding='utf-8') as f:
                json.dump(hafiza, f, ensure_ascii=False, indent=1, sort_keys=True)
        except OSError as e:
            print('  UYARI: esleme hafizasi yazilamadi: ' + str(e))
            return 0
    return yeni
