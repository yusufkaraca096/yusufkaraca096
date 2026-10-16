import re
import os
import pandas as pd
from datetime import datetime

TR_KARAKTER = str.maketrans('şçığüöŞÇİĞÜÖ', 'sciguoSCIGUO')

# ------------------------------------------------------------
# Parametrik ayarlar (kurallar.xlsx > AYARLAR sayfasi, modul bazinda)
# Excel'de satiri olmayan ayar icin 'varsayilan' kullanilir.
# Ayarlar penceresindeki "Eksik Ayarlari Ekle" butonu bu listeden doldurur.
# ------------------------------------------------------------
AYAR_TANIMLARI = [
    {'ayar': 'telefon_format', 'varsayilan': '10', 'tur': 'secim',
     'secenekler': ['10', '11', '90', '+90'],
     'aciklama': 'Telefon / cep / faks numaralarinin yazim bicimi.',
     'formul': '10 = basta 0 yok | 11 = basta 0 var | 90 = ulke kodlu | +90 = + ile ulke kodlu',
     'ornek': '0532 123 45 67 -> 10: 5321234567 | 11: 05321234567 | +90: +905321234567'},
    {'ayar': 'vkn_sifir_tamamla', 'varsayilan': True, 'tur': 'bool',
     'aciklama': "Excel VKN'yi sayi olarak saklayinca bastaki 0 silinir. 8-9 haneli VKN soldan 0 ile 10 haneye tamamlanir.",
     'formul': 'len(VKN) 8 veya 9 ise -> zfill(10)',
     'ornek': '123456789 -> 0123456789'},
    {'ayar': 'vkn_algoritma_kontrol', 'varsayilan': True, 'tur': 'bool',
     'aciklama': 'VKN 10. hanesini (kontrol hanesi) GIB algoritmasina gore dogrular; hataliysa uyari yazar, degeri degistirmez.',
     'formul': 't=(d[i]+9-i)%10; p=(t*2^(9-i))%9 (t!=0 ve p=0 ise p=9); kontrol=(10-toplam%10)%10',
     'ornek': '0123456789 -> OK'},
    {'ayar': 'tc_algoritma_kontrol', 'varsayilan': True, 'tur': 'bool',
     'aciklama': 'TC Kimlik No algoritma kontrolu; hataliysa uyari yazar, degeri degistirmez.',
     'formul': 'd10=((d1+d3+d5+d7+d9)*7-(d2+d4+d6+d8))%10; d11=(d1+..+d10)%10',
     'ornek': '10000000146 -> OK'},
    {'ayar': 'il_ilce_duzelt', 'varsayilan': True, 'tur': 'bool',
     'aciklama': 'Il/ilceyi 81 il, 973 ilce listesine gore duzeltir: ayni hucredeki il/ilceyi ayirir, yer degistirmisse duzeltir.',
     'formul': "sehir='Kadıköy/İstanbul' veya sehir='Kadıköy', ilce='İstanbul' -> sehir=İSTANBUL, ilce=KADIKÖY",
     'ornek': 'Kadıköy / İstanbul -> İSTANBUL / KADIKÖY'},
    {'ayar': 'il_bos_ise_doldur', 'varsayilan': True, 'tur': 'bool',
     'aciklama': 'Il bos, ilce tek bir ile aitse ili doldurur (Gölbaşı gibi birden fazla ilde olan ilceler doldurulmaz, uyari verilir).',
     'formul': 'ilce yalniz 1 ilde varsa -> sehir = o il',
     'ornek': "ilce='Siverek' -> sehir='ŞANLIURFA'"},
    {'ayar': 'il_ilce_yazim', 'varsayilan': 'BUYUK', 'tur': 'secim',
     'secenekler': ['BUYUK', 'BASLIK'],
     'aciklama': 'Duzeltilen il/ilce adlarinin yazimi.',
     'formul': 'BUYUK = İSTANBUL | BASLIK = İstanbul',
     'ornek': 'kadikoy -> KADIKÖY veya Kadıköy'},
    {'ayar': 'il_sutunu', 'varsayilan': 'sehir', 'tur': 'metin',
     'aciklama': 'Il bilgisinin bulundugu DIA kolonu.',
     'formul': 'kolon adi', 'ornek': 'sehir'},
    {'ayar': 'ilce_sutunu', 'varsayilan': 'ilce', 'tur': 'metin',
     'aciklama': 'Ilce bilgisinin bulundugu DIA kolonu.',
     'formul': 'kolon adi', 'ornek': 'ilce'},
]
AYAR_VARSAYILAN = {a['ayar']: a['varsayilan'] for a in AYAR_TANIMLARI}
AYAR_TURU = {a['ayar']: a['tur'] for a in AYAR_TANIMLARI}
_EVET = ('true', '1', '1.0', 'evet', 'e', 'yes', 'x')


def ayar(modul, ad):
    """Modul ayarini dondurur; tanimli degilse varsayilan."""
    deger = modul.get(ad)
    if deger is None or (isinstance(deger, str) and deger.strip() == ''):
        return AYAR_VARSAYILAN.get(ad)
    return deger


def ayar_cevir(ad, deger):
    """Excel'den okunan ayar degerini turune cevirir."""
    tur = AYAR_TURU.get(ad, 'metin')
    if deger is None or (not isinstance(deger, str) and pd.isna(deger)):
        return AYAR_VARSAYILAN.get(ad)
    if tur == 'bool':
        return str(deger).strip().lower() in _EVET
    s = sayisal_metin(deger)
    if tur == 'secim':
        s = s.upper() if s.upper() in ('BUYUK', 'BASLIK') else s
    return s if s != '' else AYAR_VARSAYILAN.get(ad)


def bosluk_temizle(df):
    for col in df.columns:
        if df[col].dtype == 'object':
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].replace({'nan': '', 'None': ''})
    return df


def birebir_mukerrer_sil(df):
    return df.drop_duplicates()


def kod_mukerrer_sil(df, kod_sutunu):
    if kod_sutunu and kod_sutunu in df.columns:
        return df.drop_duplicates(subset=[kod_sutunu], keep='first')
    return df


def kod_isle(deger):
    if pd.isna(deger) or str(deger).strip() == '':
        return '', 'BOŞ'
    orijinal = str(deger).strip()
    temiz = orijinal.translate(TR_KARAKTER)
    temiz = re.sub(r'\s+', '.', temiz)
    temiz = temiz.upper()
    if orijinal != temiz:
        return temiz, 'DÜZELTİLDİ (' + orijinal + ' -> ' + temiz + ')'
    return temiz, 'OK'


def sayisal_metin(deger):
    """Excel'in sayiya cevirdigi hucreyi metne cevirir: 5321234567.0 -> '5321234567'."""
    if isinstance(deger, float) and deger.is_integer():
        return str(int(deger))
    s = str(deger).strip()
    if re.fullmatch(r'\d+\.0', s):
        s = s[:-2]
    return s


TELEFON_FORMATLARI = ('10', '11', '90', '+90')


def telefon_isle(deger, format='10'):
    """Telefonu 10 haneli cekirdege indirir, sonra istenen bicimde yazar.

    format: '10'  -> 5321234567     (basta 0 yok)
            '11'  -> 05321234567
            '90'  -> 905321234567
            '+90' -> +905321234567
    """
    if pd.isna(deger) or str(deger).strip() == '':
        return '', 'BOŞ'
    orijinal = sayisal_metin(deger)
    rakam = re.sub(r'\D', '', orijinal)
    if not rakam:
        return '', 'GEÇERSİZ (' + orijinal + ')'
    parcalar = re.split(r'/|,|;|\bdahili\b|\bdah\.?|\bext\.?', orijinal, flags=re.IGNORECASE)
    if len(rakam) > 12 and len(parcalar) > 1:
        # '0532 123 45 67 / 102', 'dahili 12' veya iki numara: ilk numara alinir
        ilk, durum = telefon_isle(parcalar[0], format)
        if not durum.startswith('GEÇERSİZ'):
            return ilk, 'UYARI: ILK NUMARA ALINDI (' + orijinal + ' -> ' + ilk + ')'
        return rakam, 'GEÇERSİZ (' + orijinal + ')'
    if rakam.startswith('00'):
        rakam = rakam[2:]
    if len(rakam) == 13 and rakam.startswith('900'):
        rakam = rakam[3:]
    elif len(rakam) == 12 and rakam.startswith('90'):
        rakam = rakam[2:]
    elif len(rakam) == 11 and rakam.startswith('0'):
        rakam = rakam[1:]
    if len(rakam) != 10 or rakam.startswith('0'):
        return rakam, 'GEÇERSİZ (' + orijinal + ')'
    format = str(format).strip()
    if format == '11':
        rakam = '0' + rakam
    elif format == '90':
        rakam = '90' + rakam
    elif format == '+90':
        rakam = '+90' + rakam
    if orijinal != rakam:
        return rakam, 'DÜZELTİLDİ (' + orijinal + ' -> ' + rakam + ')'
    return rakam, 'OK'


def vkn_gecerli(vkn):
    """VKN (10 hane) kontrol hanesi dogrulamasi."""
    if len(vkn) != 10 or not vkn.isdigit():
        return False
    toplam = 0
    for i in range(9):
        t = (int(vkn[i]) + 9 - i) % 10
        p = (t * 2 ** (9 - i)) % 9
        if t != 0 and p == 0:
            p = 9
        toplam += p
    return (10 - toplam % 10) % 10 == int(vkn[9])


def tc_gecerli(tc):
    """TC Kimlik No algoritma kontrolu (11 hane, ilk hane 0 olamaz)."""
    if len(tc) != 11 or not tc.isdigit() or tc[0] == '0':
        return False
    d = [int(c) for c in tc]
    if ((d[0] + d[2] + d[4] + d[6] + d[8]) * 7 - (d[1] + d[3] + d[5] + d[7])) % 10 != d[9]:
        return False
    return sum(d[:10]) % 10 == d[10]


def vkn_isle(deger, sifir_tamamla=True, algoritma_kontrol=True):
    """Excel'in sildigi bastaki sifirlari geri ekler, kontrol hanesine bakar.

    Excel VKN'yi sayi olarak saklayinca 0123456789 -> 123456789 olur.
    8-9 haneli VKN soldan 0 ile 10 haneye tamamlanir.
    Hane sayisi uyarilari vkn_tc_kontrol'de verildigi icin burada tekrar edilmez.
    """
    if pd.isna(deger) or str(deger).strip() == '':
        return '', 'BOŞ'
    orijinal = sayisal_metin(deger)
    rakam = re.sub(r'\D', '', orijinal)
    if sifir_tamamla and 8 <= len(rakam) < 10:
        rakam = rakam.zfill(10)
    if len(rakam) == 10 and algoritma_kontrol and not vkn_gecerli(rakam):
        return rakam, 'UYARI: VKN KONTROL HANESI HATALI (' + rakam + ')'
    if orijinal != rakam:
        return rakam, 'DÜZELTİLDİ (' + orijinal + ' -> ' + rakam + ')'
    return rakam, 'OK'


def tc_isle(deger, algoritma_kontrol=True):
    if pd.isna(deger) or str(deger).strip() == '':
        return '', 'BOŞ'
    orijinal = sayisal_metin(deger)
    rakam = re.sub(r'\D', '', orijinal)
    if len(rakam) == 11 and algoritma_kontrol and not tc_gecerli(rakam):
        return rakam, 'UYARI: TC KIMLIK NO ALGORITMA HATALI (' + rakam + ')'
    if orijinal != rakam:
        return rakam, 'DÜZELTİLDİ (' + orijinal + ' -> ' + rakam + ')'
    return rakam, 'OK'


def mail_gecerli(deger):
    if pd.isna(deger) or str(deger).strip() == '':
        return False
    return bool(re.match(r'[^@]+@[^@]+\.[^@]+', str(deger).strip()))


def mail_isle(deger):
    if pd.isna(deger) or str(deger).strip() == '':
        return '', 'BOŞ'
    orijinal = str(deger).strip()
    orijinal = orijinal.replace('İ', 'i').replace('I', 'i')
    orijinal = orijinal.lower()
    temiz = orijinal.translate(TR_KARAKTER)
    temiz = temiz.replace(' ', '')
    if not mail_gecerli(temiz):
        return temiz, 'HATALI (' + orijinal + ')'
    if orijinal != temiz:
        return temiz, 'DÜZELTİLDİ (' + orijinal + ' -> ' + temiz + ')'
    return temiz, 'OK'


def kdv_normalize(deger):
    if pd.isna(deger) or str(deger).strip() == '':
        return ''
    s = str(deger).replace('%', '').replace(',', '.').strip()
    try:
        f = float(s)
        if 0 < f < 1:
            f = f * 100
        return int(round(f))
    except BaseException:
        return ''


def tarih_isle(deger, format='%d.%m.%Y'):
    if pd.isna(deger) or str(deger).strip() == '':
        return '', 'BOŞ'
    try:
        if isinstance(deger, datetime):
            return deger.strftime(format), 'OK'
        s = str(deger).strip()
        for fmt in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%d.%m.%Y',
                    '%d/%m/%Y', '%m/%d/%Y', '%d-%m-%Y'):
            try:
                d = datetime.strptime(s, fmt)
                return d.strftime(format), 'OK'
            except BaseException:
                continue
        return s, 'FORMAT_BİLİNMİYOR (' + s + ')'
    except BaseException:
        return str(deger), 'HATA'


def iban_mod97_gecerli(iban):
    """IBAN MOD-97 algoritmasi ile gecerlilik kontrolu."""
    if not iban or len(iban) < 5:
        return False
    iban = iban.replace(' ', '').upper()
    yeniden = iban[4:] + iban[:4]
    sayisal = ''
    for ch in yeniden:
        if ch.isdigit():
            sayisal += ch
        elif ch.isalpha():
            sayisal += str(ord(ch) - ord('A') + 10)
        else:
            return False
    try:
        return int(sayisal) % 97 == 1
    except ValueError:
        return False


def iban_isle(deger):
    if pd.isna(deger) or str(deger).strip() == '':
        return '', 'BOŞ'
    temiz = re.sub(r'\s+', '', str(deger).upper())
    if len(temiz) < 5:
        return temiz, 'GEÇERSİZ (çok kısa)'
    if temiz.startswith('TR') and len(temiz) != 26:
        return temiz, 'GEÇERSİZ (' + str(len(temiz)) + ' hane, 26 olmali)'
    if not iban_mod97_gecerli(temiz):
        return temiz, 'GEÇERSİZ (MOD-97 hatasi)'
    return temiz, 'OK'


def tc_kimlik_kontrol(deger):
    if pd.isna(deger) or str(deger).strip() == '':
        return '', 'BOŞ'
    rakam = re.sub(r'\D', '', str(deger))
    if len(rakam) == 11:
        return rakam, 'OK'
    return rakam, 'GEÇERSİZ (' + str(len(rakam)) + ' hane)'


def doviz_standart(deger):
    if pd.isna(deger) or str(deger).strip() == '':
        return 'TL'
    s = str(deger).upper().strip()
    eslesme = {
        'TÜRK LİRASI': 'TL', 'TURK LIRASI': 'TL', 'TRY': 'TL', 'TL': 'TL',
        'DOLAR': 'USD', 'AMERIKAN DOLARI': 'USD', 'USD': 'USD',
        'EURO': 'EUR', 'AVRO': 'EUR', 'EUR': 'EUR',
        'STERLIN': 'GBP', 'POUND': 'GBP', 'GBP': 'GBP',
    }
    return eslesme.get(s, s)


def evet_hayir_standart(deger):
    if pd.isna(deger) or str(deger).strip() == '':
        return 'HAYIR'
    s = str(deger).upper().strip()
    evet = ['EVET', 'E', 'YES', 'Y', 'TRUE', '1', 'VAR']
    hayir = ['HAYIR', 'H', 'NO', 'N', 'FALSE', '0', 'YOK']
    if s in evet:
        return 'EVET'
    if s in hayir:
        return 'HAYIR'
    return s


def buyuk_harf(deger):
    if pd.isna(deger):
        return ''
    return str(deger).strip().replace('i', 'İ').replace('ı', 'I').upper()


def kucuk_harf(deger):
    if pd.isna(deger):
        return ''
    return str(deger).strip().replace('I', 'ı').replace('İ', 'i').lower()


def rapor_ozet(rapor_listesi):
    if not rapor_listesi:
        return pd.DataFrame()
    df = pd.DataFrame(rapor_listesi)
    ozet = {
        'Toplam Uyarı': len(df),
        'Düzeltilen': df['durum'].str.contains('DÜZELTİLDİ', na=False).sum(),
        'Geçersiz': df['durum'].str.contains('GEÇERSİZ|HATALI', na=False).sum(),
        'Boş': (df['durum'] == 'BOŞ').sum(),
    }
    return pd.DataFrame([ozet])


DIA_ALANLARI = {
    'Adı', 'AMORTİSMAN ORANI', 'AMORTİSMAN YIL', 'Ana Birim', 'Banka Adı',
    'Banka adı', 'Barkod No', 'carikartkodu', 'ceptel', 'DEMİRBAŞ ADI',
    'DEMİRBAŞ.KODU', 'Departman', 'Doğum Tarihi', 'Döviz Türü', 'e-mail',
    'eposta', 'FATURA NO', 'FAYDALI ÖMÜR KODU', 'fax', 'Fiyat 1', 'Fiyat 2',
    'Görevi', 'GRUP KODU', 'Grup Kodu', 'grupkodu', 'Hesap Adı', 'Hesap Kodu',
    'Hesap NO', 'IBAN', 'IBAN No', 'ilce', 'İlgili', 'İşe Başlama Tarihi',
    'İşten Ayrılma Tarihi', 'KDV Alış Oran', 'KDV Satış Oran',
    'KDV Satış Toptan Oran', 'Kredi Hesabımı?', 'Maaş', 'Marka', 'MİKTAR',
    'Özel Kod 1', 'Özel Kod 2', 'ozelkod1', 'SATICI ÜNVANI', 'sehir', 'Sicil',
    'Soyadı', 'stok adı', 'stok kart kodu', 'sube', 'ŞUBE ADI', 'TARİH',
    'TC Kimlik No', 'tckimlikno2', 'telefon', 'telefon1', 'telefon2', 'TUTAR',
    'ulke', 'unvan', 'vergidairesi', 'verginumarasi', 'web adresi', 'adres1',
    'adres2',
}


def otomatik_esle(baslik):
    s = str(baslik).strip().translate(TR_KARAKTER).lower()
    for ch in (' ', '.', '_', '-', '/'):
        s = s.replace(ch, '')
    kurallar = [
        (['carikod', 'musterikod', 'musterino', 'hesapkod', 'carino', 'muhasebekod'],
         'carikartkodu'),
        (['carikod', 'carikart'], 'carikartkodu'),
        (['unvan', 'firmaad', 'firmunv', 'musteriad', 'firmaisim', 'ticariunvan',
          'adsoyad', 'firma'],
         'unvan'),
        (['eposta', 'email', 'mail'], 'eposta'),
        (['ceptel', 'gsm', 'cepno', 'cep'], 'ceptel'),
        (['telefon', 'tel', 'telno'], 'telefon1'),
        (['fax', 'faks'], 'fax'),
        (['sehir', 'vilayet'], 'sehir'),
        (['ilce', 'semt', 'kaza'], 'ilce'),
        (['adres', 'acikadres'], 'adres1'),
        (['vergino', 'vkn', 'verginumarasi'], 'verginumarasi'),
        (['vergidaire'], 'vergidairesi'),
        (['tckimlik', 'tckn'], 'tckimlikno2'),
        (['web', 'website', 'internet'], 'web adresi'),
        (['ulke'], 'ulke'),
        (['grupkod'], 'grupkodu'),
        (['sube'], 'sube'),
        (['stokkod', 'stokkodu', 'urunkod'], 'stok kart kodu'),
        (['stokad', 'urunad', 'stokadi', 'urunadi'], 'stok adı'),
        (['birim'], 'Ana Birim'),
        (['kdv'], 'KDV Satış Oran'),
        (['marka'], 'Marka'),
        (['barkod'], 'Barkod No'),
        (['fiyat', 'satis', 'alis'], 'Fiyat 1'),
        (['hesapad'], 'Hesap Adı'),
        (['iban'], 'IBAN No'),
        (['doviz', 'para'], 'Döviz Türü'),
        (['banka'], 'Banka adı'),
        (['soyad'], 'Soyadı'),
        (['maas', 'ucret'], 'Maaş'),
        (['isebaslama', 'isegiris'], 'İşe Başlama Tarihi'),
        (['dogumtarihi', 'dogum'], 'Doğum Tarihi'),
        (['gorevi', 'gorev', 'pozisyon'], 'Görevi'),
        (['departman', 'bolum'], 'Departman'),
        (['demirbas'], 'DEMİRBAŞ ADI'),
        (['tutar'], 'TUTAR'),
        (['faturano'], 'FATURA NO'),
        (['tarih'], 'TARİH'),
        (['miktar', 'adet'], 'MİKTAR'),
    ]
    for kelimeler, alan in kurallar:
        for k in kelimeler:
            if s == k:
                return alan
    for kelimeler, alan in kurallar:
        for k in kelimeler:
            if len(k) < 4:
                if ' ' + k + ' ' in ' ' + s + ' ':
                    return alan
            elif k in s:
                return alan
    return None


def sutun_esle(df, eslesme):
    harita = dict(eslesme) if eslesme else {}
    if eslesme:
        df = df.rename(columns=eslesme)
    taninmayanlar = []
    for col in list(df.columns):
        if col in DIA_ALANLARI:
            continue
        tahmin = otomatik_esle(col)
        if tahmin and tahmin not in df.columns:
            df = df.rename(columns={col: tahmin})
            harita[col] = tahmin
            print("  Otomatik eslesti: '" + str(col) + "' -> '" + str(tahmin) + "'")
        elif tahmin:
            taninmayanlar.append(str(col) + ' (-> ' + str(tahmin) + ' zaten var)')
        else:
            taninmayanlar.append(str(col))
    if taninmayanlar:
        print('  TANINMAYAN: ' + ', '.join(taninmayanlar))
    df.attrs['esleme_haritasi'] = harita
    return df


def mukerrer_kontrol_et(df, kurallar_listesi, mod='otomatik'):
    if 'KONTROL_MUKERRER' not in df.columns:
        df['KONTROL_MUKERRER'] = ''
    df = df.reset_index(drop=True)

    def excel_satir(idx):
        return idx + 2

    for kural in kurallar_listesi:
        sutunlar = kural.get('sutunlar', [])
        if not all(s in df.columns for s in sutunlar):
            continue
        gecerli = pd.Series([True] * len(df), index=df.index)
        for s in sutunlar:
            gecerli = gecerli & df[s].notna() & (df[s].astype(str).str.strip() != '')
        anahtar = df[sutunlar[0]].astype(str)
        for s in sutunlar[1:]:
            anahtar = anahtar + '|' + df[s].astype(str)
        anahtar_gecerli = anahtar.where(gecerli, '')
        tekrar = anahtar_gecerli.value_counts()
        isaretli = 0
        for idx, val in anahtar_gecerli.items():
            if val and tekrar[val] > 1:
                ayni = [excel_satir(i)
                        for i in anahtar_gecerli[anahtar_gecerli == val].index if i != idx]
                if not ayni:
                    continue
                ayni_str = ', '.join(str(x) for x in ayni[:5])
                if len(ayni) > 5:
                    ayni_str += ' ...(+' + str(len(ayni) - 5) + ')'
                uyari = ('MÜKERRER: ' + ' + '.join(sutunlar) + ' (satir: ' + ayni_str + ')')
                mevcut = df.at[idx, 'KONTROL_MUKERRER']
                if uyari not in mevcut:
                    if mevcut == '':
                        df.at[idx, 'KONTROL_MUKERRER'] = uyari
                    else:
                        df.at[idx, 'KONTROL_MUKERRER'] = mevcut + ' | ' + uyari
                    isaretli += 1
        if isaretli > 0:
            print('  Isaretlendi (' + ' + '.join(sutunlar) + '): ' + str(isaretli) + ' satir')

    silinen_df = pd.DataFrame()
    if mod == 'sadece_isaretle':
        return df, silinen_df
    silinecek_indeksler = set()
    for kural in kurallar_listesi:
        sutunlar = kural.get('sutunlar', [])
        aksiyon = kural.get('aksiyon', 'isaretle')
        if mod == 'sadece_sil' and aksiyon != 'sil':
            continue
        if aksiyon != 'sil':
            continue
        if len(sutunlar) < 2:
            continue
        if not all(s in df.columns for s in sutunlar):
            continue
        gecerli = pd.Series([True] * len(df), index=df.index)
        for s in sutunlar:
            gecerli = gecerli & df[s].notna() & (df[s].astype(str).str.strip() != '')
        df_gecerli = df[gecerli]
        onceki_indeksler = df_gecerli.index.tolist()
        gecici = df_gecerli.drop_duplicates(subset=sutunlar, keep='first')
        kalan_indeksler = set(gecici.index.tolist())
        bu_kural_silinecek = [i for i in onceki_indeksler if i not in kalan_indeksler]
        silinecek_indeksler.update(bu_kural_silinecek)
        if bu_kural_silinecek:
            print('  Silinecek (' + ' + '.join(sutunlar) + '): ' + str(len(bu_kural_silinecek)) + ' satir')
    if silinecek_indeksler:
        silinen_df = df.loc[list(silinecek_indeksler)].copy()
        silinen_df['SILINME_NEDENI'] = 'Mukerrer kural'
        df = df.drop(index=list(silinecek_indeksler))
    return df, silinen_df


def mukerrer_grup_raporu(df, kurallar_listesi):
    """Her mukerrer kurali icin grup ozeti uretir."""
    satirlar = []
    for kural in kurallar_listesi:
        sutunlar = kural.get('sutunlar', [])
        if not sutunlar or not all(s in df.columns for s in sutunlar):
            continue
        gecerli = pd.Series([True] * len(df), index=df.index)
        for s in sutunlar:
            gecerli = gecerli & df[s].notna() & (df[s].astype(str).str.strip() != '')
        df_g = df[gecerli]
        if len(df_g) == 0:
            continue
        anahtar = df_g[sutunlar[0]].astype(str)
        for s in sutunlar[1:]:
            anahtar = anahtar + ' | ' + df_g[s].astype(str)
        gruplar = anahtar.groupby(anahtar)
        for deger, indeksler in gruplar.groups.items():
            if len(indeksler) < 2:
                continue
            excel_satirlar = [i + 2 for i in indeksler]
            satirlar.append({
                'kural': ' + '.join(sutunlar),
                'deger': deger[:100],
                'satir_sayisi': len(indeksler),
                'satir_numaralari': ', '.join(str(x) for x in excel_satirlar[:10]) + (' ...' if len(excel_satirlar) > 10 else ''),
            })
    if not satirlar:
        return pd.DataFrame()
    df_rapor = pd.DataFrame(satirlar)
    df_rapor = df_rapor.sort_values(['kural', 'satir_sayisi'], ascending=[True, False])
    return df_rapor


def kdv_gecerli_mi(deger):
    """Turkiye'de gecerli KDV oranlari: 0, 1, 8, 10, 18, 20."""
    if deger is None or str(deger).strip() == '':
        return True
    try:
        v = int(float(str(deger).replace(',', '.')))
        return v in (0, 1, 8, 10, 18, 20)
    except (ValueError, TypeError):
        return False


def tarih_karsilastir(df, alan_sirasi):
    """Alan sirasina gore tarihleri karsilastirir (eski -> yeni).
Ornek: ['Dogum Tarihi', 'Ise Baslama Tarihi', 'Isten Ayrilma Tarihi']
"""
    if 'KONTROL_UYARI' not in df.columns:
        df['KONTROL_UYARI'] = 'OK'

    def parse(t):
        if pd.isna(t) or str(t).strip() == '':
            return None
        try:
            return pd.to_datetime(str(t), dayfirst=True, errors='coerce')
        except Exception:
            return None

    uyari_sayisi = 0
    for idx, row in df.iterrows():
        tarihler = []
        for alan in alan_sirasi:
            if alan not in df.columns:
                tarihler.append((alan, None))
                continue
            t = parse(row.get(alan))
            tarihler.append((alan, t))

        hatalar = []
        for i in range(len(tarihler) - 1):
            a_adi, a = tarihler[i]
            b_adi, b = tarihler[i + 1]
            if a is None or b is None:
                continue
            if pd.notna(a) and pd.notna(b) and a > b:
                hatalar.append(a_adi + ' > ' + b_adi)

        if hatalar:
            mevcut = df.at[idx, 'KONTROL_UYARI']
            yeni = 'TARIH HATASI: ' + ', '.join(hatalar)
            if mevcut == 'OK' or mevcut == '':
                df.at[idx, 'KONTROL_UYARI'] = yeni
            elif yeni not in mevcut:
                df.at[idx, 'KONTROL_UYARI'] = mevcut + ' | ' + yeni
            uyari_sayisi += 1

    if uyari_sayisi > 0:
        print('  Tarih sirasi hatasi: ' + str(uyari_sayisi) + ' satir')
    return df


def kod_seviye_kontrol(df, alan_sirasi, ad='kod'):
    """Bir dizi kod alaninda 'soldan dolu' kurali kontrol eder.
Ornek: [Grup1, Grup2, Grup3] -> Grup1 bos, Grup2 dolu olamaz.
"""
    if 'KONTROL_UYARI' not in df.columns:
        df['KONTROL_UYARI'] = 'OK'

    uyari_sayisi = 0
    for idx, row in df.iterrows():
        dolu = []
        for alan in alan_sirasi:
            if alan not in df.columns:
                dolu.append(False)
                continue
            v = row.get(alan)
            dolu.append(not (pd.isna(v) or str(v).strip() == ''))

        bulundu_bos = False
        hata_alan = None
        for i, d in enumerate(dolu):
            if d and bulundu_bos:
                hata_alan = alan_sirasi[i]
                break
            elif not d:
                bulundu_bos = True

        if hata_alan:
            mevcut = df.at[idx, 'KONTROL_UYARI']
            yeni = 'SEVIYE ATLAMA: ' + hata_alan + ' dolu ama onceki bos'
            if mevcut == 'OK' or mevcut == '':
                df.at[idx, 'KONTROL_UYARI'] = yeni
            elif yeni not in mevcut:
                df.at[idx, 'KONTROL_UYARI'] = mevcut + ' | ' + yeni
            uyari_sayisi += 1

    if uyari_sayisi > 0:
        print('  ' + ad + ' seviye atlama: ' + str(uyari_sayisi) + ' satir')
    return df


def sil_onay_sutunu_ekle(df):
    if 'SIL_ONAY' not in df.columns:
        df['SIL_ONAY'] = ''
    return df


def onayli_sil(df):
    if 'SIL_ONAY' not in df.columns:
        print('  UYARI: SIL_ONAY sutunu yok.')
        return df, pd.DataFrame()
    onay = df['SIL_ONAY'].astype(str).str.upper().str.strip()
    onayli = onay.isin(['E', 'EVET', 'YES', 'Y', '1', 'TRUE'])
    silinen_df = df[onayli].copy()
    temiz_df = df[~onayli].copy()
    temiz_df = temiz_df.drop(columns=['SIL_ONAY'], errors='ignore')
    print('  Onayli silinen: ' + str(len(silinen_df)) + ' satir')
    print('  Kalan: ' + str(len(temiz_df)) + ' satir')
    return temiz_df, silinen_df


def excel_renklendir(dosya_yolu):
    """KONTROL_* ve SIL_ONAY sutunlarina gore Excel'i renklendirir."""
    try:
        from openpyxl import load_workbook
        from openpyxl.styles import PatternFill
        wb = load_workbook(dosya_yolu)
        ws = wb.active

        sari = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
        yesil = PatternFill(start_color='90EE90', end_color='90EE90', fill_type='solid')
        turuncu = PatternFill(start_color='FFC000', end_color='FFC000', fill_type='solid')
        kirmizi = PatternFill(start_color='FF6B6B', end_color='FF6B6B', fill_type='solid')

        basliklar = [c.value for c in ws[1]]
        k_u = None; k_m = None; s_o = None
        for i, b in enumerate(basliklar, start=1):
            if b == 'KONTROL_UYARI': k_u = i
            elif b == 'KONTROL_MUKERRER': k_m = i
            elif b == 'SIL_ONAY': s_o = i

        if not k_u and not k_m and not s_o:
            return False

        for row in ws.iter_rows(min_row=2):
            if k_m:
                c = row[k_m - 1]
                if c.value:
                    c.fill = sari
            if s_o:
                c = row[s_o - 1]
                if c.value and str(c.value).upper().strip() in ('E', 'EVET', 'YES', 'Y', '1', 'TRUE'):
                    c.fill = yesil
            if k_u:
                c = row[k_u - 1]
                if c.value:
                    v = str(c.value)
                    if 'GEÇERSİZ' in v or 'HATALI' in v:
                        c.fill = kirmizi
                    elif 'DÜZELTİLDİ' in v or 'MÜKERRER' in v:
                        c.fill = turuncu
        wb.save(dosya_yolu)
        return True
    except Exception as e:
        print('  UYARI: Renklendirme yapilamadi: ' + str(e))
        return False


def inceleme_raporu(df, cikti_yolu):
    try:
        from openpyxl import load_workbook
        from openpyxl.styles import PatternFill
        df.to_excel(cikti_yolu, index=False)
        wb = load_workbook(cikti_yolu)
        ws = wb.active
        sari = PatternFill(
            start_color='FFFF00',
            end_color='FFFF00',
            fill_type='solid')
        yesil = PatternFill(
            start_color='90EE90',
            end_color='90EE90',
            fill_type='solid')
        basliklar = [c.value for c in ws[1]]
        for i, baslik in enumerate(basliklar, start=1):
            if baslik == 'KONTROL_MUKERRER':
                for row in ws.iter_rows(min_row=2, min_col=i, max_col=i):
                    for cell in row:
                        if cell.value:
                            cell.fill = sari
            if baslik == 'SIL_ONAY':
                for row in ws.iter_rows(min_row=2, min_col=i, max_col=i):
                    for cell in row:
                        cell.fill = yesil
        wb.save(cikti_yolu)
    except Exception as e:
        print('  UYARI: Renklendirme yapilamadi: ' + str(e))
        df.to_excel(cikti_yolu, index=False)


def bos_satir_temizle(df):
    """Sadece TAMAMEN bos satirlari ve ozet satirlarini siler."""
    onceki = len(df)

    # Tamamen bos satirlar
    df = df.dropna(how='all')

    # Ilk sutunda ozet ifadesi olan satirlar
    if len(df) > 0 and len(df.columns) > 0:
        ilk_sutun = df.columns[0]

        def ozet_satiri(deger):
            if pd.isna(deger):
                return False
            s = str(deger).strip().upper()
            anahtarlar = ['TOPLAM', 'GENEL TOPLAM', 'ARA TOPLAM', 'LISTELENEN', '***', '---', '===']

            for a in anahtarlar:
                if s == a or s.startswith(a):
                    return True
            return False

        df = df[~df[ilk_sutun].apply(ozet_satiri)]

    df = df.reset_index(drop=True)
    silinen = onceki - len(df)
    if silinen > 0:
        print('  Bos/ozet satir silindi: ' + str(silinen))
    return df


def coklu_bosluk_temizle(df):
    duzeltilen = 0
    for col in df.columns:
        if df[col].dtype == 'object':
            def temizle(x):
                nonlocal duzeltilen
                if pd.isna(x) or str(x) == 'nan':
                    return ''
                s = str(x)
                yeni = re.sub(r'\s+', ' ', s).strip()
                if s != yeni:
                    duzeltilen += 1
                return yeni
            df[col] = df[col].apply(temizle)
    if duzeltilen > 0:
        print('  Coklu bosluk duzeltildi: ' + str(duzeltilen) + ' hucre')
    return df


def bos_sutun_sil(df):
    bos = []
    for col in df.columns:
        if df[col].isna().all():
            bos.append(col)
        elif (df[col].astype(str).str.strip() == '').all():
            bos.append(col)
    if bos:
        print('  Bos sutunlar silindi: ' + ', '.join(str(s) for s in bos))
        df = df.drop(columns=bos)
    return df


def kod_bas_son_temizle(df, kod_sutunlari):
    duzeltilen = 0
    for alan in kod_sutunlari:
        if alan in df.columns:
            def temizle(x):
                nonlocal duzeltilen
                if pd.isna(x) or str(x) == 'nan':
                    return ''
                s = str(x).strip()
                yeni = re.sub(r'^[\s\.\-,;:]+|[\s\.\-,;:]+$', '', s)
                if s != yeni:
                    duzeltilen += 1
                return yeni
            df[alan] = df[alan].apply(temizle)
    if duzeltilen > 0:
        print('  Kod bas/son temizlendi: ' + str(duzeltilen) + ' hucre')
    return df


def posta_kodu_isle(deger):
    if pd.isna(deger) or str(deger).strip() == '':
        return ''
    s = re.sub(r'\D', '', str(deger))
    if not s:
        return ''
    return s.zfill(5)


def sayi_format_duzelt(deger):
    if pd.isna(deger) or str(deger).strip() == '':
        return ''
    s = str(deger).strip()
    if ',' in s:
        s = s.replace('.', '')
        s = s.replace(',', '.')
    try:
        return float(s)
    except BaseException:
        return deger


def sutun_sirala(df, modul):
    SIRALAR = {
        'cari': ['cari kayitturu', 'carikarttipi', 'carikartkodu', 'doviz', 'eposta',
                 'grupkodu', 'ilgili', 'indirimorani', 'ozelkod1', 'sube', 'unvan',
                 'vergidairesi', 'verginumarasi', 'tckimlikno2', 'web adresi', 'yurtdisi',
                 'adresadi', 'adres1', 'adres2', 'ilce', 'postakodu', 'sehir', 'adresno',
                 'adrestipi', 'anaadres', 'ceptel', 'fax', 'telefon1', 'telefon2', 'ulke'],
        'stok': ['Stok kart türü', 'stok kart kodu', 'stok adı', 'KDV Alış Oran',
                 'KDV Satış Oran', 'KDV Satış Toptan Oran', 'Özel Kod 1', 'Özel Kod 2',
                 'Özel Kod 3', 'Özel Kod 4', 'Özel Kod 5', 'Özel Kod 6', 'Özel Kod 7',
                 'Özel Kod 8', 'Özel Kod 9', 'Özel Kod 10', 'Yetki kodu',
                 'Mağaza Ürün Grubu', 'Marka', 'Not', 'Birim Kodu', 'Ana Birim',
                 'Fiyat 1', 'Fiyat 2', 'Fiyat 3', 'Fiyat 4', 'Fiyat 5', 'Fiyat 6',
                 'Fiyat 7', 'Fiyat 8', 'Fiyat 9', 'Fiyat 10', 'Döviz 1', 'Döviz 2',
                 'Döviz 3', 'Döviz 4', 'Döviz 5', 'Döviz 6', 'Döviz 7', 'Döviz 8',
                 'Döviz 9', 'Döviz 10', 'Resim URL', 'Barkod No', 'Azami Miktar',
                 'Asgari Miktar', 'GTIP Kodu'],
        'banka': ['Banka adı', 'Hesap Kodu', 'Hesap Adı', 'Hesap NO', 'IBAN No',
                  'Yetki Kodu', 'Özel Kod', 'Döviz Türü', 'Kredi Hesabımı?', 'İlgili',
                  'e-mail', 'telefon'],
        'personel': ['Sicil', 'Adı', 'Soyadı', 'TC Kimlik No', 'Doğum Tarihi',
                     'İşe Başlama Tarihi', 'İşten Ayrılma Tarihi', 'Durum', 'Görevi',
                     'Departman', 'Maaş', 'Banka Adı', 'IBAN', 'telefon1', 'Telefon 1',
                     'Telefon 2', 'Fax', 'eposta', 'E-Posta'],
    }
    siralama = SIRALAR.get(modul)
    if not siralama:
        return df
    mevcut = [s for s in siralama if s in df.columns]
    digerleri = [s for s in df.columns if s not in siralama]
    if mevcut:
        df = df[mevcut + digerleri]
        print('  Sutunlar DIA sirasina dizildi.')
    return df


def temizlik_uygula(df, modul, temizlik_kurallari, kod_alanlari, sayi_alanlari):
    aktif = {}
    for k in temizlik_kurallari:
        aktif[k['kural']] = k.get('aktif', True)
    if aktif.get('bos_satir_temizle', True):
        df = bos_satir_temizle(df)
    if aktif.get('toplam_satirlarini_temizle', True):
        df = toplam_satirlarini_temizle(df)
    if aktif.get('coklu_bosluk', True):
        df = coklu_bosluk_temizle(df)
    if aktif.get('bos_sutun_sil', True):
        df = bos_sutun_sil(df)
    if aktif.get('kod_bas_son_temizle', True):
        df = kod_bas_son_temizle(df, kod_alanlari)
    if aktif.get('posta_kodu_5_hane', False) and 'postakodu' in df.columns:
        df['postakodu'] = df['postakodu'].apply(posta_kodu_isle)
    if aktif.get('sayi_format_duzelt', False):
        for alan in sayi_alanlari:
            if alan in df.columns:
                df[alan] = df[alan].apply(sayi_format_duzelt)
    if aktif.get('sutun_sirasi', False):
        df = sutun_sirala(df, modul)
    return df


def excel_oku_coklu_sayfa(dosya_yolu, sayfa_ayari):
    xl = pd.ExcelFile(dosya_yolu)
    tum_sayfalar = xl.sheet_names

    if not sayfa_ayari or str(sayfa_ayari).strip() in ('', '0', 'None', 'nan'):
        print('  Sayfa belirtilmemis, ilk sayfa okunuyor: ' +
              str(tum_sayfalar[0]))
        return pd.read_excel(dosya_yolu, sheet_name=tum_sayfalar[0])

    istenen = [s.strip() for s in str(sayfa_ayari).split(',') if s.strip()]
    gecerli = [s for s in istenen if s in tum_sayfalar]

    if not gecerli:
        print('  UYARI: Belirtilen sayfalar bulunamadi: ' + str(istenen))
        print("  Excel'deki sayfalar: " + str(tum_sayfalar))
        print('  Ilk sayfa okunuyor: ' + str(tum_sayfalar[0]))
        return pd.read_excel(dosya_yolu, sheet_name=tum_sayfalar[0])

    bulunmayan = [s for s in istenen if s not in tum_sayfalar]
    if bulunmayan:
        print('  UYARI: Su sayfalar bulunamadi, atlaniyor: ' + str(bulunmayan))

    if len(gecerli) == 1:
        print('  Sayfa okunuyor: ' + gecerli[0])
        return pd.read_excel(dosya_yolu, sheet_name=gecerli[0])

    print('  ' +
          str(len(gecerli)) +
          ' sayfa okunuyor ve birlestiriliyor: ' +
          ', '.join(gecerli))
    parcalar = []
    for s in gecerli:
        df_s = pd.read_excel(dosya_yolu, sheet_name=s)
        df_s['_KAYNAK_SAYFA'] = s
        parcalar.append(df_s)
        print('    ' + s + ': ' + str(len(df_s)) + ' satir')
    birlesik = pd.concat(parcalar, ignore_index=True)
    birlesik = birlesik.drop(columns=['_KAYNAK_SAYFA'], errors='ignore')
    print('  Toplam birlestirilmis satir: ' + str(len(birlesik)))
    return birlesik


def kurallari_excelden_oku(yol):
    moduller = ['cari', 'stok', 'banka', 'personel', 'demirbas', 'hesapplan', 'muhkod', 'urunagaci']

    K = {'global': {'bosluk_temizle': True, 'birebir_mukerrer_sil': True,
                    'kod_mukerrer_sil': True, 'turkce_karakter_kodlarda': True,
                    'rapor_olustur': True}}
    for m in moduller:
        K[m] = {'mukerrer_kontrol': [], 'eslesme': {}, 'varsayilan': {},
                'telefon_alanlari': [], 'buyuk_harf_alanlari': [],
                'kucuk_harf_alanlari': [], 'kod_alanlari': [],
                'kdv_alanlari': [], 'tarih_alanlari': []}

    # MUKERRER
    try:
        df = pd.read_excel(yol, sheet_name='MUKERRER')
        for _, row in df.iterrows():
            modul = str(row.get('modul', '')).strip()
            if modul in K:
                K[modul]['mukerrer_kontrol'].append({
                    'sutunlar': [s.strip() for s in str(row.get('sutunlar', '')).split(',') if s.strip()],
                    'aksiyon': str(row.get('aksiyon', 'isaretle')).strip().lower()
                })
    except Exception as e:
        print('MUKERRER okunamadi: ' + str(e))

    # ESLESME
    try:
        df = pd.read_excel(yol, sheet_name='ESLESME')
        for _, row in df.iterrows():
            modul = str(row.get('modul', '')).strip()
            if modul in K:
                baslik = str(row.get('excel_baslik', '')).strip()
                alan = str(row.get('dia_alan', '')).strip()
                if baslik and alan:
                    K[modul]['eslesme'][baslik] = alan
    except Exception as e:
        print('ESLESME okunamadi: ' + str(e))

    # VARSAYILAN
    try:
        df = pd.read_excel(yol, sheet_name='VARSAYILAN')
        for _, row in df.iterrows():
            modul = str(row.get('modul', '')).strip()
            if modul in K:
                alan = str(row.get('alan', '')).strip()
                deger = row.get('deger', '')
                if not isinstance(deger, str) and pd.isna(deger):
                    deger = ''
                if alan:
                    K[modul]['varsayilan'][alan] = deger
    except Exception as e:
        print('VARSAYILAN okunamadi: ' + str(e))

    # AYARLAR
    try:
        df = pd.read_excel(yol, sheet_name='AYARLAR')
        liste_alanlari = ['telefon_alanlari', 'buyuk_harf_alanlari',
                          'kucuk_harf_alanlari', 'kod_alanlari',
                          'kdv_alanlari', 'tarih_alanlari']
        tek_alanlar = ['sayfa', 'girdi_dosya', 'cikti_dosya', 'rapor_dosya']
        bool_alanlari = ['mail_kontrol', 'iban_temizle', 'doviz_standart',
                         'tc_kimlik_11_hane', 'tum_alanlar_zorunlu', 'kist_kontrol',
                         'bosluk_yerine_nokta', 'cari_hesap_120_320']

        for _, row in df.iterrows():
            modul = str(row.get('modul', '')).strip()
            if modul not in K:
                continue
            ayar = str(row.get('ayar', '')).strip()
            deger = row.get('deger', '')
            if ayar in tek_alanlar:
                K[modul][ayar] = str(deger)
            elif ayar in liste_alanlari:
                if pd.isna(deger) or str(deger).strip() == '':
                    K[modul][ayar] = []
                else:
                    K[modul][ayar] = [s.strip()
                                      for s in str(deger).split(',') if s.strip()]
            elif ayar in bool_alanlari:
                K[modul][ayar] = str(deger).strip().lower() in _EVET
            elif ayar in AYAR_TURU:
                K[modul][ayar] = ayar_cevir(ayar, deger)
    except Exception as e:
        print('AYARLAR okunamadi: ' + str(e))

    return K


def hiyerarsi_kontrol(df, hesap_kodu_sutunu):
    """Her hesabin ust hesabi listede var mi kontrol eder.
Yoksa KONTROL_UYARI sutununa yazar.
"""
    if hesap_kodu_sutunu not in df.columns:
        return df
    if 'KONTROL_UYARI' not in df.columns:
        df['KONTROL_UYARI'] = 'OK'

    mevcut = set()
    for kod in df[hesap_kodu_sutunu]:
        if pd.notna(kod) and str(kod).strip():
            mevcut.add(str(kod).strip())

    uyari_sayisi = 0
    for idx, kod in enumerate(df[hesap_kodu_sutunu]):
        k = str(kod).strip() if pd.notna(kod) else ''
        if not k:
            continue
        parcalar = k.split('.')
        if len(parcalar) < 2:
            continue
        parent = '.'.join(parcalar[:-1])
        if parent not in mevcut:
            mevcut_uyari = df.at[idx, 'KONTROL_UYARI']
            yeni = 'UST HESAP YOK: ' + parent
            if mevcut_uyari == 'OK' or mevcut_uyari == '':
                df.at[idx, 'KONTROL_UYARI'] = yeni
            elif yeni not in mevcut_uyari:
                df.at[idx, 'KONTROL_UYARI'] = mevcut_uyari + ' | ' + yeni
            uyari_sayisi += 1

    if uyari_sayisi > 0:
        print(
            '  Hiyerarsi uyarisi: ' +
            str(uyari_sayisi) +
            ' satirda ust hesap eksik'
        )
    return df


EK_ALANLAR = {
    'Asgari Miktar', 'Azami Miktar', 'Birim Kodu', 'Dogum Tarihi', 'Durum',
    'Döviz 1', 'Döviz 2', 'Döviz 3', 'Döviz 4', 'Döviz 5', 'Döviz 6', 'Döviz 7',
    'Döviz 8', 'Döviz 9', 'Döviz 10', 'Fax', 'Fiyat 3', 'Fiyat 4', 'Fiyat 5',
    'Fiyat 6', 'Fiyat 7', 'Fiyat 8', 'Fiyat 9', 'Fiyat 10', 'Gorevi', 'GRUP KODU',
    'Grup1', 'GRUP2', 'GRUP3', 'GRUP4', 'GRUP5', 'GRUP6', 'GTIP Kodu',
    'Ise Baslama Tarihi', 'Isten Ayrilma Tarihi', 'Maas', 'Mağaza Ürün Grubu',
    'Muhasebe Hesap Kodu', 'Not', 'Özel Kod 1', 'Özel Kod 2', 'Özel Kod 3',
    'Özel Kod 4', 'Özel Kod 5', 'Özel Kod 6', 'Özel Kod 7', 'Özel Kod 8',
    'Özel Kod 9', 'Özel Kod 10', 'Resim URL', 'Sicil', 'STOK OZEL KOD 3',
    'STOK OZEL KOD 4', 'STOK OZEL KOD 5', 'STOK ÖZEL KOD 1', 'STOK ÖZEL KOD 2',
    'STOK ÖZEL KOD 3', 'STOK ÖZEL KOD 4', 'STOK ÖZEL KOD 5', 'Telefon 1',
    'Telefon 2', 'Yetki kodu',
}
DIA_ALANLARI = DIA_ALANLARI | EK_ALANLAR


def xls_olarak_kaydet(df, xlsx_yolu):
    """DataFrame'i hem .xlsx hem .xls olarak kaydeder."""
    try:
        import xlwt
        xls_yolu = xlsx_yolu.replace('.xlsx', '.xls')
        wb = xlwt.Workbook(encoding='utf-8')
        ws = wb.add_sheet('Sheet1')

        for j, col in enumerate(df.columns):
            ws.write(0, j, str(col))

        for i, (_, row) in enumerate(df.iterrows()):
            for j, val in enumerate(row):
                if pd.isna(val):
                    val = ''
                ws.write(i + 1, j, str(val))
        wb.save(xls_yolu)
        print('  XLS kaydedildi: ' + os.path.basename(xls_yolu))
        return xls_yolu
    except ImportError:
        print("  UYARI: xlwt kurulu degil. 'python -m pip install xlwt' calistir.")
        return None
    except Exception as e:
        print('  UYARI: XLS kaydedilemedi: ' + str(e))
        return None


def sayisal_metin_temizle(df, sutunlar):
    """TC, vergi no gibi sutunlardaki .0 uzantisini temizler."""
    for alan in sutunlar:
        if alan in df.columns:
            def temizle(x):
                if pd.isna(x):
                    return ''
                s = str(x).strip()
                if s.endswith('.0'):
                    s = s[:-2]
                return s
            df[alan] = df[alan].apply(temizle)
    return df


def vkn_tc_kontrol(df, vkn_sutun='verginumarasi', tc_sutun='tckimlikno2'):
    """VKN 10 hane, TC 11 hane kontrolu yapar.
Karisikliklari ve eksikleri tespit eder.
"""
    if 'KONTROL_UYARI' not in df.columns:
        df['KONTROL_UYARI'] = 'OK'

    def sadece_rakam(x):
        if pd.isna(x):
            return ''
        return re.sub(r'\D', '', str(x))

    for idx, row in df.iterrows():
        vkn = sadece_rakam(row.get(vkn_sutun, '')) if vkn_sutun in df.columns else ''
        tc = sadece_rakam(row.get(tc_sutun, '')) if tc_sutun in df.columns else ''

        uyarilar = []

        if vkn and tc:
            uyarilar.append('HEM VKN HEM TC DOLU')

        if vkn:
            if len(vkn) == 10:
                pass
            elif len(vkn) == 11:
                uyarilar.append('VKN 11 HANE (TC OLABILIR)')
            else:
                uyarilar.append('VKN ' + str(len(vkn)) + ' HANE (10 OLMALI)')

        if tc:
            if len(tc) == 11:
                pass
            elif len(tc) == 10:
                uyarilar.append('TC 10 HANE (VKN OLABILIR)')
            else:
                uyarilar.append('TC ' + str(len(tc)) + ' HANE (11 OLMALI)')

        if uyarilar:
            mevcut = df.at[idx, 'KONTROL_UYARI']
            yeni = ' | '.join(uyarilar)
            if mevcut == 'OK' or mevcut == '':
                df.at[idx, 'KONTROL_UYARI'] = yeni
            elif yeni not in mevcut:
                df.at[idx, 'KONTROL_UYARI'] = mevcut + ' | ' + yeni

    return df


def toplam_satirlarini_temizle(df):
    """Listelenen, TOPLAM, *** gibi ozet satirlarini siler."""
    onceki = len(df)
    if len(df.columns) == 0:
        return df
    ilk_sutun = df.columns[0]
    def ozet_satiri(deger):
        if pd.isna(deger):
            return False
        s = str(deger).strip().upper()
        if s.startswith('LISTELENEN') or s.startswith('TOPLAM'):
            return True
        if s.startswith('GENEL TOPLAM') or s.startswith('ARA TOPLAM'):
            return True
        if s.startswith('***') or s.startswith('---'):
            return True
        return False
    df = df[~df[ilk_sutun].apply(ozet_satiri)]
    df = df.dropna(how='all').reset_index(drop=True)
    silinen = onceki - len(df)
    if silinen > 0:
        print('  Ozet satirlari silindi: ' + str(silinen))
    return df


def cari_kayitturu_ata(df):
    """Cari kayitturu atar. Oncelik:
1. VKN 10 hane ise -> KRLS (sirket)
2. TC 11 hane ise -> SHS (sahis)
3. Ikisi de yoksa -> unvanda LTD/AS/STI ara -> KRLS, yoksa SHS
"""
    if 'cari kayitturu' not in df.columns:
        df['cari kayitturu'] = 'SHS'

    def sadece_rakam(x):
        if pd.isna(x):
            return ''
        return re.sub(r'\D', '', str(x))

    sirket_kelimeleri = ['LTD', 'LIMITED', 'A.S', 'AS', 'ANONIM', 'STI', 'STI.',
                         'KOOPERATIF', 'KOOP', 'TIC.', 'TICARET', 'SANAYI', 'SAN.',
                         'HOLDING', 'GRUP', 'GROUP', 'ITHALAT', 'IHRACAT']

    def ata(satir):
        vkn = sadece_rakam(satir.get('verginumarasi', ''))
        if len(vkn) == 10:
            return 'KRLS'

        tc = sadece_rakam(satir.get('tckimlikno2', ''))
        if len(tc) == 11:
            return 'SHS'

        unvan = str(satir.get('unvan', '')).upper()
        unvan = unvan.replace('İ', 'I').replace('Ş', 'S').replace('Ç', 'C')
        unvan = unvan.replace('Ü', 'U').replace('Ö', 'O').replace('Ğ', 'G')
        for kelime in sirket_kelimeleri:
            # kelime olarak ara: 'AS' -> 'DENIZ AS' evet, 'KASAP' hayir
            if re.search(r'(?<![A-Z])' + re.escape(kelime.strip('.')) + r'(?![A-Z])', unvan):
                return 'KRLS'
        return 'SHS'

    df['cari kayitturu'] = df.apply(ata, axis=1)
    krls_sayisi = (df['cari kayitturu'] == 'KRLS').sum()
    shs_sayisi = (df['cari kayitturu'] == 'SHS').sum()
    print('  cari kayitturu atandi: ' + str(krls_sayisi) + ' KRLS, ' + str(shs_sayisi) + ' SHS')
    return df


def kurallari_excele_yaz(kurallar, yol):
    """KURALLAR sozlugunden (kurallar.py) Ayarlar penceresinin duzenleyebilecegi
    kurallar.xlsx dosyasini olusturur. Yeni parametrik ayarlar da varsayilanlariyla eklenir."""
    liste_alanlari = ['telefon_alanlari', 'buyuk_harf_alanlari', 'kucuk_harf_alanlari',
                      'kod_alanlari', 'kdv_alanlari', 'tarih_alanlari']
    tek_alanlar = ['sayfa', 'girdi_dosya', 'cikti_dosya', 'rapor_dosya']
    bool_alanlari = ['mail_kontrol', 'iban_temizle', 'doviz_standart', 'tc_kimlik_11_hane',
                     'tum_alanlar_zorunlu', 'kist_kontrol', 'bosluk_yerine_nokta',
                     'cari_hesap_120_320']
    mukerrer, eslesme, varsayilan, ayarlar = [], [], [], []
    for modul, m in kurallar.items():
        if modul == 'global':
            continue
        for k in m.get('mukerrer_kontrol', []):
            mukerrer.append({'modul': modul, 'sutunlar': ', '.join(k.get('sutunlar', [])),
                             'aksiyon': k.get('aksiyon', 'isaretle'), 'aciklama': ''})
        for baslik, alan in m.get('eslesme', {}).items():
            eslesme.append({'modul': modul, 'excel_baslik': baslik, 'dia_alan': alan})
        for alan, deger in m.get('varsayilan', {}).items():
            varsayilan.append({'modul': modul, 'alan': alan, 'deger': deger})
        for a in tek_alanlar:
            if a in m:
                ayarlar.append({'modul': modul, 'ayar': a, 'deger': m[a]})
        for a in liste_alanlari:
            if a in m:
                ayarlar.append({'modul': modul, 'ayar': a, 'deger': ', '.join(m[a])})
        for a in bool_alanlari:
            if a in m:
                ayarlar.append({'modul': modul, 'ayar': a, 'deger': 'True' if m[a] else 'False'})
        for a in AYAR_TANIMLARI:
            deger = m.get(a['ayar'], a['varsayilan'])
            if isinstance(deger, bool):
                deger = 'True' if deger else 'False'
            ayarlar.append({'modul': modul, 'ayar': a['ayar'], 'deger': deger,
                            'aciklama': a['aciklama'], 'formul': a['formul'],
                            'ornek': a['ornek']})
    sayfalar = {
        'MUKERRER': pd.DataFrame(mukerrer, columns=['modul', 'sutunlar', 'aksiyon', 'aciklama']),
        'ESLESME': pd.DataFrame(eslesme, columns=['modul', 'excel_baslik', 'dia_alan']),
        'VARSAYILAN': pd.DataFrame(varsayilan, columns=['modul', 'alan', 'deger']),
        'AYARLAR': pd.DataFrame(ayarlar, columns=['modul', 'ayar', 'deger', 'aciklama',
                                                  'formul', 'ornek']),
    }
    with pd.ExcelWriter(yol, engine='openpyxl') as w:
        for ad, df in sayfalar.items():
            df.to_excel(w, sheet_name=ad, index=False)
    print('kurallar.xlsx olusturuldu: ' + yol)
    return yol
