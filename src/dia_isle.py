import sys
import os
import shutil
import pandas as pd
from datetime import datetime
from _ortak import (bosluk_temizle, birebir_mukerrer_sil, kod_mukerrer_sil, kod_isle,
                    telefon_isle, mail_isle, kdv_normalize, tarih_isle, iban_isle,
                    tc_kimlik_kontrol, doviz_standart, evet_hayir_standart, buyuk_harf,
                    kucuk_harf, rapor_ozet, sutun_esle, mukerrer_kontrol_et,
                    sil_onay_sutunu_ekle, onayli_sil, inceleme_raporu,
                    kurallari_excelden_oku, temizlik_uygula, excel_oku_coklu_sayfa,
                    hiyerarsi_kontrol, sayisal_metin_temizle, vkn_tc_kontrol,
                    mukerrer_grup_raporu, cari_kayitturu_ata, sutun_sirala,
                    kdv_gecerli_mi, tarih_karsilastir, kod_seviye_kontrol)

# Calisma klasoru: once bulunulan klasor (kurallar.xlsx varsa), yoksa script klasoru
_cwd = os.getcwd()
_script = os.path.dirname(os.path.abspath(__file__))
if os.path.exists(os.path.join(_cwd, 'kurallar.xlsx')):
    KLASOR = _cwd
else:
    KLASOR = _script


def kurallari_yukle():
    excel_yolu = os.path.join(KLASOR, 'kurallar.xlsx')
    if os.path.exists(excel_yolu):
        print("Kurallar Excel'den okunuyor: kurallar.xlsx")
        return kurallari_excelden_oku(excel_yolu)
    print("Kurallar Python'dan okunuyor: kurallar.py")
    from kurallar import KURALLAR
    return KURALLAR


def cikti_dosya_adi(modul, kurallar, tip='hazir'):
    """Tarihli veya normal dosya adi uretir."""
    global_k = kurallar.get('global', {})
    temel = modul['cikti_dosya'].replace('.xlsx', '')
    if tip == 'inceleme':
        temel = temel + '_INCELEME'
    elif tip == 'silinenler':
        temel = temel + '_SILINENLER'
    elif tip == 'ozet':
        temel = temel.replace('_dia_hazir', '') + '_OZET'
    if global_k.get('tarihli_dosya_adi', False):
        tarih = datetime.now().strftime('%Y-%m-%d_%H%M')
        temel = temel + '_' + tarih
    return temel + '.xlsx'


def yedekle(dosya_yolu):
    """Eski dosya varsa .bak olarak yedekler."""
    if os.path.exists(dosya_yolu):
        yedek = dosya_yolu + '.bak'
        try:
            shutil.copy2(dosya_yolu, yedek)
            print('  Yedek: ' + os.path.basename(yedek))
        except:
            pass


def kaydet(df, dosya_yolu, dry_run=False):
    if dry_run:
        print('  [DRY-RUN] Kaydedilecek: ' + os.path.basename(dosya_yolu) + ' (' + str(len(df)) + ' satir)')
        return
    yedekle(dosya_yolu)
    df.to_excel(dosya_yolu, index=False)
    print('  Kaydedildi: ' + os.path.basename(dosya_yolu) + ' (' + str(len(df)) + ' satir)')
    try:
        from _ortak import excel_renklendir
        excel_renklendir(dosya_yolu)
    except Exception as e:
        print('  UYARI: Renklendirme: ' + str(e))
    # XLS kopyasi
    try:
        from _ortak import xls_olarak_kaydet
        xls_olarak_kaydet(df, dosya_yolu)
    except Exception as e:
        print('  UYARI: XLS olusturulamadi: ' + str(e))


def kontrol_uyari_ekle(df, rapor):
    if not rapor:
        return df
    df = df.reset_index(drop=True)
    if 'KONTROL_UYARI' not in df.columns:
        df['KONTROL_UYARI'] = 'OK'
    for r in rapor:
        idx = r['satir'] - 2
        if 0 <= idx < len(df):
            mevcut = df.at[idx, 'KONTROL_UYARI']
            if mevcut == 'OK':
                df.at[idx, 'KONTROL_UYARI'] = r['alan'] + ': ' + r['durum']
            else:
                df.at[idx, 'KONTROL_UYARI'] += ' | ' + r['alan'] + ': ' + r['durum']
    return df


def veri_temizle(df, modul, modul_adi, kurallar):
    global_k = kurallar.get('global', {})

    TEMIZLIK_MODUL = {
        'cari': ['kod_bas_son_temizle', 'posta_kodu_5_hane', 'sutun_sirasi'],
        'stok': ['kod_bas_son_temizle', 'sayi_format_duzelt', 'sutun_sirasi'],
        'banka': ['kod_bas_son_temizle', 'sayi_format_duzelt', 'sutun_sirasi'],
        'personel': ['posta_kodu_5_hane', 'sayi_format_duzelt'],
        'demirbas': ['kod_bas_son_temizle', 'sayi_format_duzelt'],
        'hesapplan': ['kod_bas_son_temizle'],
        'muhkod': ['kod_bas_son_temizle'],
        'urunagaci': ['kod_bas_son_temizle'],
    }
    temizlik_kurallari = [
        {'kural': 'bos_satir_temizle', 'aktif': True},
        {'kural': 'coklu_bosluk', 'aktif': True},
        {'kural': 'bos_sutun_sil', 'aktif': True},
        {'kural': 'toplam_satirlarini_temizle', 'aktif': True},
    ]
    for k in TEMIZLIK_MODUL.get(modul_adi, []):
        temizlik_kurallari.append({'kural': k, 'aktif': True})

    kod_alanlari = modul.get('kod_alanlari', [])
    sayi_map = {
        'stok': ['Fiyat 1', 'Fiyat 2', 'Fiyat 3', 'Fiyat 4', 'Fiyat 5'],
        'personel': ['Maaş'],
        'demirbas': ['TUTAR', 'MİKTAR'],
    }
    sayi_alanlari = sayi_map.get(modul_adi, [])

    df = temizlik_uygula(df, modul_adi, temizlik_kurallari,
                         kod_alanlari, sayi_alanlari)

    if global_k.get('bosluk_temizle', True):
        df = bosluk_temizle(df)

    sayisal_metin = ['verginumarasi', 'tckimlikno2', 'TC Kimlik No', 'Hesap NO',
                     'IBAN No', 'Barkod No', 'GTIP Kodu', 'telefon1', 'telefon2',
                     'ceptel', 'fax']
    df = sayisal_metin_temizle(df, sayisal_metin)
    if modul_adi == 'cari':
        df = vkn_tc_kontrol(df, 'verginumarasi', 'tckimlikno2')

    rapor = []

    if global_k.get('turkce_karakter_kodlarda', True):
        for alan in kod_alanlari:
            if alan in df.columns:
                sonuclar = df[alan].apply(kod_isle)
                df[alan] = [s[0] for s in sonuclar]
                for i, s in enumerate(sonuclar):
                    if s[1] not in ('OK', 'BOŞ'):
                        rapor.append({'satir': i + 2, 'alan': alan,
                                      'deger': str(df[alan].iloc[i]), 'durum': s[1]})

    for alan in modul.get('telefon_alanlari', []):
        if alan in df.columns:
            sonuclar = df[alan].apply(telefon_isle)
            df[alan] = [s[0] for s in sonuclar]
            for i, s in enumerate(sonuclar):
                if s[1] not in ('OK', 'BOŞ'):
                    rapor.append({'satir': i + 2, 'alan': alan,
                                  'deger': str(df[alan].iloc[i]), 'durum': s[1]})

    if modul.get('mail_kontrol', False) and 'eposta' in df.columns:
        sonuclar = df['eposta'].apply(mail_isle)
        df['eposta'] = [s[0] for s in sonuclar]
        for i, s in enumerate(sonuclar):
            if s[1] not in ('OK', 'BOŞ'):
                rapor.append({'satir': i + 2, 'alan': 'eposta',
                              'deger': str(df['eposta'].iloc[i]), 'durum': s[1]})

    for alan in modul.get('kdv_alanlari', []):
        if alan in df.columns:
            df[alan] = df[alan].apply(kdv_normalize)

    # KDV oran gecerlilik kontrolu (stok)
    if modul_adi == 'stok' and modul.get('kdv_alanlari'):
        if 'KONTROL_UYARI' not in df.columns:
            df['KONTROL_UYARI'] = 'OK'
        for alan in modul.get('kdv_alanlari', []):
            if alan not in df.columns:
                continue
            for i, v in enumerate(df[alan]):
                if v is None or str(v).strip() == '':
                    continue
                if not kdv_gecerli_mi(v):
                    mevcut = df.at[i, 'KONTROL_UYARI']
                    yeni = alan + ': GECERSIZ KDV (' + str(v) + ')'
                    if mevcut == 'OK' or mevcut == '':
                        df.at[i, 'KONTROL_UYARI'] = yeni
                    elif yeni not in mevcut:
                        df.at[i, 'KONTROL_UYARI'] = mevcut + ' | ' + yeni

    for alan in modul.get('tarih_alanlari', []):
        if alan in df.columns:
            sonuclar = df[alan].apply(tarih_isle)
            df[alan] = [s[0] for s in sonuclar]
            for i, s in enumerate(sonuclar):
                if s[1] not in ('OK', 'BOŞ'):
                    rapor.append({'satir': i + 2, 'alan': alan,
                                  'deger': str(df[alan].iloc[i]), 'durum': s[1]})

    if modul.get('iban_temizle', False):
        for alan in ('IBAN No', 'IBAN'):
            if alan in df.columns:
                sonuclar = df[alan].apply(iban_isle)
                df['IBAN No'] = [s[0] for s in sonuclar]
                for i, s in enumerate(sonuclar):
                    if s[1] not in ('OK', 'BOŞ', 'BOS'):
                        rapor.append({'satir': i + 2, 'alan': 'IBAN No',
                                      'deger': str(df['IBAN No'].iloc[i]), 'durum': s[1]})
                break

    if modul.get('doviz_standart', False) and 'Döviz Türü' in df.columns:
        df['Döviz Türü'] = df['Döviz Türü'].apply(doviz_standart)

    if 'Kredi Hesabımı?' in df.columns:
        df['Kredi Hesabımı?'] = df['Kredi Hesabımı?'].apply(evet_hayir_standart)

    # Personel tarih sirasi
    if modul_adi == 'personel':
        df = tarih_karsilastir(df, ['Doğum Tarihi', 'İşe Başlama Tarihi', 'İşten Ayrılma Tarihi'])

    if modul.get('tc_kimlik_11_hane', False):
        for alan in ('TC Kimlik No', 'tckimlikno2'):
            if alan in df.columns:
                sonuclar = df[alan].apply(tc_kimlik_kontrol)
                df[alan] = [s[0] for s in sonuclar]
                for i, s in enumerate(sonuclar):
                    if s[1] not in ('OK', 'BOŞ'):
                        rapor.append({'satir': i + 2, 'alan': alan,
                                      'deger': str(df[alan].iloc[i]), 'durum': s[1]})

    for alan in modul.get('buyuk_harf_alanlari', []):
        if alan in df.columns:
            df[alan] = df[alan].apply(buyuk_harf)

    for alan in modul.get('kucuk_harf_alanlari', []):
        if alan in df.columns:
            df[alan] = df[alan].apply(kucuk_harf)

    for alan, deger in modul.get('varsayilan', {}).items():
        if alan not in df.columns:
            df[alan] = deger
        else:
            df[alan] = df[alan].fillna(deger)

    # Urun agaci grup seviye kontrolu
    if modul_adi == 'urunagaci':
        df = kod_seviye_kontrol(df, ['Grup1', 'GRUP2', 'GRUP3', 'GRUP4', 'GRUP5', 'GRUP6'], ad='Grup')

    # Muhasebe kodu ozel kod seviye kontrolu
    if modul_adi == 'muhkod':
        df = kod_seviye_kontrol(df, ['STOK ÖZEL KOD 1', 'STOK ÖZEL KOD 2'], ad='Ozel Kod')

    if modul_adi == 'hesapplan':
        hesap_kodu_sutunu = None
        for aday in ('Hesap Kodu', 'hesap kodu', 'Kodu'):
            if aday in df.columns:
                hesap_kodu_sutunu = aday
                break
        if hesap_kodu_sutunu:
            df = hiyerarsi_kontrol(df, hesap_kodu_sutunu)

    # Cari kayit turu
    if modul_adi == 'cari':
        df = cari_kayitturu_ata(df)
        if 'TARİH' in df.columns:
            df = df.drop(columns=['TARİH'])
        if 'tarih' in df.columns:
            df = df.drop(columns=['tarih'])

    if any(k.get('kural') == 'sutun_sirasi' and k.get('aktif', True) for k in temizlik_kurallari):
        df = sutun_sirala(df, modul_adi)

    return df, rapor


def modul_isle(modul_adi, kurallar, mukerrer_mod='incele', dry_run=False, klasor=None):
    if klasor is None:
        klasor = KLASOR
    print('')
    print('============================================================')
    print('MODUL: ' + modul_adi.upper() + '  (mukerrer mod: ' + mukerrer_mod + ')')
    if dry_run:
        print('[DRY-RUN MODU] Dosya yazilmayacak, sadece rapor gosterilecek.')
    print('============================================================')

    modul = kurallar.get(modul_adi)
    if not modul:
        print("HATA: '" + modul_adi + "' modulu tanimli degil.")
        return

    girdi_yolu = os.path.join(klasor, modul['girdi_dosya'])
    if not os.path.exists(girdi_yolu):
        print('UYARI: ' + modul['girdi_dosya'] + ' bulunamadi.')
        return

    print('Dosya okunuyor: ' + modul['girdi_dosya'])
    sayfa_ayari = modul.get('sayfa', '')
    df = excel_oku_coklu_sayfa(girdi_yolu, sayfa_ayari)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.str.startswith('Unnamed')]
    df = df.loc[:, ~df.columns.str.match('^$')]
    print('Okunan satir: ' + str(len(df)))

    if kurallar.get('global', {}).get('birebir_mukerrer_sil', True):
        df = birebir_mukerrer_sil(df)
        print('Birebir mukerrer sonrasi: ' + str(len(df)))

    eslesme = modul.get('eslesme', {})
    df = sutun_esle(df, eslesme)

    df, rapor = veri_temizle(df, modul, modul_adi, kurallar)

    mukerrer_kurallari = modul.get('mukerrer_kontrol', [])

    if mukerrer_kurallari:
        if mukerrer_mod == 'incele':
            sonuc = mukerrer_kontrol_et(df, mukerrer_kurallari, mod='sadece_isaretle')
            df = sonuc[0] if isinstance(sonuc, tuple) else sonuc
            df = sil_onay_sutunu_ekle(df)
            df = kontrol_uyari_ekle(df, rapor)
            inceleme_adi = cikti_dosya_adi(modul, kurallar, 'inceleme')
            inceleme_yolu = os.path.join(klasor, inceleme_adi)
            if dry_run:
                print('  [DRY-RUN] Inceleme dosyasi kaydedilmeyecek: ' + inceleme_adi)
            else:
                inceleme_raporu(df, inceleme_yolu)
            # Mukerrer grup raporu
            try:
                gdf = mukerrer_grup_raporu(df, mukerrer_kurallari)
                if not gdf.empty:
                    gadi = modul['cikti_dosya'].replace('.xlsx', '_MUKERRER_GRUPLAR.xlsx')
                    gyolu = os.path.join(klasor, gadi)
                    gdf.to_excel(gyolu, index=False)
                    print('  Mukerrer grup raporu: ' + gadi + ' (' + str(len(gdf)) + ' grup)')
            except Exception as e:
                print('  UYARI: Grup raporu: ' + str(e))
                print('  Inceleme dosyasi kaydedildi: ' + inceleme_adi)
            print(">> Excel'de ac, silinecek satirlara SIL_ONAY sutununa 'E' yaz.")
            print('>> Sonra: python dia_isle.py ' + modul_adi + ' --uygula')
            return

        elif mukerrer_mod == 'uygula':
            inceleme_adi = cikti_dosya_adi(modul, kurallar, 'inceleme')
            inceleme_yolu = os.path.join(klasor, inceleme_adi)
            if not os.path.exists(inceleme_yolu):
                inceleme_adi2 = modul['cikti_dosya'].replace('.xlsx', '_INCELEME.xlsx')
                inceleme_yolu = os.path.join(klasor, inceleme_adi2)
            if not os.path.exists(inceleme_yolu):
                print('HATA: Inceleme dosyasi bulunamadi.')
                return
            print('Inceleme dosyasi okunuyor: ' + os.path.basename(inceleme_yolu))
            df = pd.read_excel(inceleme_yolu)
            df, silinen_df = onayli_sil(df)
            if not silinen_df.empty:
                silinen_adi = cikti_dosya_adi(modul, kurallar, 'silinenler')
                silinen_yolu = os.path.join(klasor, silinen_adi)
                kaydet(silinen_df, silinen_yolu, dry_run)

        else:
            sonuc = mukerrer_kontrol_et(df, mukerrer_kurallari, mod=mukerrer_mod)
            df = sonuc[0] if isinstance(sonuc, tuple) else sonuc
            silinen_df = sonuc[1] if isinstance(sonuc, tuple) else pd.DataFrame()
            print('Mukerrer kontrol sonrasi: ' + str(len(df)) + ' satir')
            if not silinen_df.empty:
                silinen_adi = cikti_dosya_adi(modul, kurallar, 'silinenler')
                silinen_yolu = os.path.join(klasor, silinen_adi)
                kaydet(silinen_df, silinen_yolu, dry_run)

    df = kontrol_uyari_ekle(df, rapor)

    cikti_adi = cikti_dosya_adi(modul, kurallar, 'hazir')
    cikti_yolu = os.path.join(klasor, cikti_adi)
    kaydet(df, cikti_yolu, dry_run)

    if rapor:
        rapor_df = pd.DataFrame(rapor)
        ozet_df = rapor_ozet(rapor)
        ozet_adi = cikti_dosya_adi(modul, kurallar, 'ozet')
        ozet_yolu = os.path.join(klasor, ozet_adi)
        if dry_run:
            print('  [DRY-RUN] Ozet rapor kaydedilmeyecek: ' + ozet_adi)
        else:
            with pd.ExcelWriter(ozet_yolu, engine='openpyxl') as writer:
                ozet_df.to_excel(writer, sheet_name='OZET', index=False)
                rapor_df.to_excel(writer, sheet_name='DETAY', index=False)
            print('  Ozet rapor: ' + ozet_adi)

    if dry_run:
        print('')
        print('============================================================')
        print('[DRY-RUN OZET] ' + modul_adi.upper())
        print('  Toplam satir: ' + str(len(df)))
        print('  Uyari sayisi: ' + str(len(rapor)))
        if rapor:
            duzeltilen = sum(1 for r in rapor if 'DÜZELTİLDİ' in r['durum'])
            hatali = sum(1 for r in rapor if 'HATALI' in r['durum'] or 'GEÇERSİZ' in r['durum'])
            print('  Duzeltilen: ' + str(duzeltilen))
            print('  Gecersiz/Hatali: ' + str(hatali))
        print('============================================================')


def musteri_hepsi_isle(modul, kurallar, mukerrer_mod, dry_run):
    """musteriler/ altindaki her klasoru sirayla isle."""
    adaylar = [
        os.path.join(KLASOR, 'musteriler'),
        os.path.join(os.path.dirname(KLASOR), 'musteriler'),
        os.path.join(os.getcwd(), 'musteriler'),
    ]
    musteri_root = None
    for a in adaylar:
        if os.path.exists(a) and os.path.isdir(a):
            musteri_root = a
            break
    if not musteri_root:
        print("HATA: 'musteriler' klasoru bulunamadi.")
        for a in adaylar:
            print('  Arandi: ' + a)
        return

    klasorler = []
    for ad in sorted(os.listdir(musteri_root)):
        tam = os.path.join(musteri_root, ad)
        if os.path.isdir(tam):
            klasorler.append((ad, tam))

    if not klasorler:
        print("UYARI: 'musteriler' altinda klasor yok.")
        return

    print('')
    print('############################################################')
    print('# COKLU MUSTERI ISLEME')
    print('# Konum  : ' + musteri_root)
    print('# Musteri: ' + str(len(klasorler)))
    print('# Modul  : ' + modul)
    print('############################################################')

    moduller = ['cari', 'stok', 'banka', 'personel', 'demirbas', 'hesapplan', 'muhkod', 'urunagaci']

    islenecek = moduller if modul == 'hepsi' else [modul]

    basarili = 0
    hatali = 0

    for ad, klasor in klasorler:
        print('')
        print('############################################################')
        print('# MUSTERI: ' + ad)
        print('# KLASOR : ' + klasor)
        print('############################################################')

        m_kurallar_dosya = os.path.join(klasor, 'kurallar.xlsx')
        if os.path.exists(m_kurallar_dosya):
            try:
                m_kurallar = kurallari_excelden_oku(m_kurallar_dosya)
            except Exception as e:
                print('UYARI: kurallar.xlsx okunamadi: ' + str(e))
                m_kurallar = kurallar
        else:
            m_kurallar = kurallar
            kaynak = os.path.join(KLASOR, 'kurallar.xlsx')
            if os.path.exists(kaynak):
                try:
                    import shutil as _sh
                    _sh.copy2(kaynak, m_kurallar_dosya)
                    print('  kurallar.xlsx kopyalandi: ' + m_kurallar_dosya)
                except Exception as e:
                    print('  UYARI: kurallar.xlsx kopyalanamadi: ' + str(e))

        for m in islenecek:
            try:
                modul_isle(m, m_kurallar, mukerrer_mod, dry_run, klasor=klasor)
                basarili += 1
            except Exception as e:
                print('HATA (' + ad + ' / ' + m + '): ' + str(e))
                hatali += 1

    print('')
    print('############################################################')
    print('# COKLU ISLEME TAMAMLANDI')
    print('# Basarili: ' + str(basarili))
    print('# Hatali  : ' + str(hatali))
    print('############################################################')


def arg_parse(argv):
    modul = None
    mukerrer_mod = 'incele'
    dry_run = False
    for arg in argv:
        if arg.startswith('--mod='):
            mukerrer_mod = arg.split('=', 1)[1]
        elif arg == '--uygula':
            mukerrer_mod = 'uygula'
        elif arg == '--sil':
            mukerrer_mod = 'otomatik'
        elif arg == '--incele':
            mukerrer_mod = 'incele'
        elif arg == '--dry-run':
            dry_run = True
        elif not arg.startswith('--'):
            modul = arg.lower()
    return modul, mukerrer_mod, dry_run


def main():
    if len(sys.argv) < 2:
        print('Kullanim: python dia_isle.py [cari|stok|...|hepsi]')
        print('Opsiyonlar: --incele, --uygula, --sil, --dry-run')
        print('Coklu musteri: python dia_isle.py hepsi --musteri-hepsi')
        return
    modul, mukerrer_mod, dry_run = arg_parse(sys.argv[1:])
    if not modul:
        print('HATA: Modul adi verilmedi.')
        return
    kurallar = kurallari_yukle()
    moduller = ['cari', 'stok', 'banka', 'personel', 'demirbas', 'hesapplan', 'muhkod', 'urunagaci']

    if '--musteri-hepsi' in sys.argv:
        musteri_hepsi_isle(modul, kurallar, mukerrer_mod, dry_run)
        return
    if modul == 'hepsi':
        for m in moduller:
            modul_isle(m, kurallar, mukerrer_mod, dry_run)
    elif modul in moduller:
        modul_isle(modul, kurallar, mukerrer_mod, dry_run)
    else:
        print("HATA: '" + modul + "' bilinmeyen modul.")


if __name__ == '__main__':
    main()
