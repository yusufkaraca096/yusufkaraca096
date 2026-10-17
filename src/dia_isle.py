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
                    kdv_gecerli_mi, tarih_karsilastir, kod_seviye_kontrol,
                    vkn_isle, tc_isle, ayar, KAYNAK_SATIR)
from il_ilce import il_ilce_duzelt
import ek_kontroller

# Calisma klasoru: once bulunulan klasor (kurallar.xlsx varsa), yoksa script klasoru
_cwd = os.getcwd()
_script = os.path.dirname(os.path.abspath(__file__))
if os.path.exists(os.path.join(_cwd, 'kurallar.xlsx')):
    KLASOR = _cwd
else:
    KLASOR = _script


# Excel'in sayiya cevirip bastaki 0'i sildigi, metin olarak kalmasi gereken kolonlar
METIN_KOLONLAR = ['verginumarasi', 'tckimlikno2', 'TC Kimlik No', 'Hesap NO',
                  'IBAN No', 'Barkod No', 'GTIP Kodu', 'telefon1', 'telefon2',
                  'ceptel', 'fax', 'Telefon 1', 'Telefon 2', 'Fax', 'postakodu']


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
    # df.at ile tek tek yazmak buyuk dosyada yavas: listede biriktirip bir kerede yaz
    uyarilar = ['OK' if pd.isna(v) or v == '' else str(v) for v in df['KONTROL_UYARI'].tolist()]
    for r in rapor:
        idx = r['satir'] - 2
        if 0 <= idx < len(df):
            yeni = r['alan'] + ': ' + r['durum']
            uyarilar[idx] = yeni if uyarilar[idx] == 'OK' else uyarilar[idx] + ' | ' + yeni
    df['KONTROL_UYARI'] = uyarilar
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

    # Hata degerleri (#YOK), gorunmez karakterler, bozuk Turkce: bos satir silinmeden once
    df = ek_kontroller.hucre_on_temizlik(df, modul, modul_adi)

    df = temizlik_uygula(df, modul_adi, temizlik_kurallari,
                         kod_alanlari, sayi_alanlari)

    if global_k.get('bosluk_temizle', True):
        df = bosluk_temizle(df)

    df = sayisal_metin_temizle(df, METIN_KOLONLAR)

    rapor = []

    # VKN: Excel'in sildigi bastaki 0'lari geri ekle + kontrol hanesi
    vkn_sifir = ayar(modul, 'vkn_sifir_tamamla')
    vkn_algoritma = ayar(modul, 'vkn_algoritma_kontrol')
    if 'verginumarasi' in df.columns and (vkn_sifir or vkn_algoritma):
        sonuclar = [vkn_isle(v, vkn_sifir, vkn_algoritma) for v in df['verginumarasi']]
        df['verginumarasi'] = [s[0] for s in sonuclar]
        for i, s in enumerate(sonuclar):
            if s[1] not in ('OK', 'BOŞ'):
                rapor.append({'satir': i + 2, 'alan': 'verginumarasi',
                              'deger': s[0], 'durum': s[1]})

    # TC Kimlik No algoritma kontrolu
    if ayar(modul, 'tc_algoritma_kontrol'):
        for alan in ('tckimlikno2', 'TC Kimlik No'):
            if alan in df.columns:
                sonuclar = [tc_isle(v) for v in df[alan]]
                df[alan] = [s[0] for s in sonuclar]
                for i, s in enumerate(sonuclar):
                    if s[1].startswith('UYARI'):
                        rapor.append({'satir': i + 2, 'alan': alan,
                                      'deger': s[0], 'durum': s[1]})

    if modul_adi == 'cari':
        df = vkn_tc_kontrol(df, 'verginumarasi', 'tckimlikno2')

    if global_k.get('turkce_karakter_kodlarda', True):
        for alan in kod_alanlari:
            if alan in df.columns:
                sonuclar = df[alan].apply(kod_isle)
                df[alan] = [s[0] for s in sonuclar]
                for i, s in enumerate(sonuclar):
                    if s[1] not in ('OK', 'BOŞ'):
                        rapor.append({'satir': i + 2, 'alan': alan,
                                      'deger': str(df[alan].iloc[i]), 'durum': s[1]})

    telefon_format = ayar(modul, 'telefon_format')
    for alan in modul.get('telefon_alanlari', []):
        if alan in df.columns:
            sonuclar = [telefon_isle(v, telefon_format) for v in df[alan]]
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
                df[alan] = [s[0] for s in sonuclar]
                for i, s in enumerate(sonuclar):
                    if s[1] not in ('OK', 'BOŞ', 'BOS'):
                        rapor.append({'satir': i + 2, 'alan': alan,
                                      'deger': str(df[alan].iloc[i]), 'durum': s[1]})
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

    # Ek duzeltmeler: unvan, vergi dairesi, adresten il/ilce, cep, birim, fiyat, banka adi, hesap kodu ...
    df = ek_kontroller.on_islem(df, modul, modul_adi, rapor)

    # Il / ilce: ayni hucreye yazilmis, yer degistirmis veya eksik il/ilce
    il_sutun = ayar(modul, 'il_sutunu')
    ilce_sutun = ayar(modul, 'ilce_sutunu')
    il_ilce_aktif = ayar(modul, 'il_ilce_duzelt') and (il_sutun in df.columns or ilce_sutun in df.columns)
    if il_ilce_aktif:
        df, il_rapor = il_ilce_duzelt(df, il_sutun, ilce_sutun,
                                      il_bos_ise_doldur=ayar(modul, 'il_bos_ise_doldur'),
                                      yazim=ayar(modul, 'il_ilce_yazim'))
        rapor.extend(il_rapor)

    for alan in modul.get('buyuk_harf_alanlari', []):
        if il_ilce_aktif and alan in (il_sutun, ilce_sutun):
            continue  # yazimi il_ilce_yazim ayari belirler
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

    # Ek kontroller (sadece uyari): zorunlu alan, uzunluk, benzer kayit, barkod, TDHP ...
    df = ek_kontroller.son_kontroller(df, modul, modul_adi, rapor)

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

    sonuc_ozet = {'modul': modul_adi, 'satir': 0, 'uyari': 0, 'duzeltilen': 0, 'gecersiz': 0, 'durum': ''}
    girdi_yolu = os.path.join(klasor, modul['girdi_dosya'])
    if not os.path.exists(girdi_yolu):
        print('UYARI: ' + modul['girdi_dosya'] + ' bulunamadi.')
        sonuc_ozet['durum'] = 'DOSYA YOK'
        return sonuc_ozet

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

    eslesme = dict(modul.get('eslesme', {}))
    hafiza_aktif = ayar(modul, 'esleme_hafizasi', modul_adi)
    if hafiza_aktif:
        hafiza = ek_kontroller.esleme_hafizasi_yukle(KLASOR).get(modul_adi, {})
        kullanilan = {b: a for b, a in hafiza.items() if b not in eslesme and b in df.columns and a not in df.columns}
        if kullanilan:
            print('  Esleme hafizasindan: ' + ', '.join(b + ' -> ' + a for b, a in kullanilan.items()))
            eslesme.update(kullanilan)
    df = sutun_esle(df, eslesme)
    if hafiza_aktif and modul.get('eslesme'):
        ek_kontroller.esleme_hafizasi_kaydet(KLASOR, modul_adi, modul['eslesme'])

    # Degisiklik raporu icin satirlarin okunan Excel'deki numarasi (gecici kolon)
    df = df.copy()
    df[KAYNAK_SATIR] = [i + 2 for i in df.index]
    once = df.copy()

    df, rapor = veri_temizle(df, modul, modul_adi, kurallar)
    rapor.extend(ek_kontroller.capraz_kontrol(df, modul, modul_adi, kurallar, klasor))
    # Uyari satir numaralari bu andaki satirlara gore; mukerrer silinmeden once yazilir
    df = kontrol_uyari_ekle(df, rapor)

    if ayar(modul, 'degisiklik_raporu', modul_adi) and mukerrer_mod != 'uygula':
        fark = ek_kontroller.degisiklikler(once, df)
        fark_adi = modul['cikti_dosya'].replace('.xlsx', '_DEGISIKLIKLER.xlsx')
        if dry_run:
            print('  [DRY-RUN] Degisiklik: ' + str(len(fark)) + ' hucre (kaydedilmeyecek)')
        elif not fark.empty:
            try:
                ek_kontroller.degisiklik_raporu_yaz(fark, df, os.path.join(klasor, fark_adi))
                print('  Degisiklik raporu: ' + fark_adi + ' (' + str(len(fark)) + ' degisiklik)')
            except Exception as e:
                print('  UYARI: Degisiklik raporu yazilamadi: ' + str(e))
    df = df.drop(columns=[KAYNAK_SATIR])

    mukerrer_kurallari = modul.get('mukerrer_kontrol', [])
    sonuc_ozet.update(_rapor_sayilari(rapor))

    if mukerrer_kurallari:
        if mukerrer_mod == 'incele':
            sonuc = mukerrer_kontrol_et(df, mukerrer_kurallari, mod='sadece_isaretle')
            df = sonuc[0] if isinstance(sonuc, tuple) else sonuc
            df = sil_onay_sutunu_ekle(df)
            inceleme_adi = cikti_dosya_adi(modul, kurallar, 'inceleme')
            inceleme_yolu = os.path.join(klasor, inceleme_adi)
            if dry_run:
                print('  [DRY-RUN] Inceleme dosyasi kaydedilmeyecek: ' + inceleme_adi)
            else:
                inceleme_raporu(df, inceleme_yolu)
                print('  Inceleme dosyasi kaydedildi: ' + inceleme_adi)
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
            print(">> Excel'de ac, silinecek satirlara SIL_ONAY sutununa 'E' yaz.")
            print('>> Sonra: python dia_isle.py ' + modul_adi + ' --uygula')
            _ozet_rapor_yaz(rapor, modul, kurallar, klasor, dry_run)
            sonuc_ozet.update(satir=len(df), durum='INCELEME BEKLIYOR')
            return sonuc_ozet

        elif mukerrer_mod == 'uygula':
            inceleme_adi = cikti_dosya_adi(modul, kurallar, 'inceleme')
            inceleme_yolu = os.path.join(klasor, inceleme_adi)
            if not os.path.exists(inceleme_yolu):
                inceleme_adi2 = modul['cikti_dosya'].replace('.xlsx', '_INCELEME.xlsx')
                inceleme_yolu = os.path.join(klasor, inceleme_adi2)
            if not os.path.exists(inceleme_yolu):
                print('HATA: Inceleme dosyasi bulunamadi.')
                sonuc_ozet['durum'] = 'INCELEME DOSYASI YOK'
                return sonuc_ozet
            print('Inceleme dosyasi okunuyor: ' + os.path.basename(inceleme_yolu))
            # VKN/telefon gibi kolonlar metin okunur (bastaki 0 kaybolmasin)
            metin_kolonlar = {k: str for k in METIN_KOLONLAR}
            df = pd.read_excel(inceleme_yolu, dtype=metin_kolonlar)
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

    cikti_adi = cikti_dosya_adi(modul, kurallar, 'hazir')
    cikti_yolu = os.path.join(klasor, cikti_adi)
    kaydet(df, cikti_yolu, dry_run)
    sonuc_ozet.update(satir=len(df), durum='HAZIR')

    _ozet_rapor_yaz(rapor, modul, kurallar, klasor, dry_run)

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
    return sonuc_ozet


def _rapor_sayilari(rapor):
    return {'uyari': len(rapor),
            'duzeltilen': sum(1 for r in rapor if 'DÜZELTİLDİ' in r['durum']),
            'gecersiz': sum(1 for r in rapor if 'GEÇERSİZ' in r['durum'] or 'HATALI' in r['durum'])}


def _ozet_rapor_yaz(rapor, modul, kurallar, klasor, dry_run):
    if not rapor:
        return
    ozet_adi = cikti_dosya_adi(modul, kurallar, 'ozet')
    if dry_run:
        print('  [DRY-RUN] Ozet rapor kaydedilmeyecek: ' + ozet_adi)
        return
    with pd.ExcelWriter(os.path.join(klasor, ozet_adi), engine='openpyxl') as writer:
        rapor_ozet(rapor).to_excel(writer, sheet_name='OZET', index=False)
        pd.DataFrame(rapor).to_excel(writer, sheet_name='DETAY', index=False)
    print('  Ozet rapor: ' + ozet_adi)


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
    toplu = []

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
                ozet = modul_isle(m, m_kurallar, mukerrer_mod, dry_run, klasor=klasor) or {'modul': m}
                basarili += 1
            except Exception as e:
                print('HATA (' + ad + ' / ' + m + '): ' + str(e))
                ozet = {'modul': m, 'durum': 'HATA: ' + str(e)}
                hatali += 1
            if ozet.get('durum') != 'DOSYA YOK':
                toplu.append(dict(ozet, musteri=ad))

    print('')
    print('############################################################')
    print('# COKLU ISLEME TAMAMLANDI')
    print('# Basarili: ' + str(basarili))
    print('# Hatali  : ' + str(hatali))
    print('############################################################')
    if toplu and not dry_run:
        kolonlar = ['musteri', 'modul', 'durum', 'satir', 'uyari', 'duzeltilen', 'gecersiz']
        adi = 'TOPLU_RAPOR_' + datetime.now().strftime('%Y-%m-%d_%H%M') + '.xlsx'
        try:
            pd.DataFrame(toplu).reindex(columns=kolonlar).to_excel(os.path.join(musteri_root, adi), index=False)
            print('Toplu rapor: ' + os.path.join(musteri_root, adi))
        except Exception as e:
            print('UYARI: Toplu rapor yazilamadi: ' + str(e))


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
