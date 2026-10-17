"""Ek kontrollerin (ek_kontroller.py) parametrik ayarlari.

Her ayar kurallar.xlsx > AYARLAR sayfasinda modul bazinda duzenlenebilir.
  moduller          : ayarin gecerli oldugu moduller (None = hepsi)
  varsayilan_modul  : modul bazinda farkli varsayilan
Liste/eslesme ayarlarinda deger bos birakilirsa varsayilan kullanilir;
tamamen kapatmak icin YOK yazin.
"""

TUM = None

UNVAN_KISALTMALARI = ('LIMITED SIRKETI=LTD. ŞTİ.; LIMITED STI=LTD. ŞTİ.; LTD SIRKETI=LTD. ŞTİ.; '
                      'LTD STI=LTD. ŞTİ.; ANONIM SIRKETI=A.Ş.; ANONIM STI=A.Ş.; A S=A.Ş.; AS$=A.Ş.; '
                      'TICARET=TİC.; TIC=TİC.; SANAYI=SAN.; SAN=SAN.')
ADRES_KISALTMALARI = ('MAHALLESI=MAH.; MAHALLE=MAH.; MH=MAH.; CADDESI=CAD.; CADDE=CAD.; CD=CAD.; '
                      'SOKAGI=SK.; SOKAK=SK.; SOK=SK.; BULVARI=BLV.; BULVAR=BLV.; NUMARA=NO')
BIRIM_ESLESME = ('AD=ADET; ADT=ADET; ADET=ADET; PCS=ADET; PC=ADET; PIECE=ADET; ST=ADET; '
                 'KG=KG; KILO=KG; KILOGRAM=KG; KGM=KG; GR=GR; G=GR; GRAM=GR; '
                 'LT=LT; L=LT; LITRE=LT; LTR=LT; ML=ML; MILILITRE=ML; '
                 'MT=MT; M=MT; METRE=MT; MTR=MT; CM=CM; MM=MM; M2=M2; METREKARE=M2; M3=M3; METREKUP=M3; '
                 'KOLI=KOLİ; KL=KOLİ; PAKET=PAKET; PK=PAKET; PKT=PAKET; KUTU=KUTU; KT=KUTU; '
                 'TAKIM=TAKIM; TK=TAKIM; SET=TAKIM; CIFT=ÇİFT; DUZINE=DÜZİNE; DZ=DÜZİNE; '
                 'TON=TON; TN=TON; RULO=RULO; TOP=TOP; SAAT=SAAT; GUN=GÜN')
HATA_DEGERLERI = ('#YOK, #N/A, #DEĞER!, #VALUE!, #REF!, #BAŞV!, #SAYI/0!, #DIV/0!, #AD?, #NAME?, '
                  '#SAYI!, #NUM!, #NULL!, #BOŞ!, -, --, ., ?, N/A, NULL')
GRUP_KOLONLARI = 'Grup1,GRUP2,GRUP3,GRUP4,GRUP5,GRUP6'

EK_AYAR_TANIMLARI = [
    # ---------------- Tum moduller ----------------
    {'ayar': 'hata_degerleri_temizle', 'varsayilan': True, 'tur': 'bool', 'moduller': TUM,
     'aciklama': "Excel formul hatalari ve anlamsiz degerler (#YOK, #N/A, '-', '.') bos hucreye cevrilir.",
     'formul': 'hucre hata_degerleri listesinde ise -> bos', 'ornek': '#YOK -> (bos)'},
    {'ayar': 'hata_degerleri', 'varsayilan': HATA_DEGERLERI, 'tur': 'metin', 'moduller': TUM,
     'aciklama': 'Bos sayilacak degerler (virgulle). 0 bilerek yok: KDV 0 gibi gercek degerleri silmesin.',
     'formul': 'deger1, deger2, ...', 'ornek': '#YOK, #N/A, -'},
    {'ayar': 'gorunmez_karakter_temizle', 'varsayilan': True, 'tur': 'bool', 'moduller': TUM,
     'aciklama': "Sekme, satir sonu, Excel'in _x000D_ kalintisi, sifir genislikli ve bolunmez bosluklar temizlenir.",
     'formul': '\\t \\r \\n _x000D_ NBSP -> bosluk | ZWSP BOM -> silinir', 'ornek': 'ABC_x000D_ LTD -> ABC LTD'},
    {'ayar': 'bozuk_turkce_duzelt', 'varsayilan': True, 'tur': 'bool', 'moduller': TUM,
     'aciklama': 'Yanlis kod sayfasiyla acilmis dosyalardaki bozuk Turkce harfler duzeltilir.',
     'formul': 'Ã‡->Ç Ä°->İ Åž->Ş Ã¼->ü ... | Ý->İ Þ->Ş Ð->Ğ ý->ı þ->ş ð->ğ',
     'ornek': 'Ä°STANBUL -> İSTANBUL'},
    {'ayar': 'bilimsel_gosterim_kontrol', 'varsayilan': True, 'tur': 'bool', 'moduller': TUM,
     'aciklama': "Kod/barkod/telefon kolonlarinda 1,23457E+11 gibi bilimsel gosterim kurtarilir; "
                 "hane kaybi varsa GECERSIZ isaretlenir.",
     'formul': 'mantis hanesi yeterliyse sayiya cevir, degilse uyari', 'ornek': '8,69E+12 -> GEÇERSİZ (veri kaybi)'},
    {'ayar': 'benzer_kayit_kontrol', 'varsayilan': True, 'tur': 'bool',
     'moduller': ['cari', 'stok', 'banka', 'personel', 'demirbas', 'hesapplan'],
     'aciklama': "Yazimi farkli ama ayni olabilecek kayitlari (ABC LTD STI / ABC Limited Sirketi) "
                 "benzerlik puaniyla isaretler. Silmez.",
     'formul': 'unvan normalize (LTD/STI/AS/TIC atilir) -> benzerlik >= esik',
     'ornek': "'ABC Gıda Ltd Şti' ~ 'ABC GIDA LİMİTED ŞİRKETİ' %100"},
    {'ayar': 'benzer_kayit_sutunu', 'varsayilan': '', 'tur': 'metin',
     'moduller': ['cari', 'stok', 'banka', 'personel', 'demirbas', 'hesapplan'],
     'varsayilan_modul': {'cari': 'unvan', 'stok': 'stok adı', 'banka': 'Hesap Adı',
                          'personel': 'Adı, Soyadı', 'demirbas': 'DEMİRBAŞ ADI', 'hesapplan': 'Hesap Adı'},
     'aciklama': 'Benzerligi karsilastirilacak kolon(lar). Birden fazlaysa virgulle; birlestirilerek karsilastirilir.',
     'formul': 'kolon1, kolon2', 'ornek': 'Adı, Soyadı'},
    {'ayar': 'benzer_kayit_esik', 'varsayilan': '88', 'tur': 'metin',
     'moduller': ['cari', 'stok', 'banka', 'personel', 'demirbas', 'hesapplan'],
     'aciklama': 'Benzerlik esigi (yuzde). Dusurulurse daha cok, yukseltilirse daha az kayit isaretlenir.',
     'formul': '0-100', 'ornek': '88'},
    {'ayar': 'degisiklik_raporu', 'varsayilan': True, 'tur': 'bool', 'moduller': TUM,
     'aciklama': "Her hucre icin 'eski deger -> yeni deger' listesi (_DEGISIKLIKLER.xlsx). Ikinci sayfada "
                 "temiz veri, degisen hucreler turuncu ve eski deger not olarak.",
     'formul': 'okunan veri ile temizlenmis veri hucre hucre karsilastirilir',
     'ornek': 'satir 5 | telefon1 | 0532 123 45 67 | 5321234567'},
    {'ayar': 'zorunlu_alanlar', 'varsayilan': '', 'tur': 'metin', 'moduller': TUM,
     'varsayilan_modul': {'cari': 'carikartkodu, unvan', 'stok': 'stok kart kodu, stok adı',
                          'banka': 'Hesap Kodu, Hesap Adı', 'personel': 'Adı, Soyadı, TC Kimlik No',
                          'demirbas': 'DEMİRBAŞ.KODU, DEMİRBAŞ ADI', 'hesapplan': 'Hesap Kodu, Hesap Adı'},
     'aciklama': "DIA'ya gitmeden once bos olmamasi gereken kolonlar. Bos ise GECERSIZ isaretlenir.",
     'formul': 'kolon1, kolon2 | kapatmak icin YOK', 'ornek': 'carikartkodu, unvan'},
    {'ayar': 'alan_uzunluklari', 'varsayilan': '', 'tur': 'metin', 'moduller': TUM,
     'varsayilan_modul': {'cari': 'carikartkodu=25; unvan=200; vergidairesi=50; eposta=100; adres1=200; adres2=200',
                          'stok': 'stok kart kodu=25; stok adı=200; Barkod No=50',
                          'banka': 'Hesap Kodu=25; Hesap Adı=100',
                          'personel': 'Adı=50; Soyadı=50; Sicil=25',
                          'demirbas': 'DEMİRBAŞ.KODU=25; DEMİRBAŞ ADI=200',
                          'hesapplan': 'Hesap Kodu=50; Hesap Adı=200'},
     'aciklama': "Kolonlarin azami karakter sayisi. Uzun olan GECERSIZ isaretlenir (DIA'da kesilmesin/hata vermesin). "
                 "Degerler tahmindir, DIA ekranina gore duzeltin.",
     'formul': 'kolon=uzunluk; kolon=uzunluk', 'ornek': 'carikartkodu=25; unvan=200'},
    {'ayar': 'esleme_hafizasi', 'varsayilan': True, 'tur': 'bool', 'moduller': TUM,
     'aciklama': "Kurallar'a kaydedilen sutun eslemeleri program klasorundeki esleme_hafizasi.json'a da yazilir; "
                 "baska musteride ayni baslik gelirse sormadan uygulanir.",
     'formul': "musteri eslemesi > hafiza (musterininki onceliklidir)", 'ornek': "'Firma Ünvanı' -> unvan"},
    {'ayar': 'capraz_kontrol', 'varsayilan': '', 'tur': 'metin', 'moduller': TUM,
     'varsayilan_modul': {'stok': 'Grup Kodu>urunagaci:' + GRUP_KOLONLARI,
                          'muhkod': 'GRUP KODU>urunagaci:' + GRUP_KOLONLARI},
     'aciklama': 'Bu moduldeki kodlar diger modulun dosyasinda var mi? (ayni musteri klasoru). Yoksa uyari.',
     'formul': 'kolon>modul:kolon1,kolon2; ...', 'ornek': 'Grup Kodu>urunagaci:Grup1,GRUP2'},

    # ---------------- Cari ----------------
    {'ayar': 'unvan_standart', 'varsayilan': True, 'tur': 'bool', 'moduller': ['cari'],
     'aciklama': 'Unvandaki sirket turu kisaltmalarini tek bicime getirir.',
     'formul': 'unvan_kisaltmalari listesine gore', 'ornek': 'ABC Ltd.Şti -> ABC LTD. ŞTİ.'},
    {'ayar': 'unvan_kisaltmalari', 'varsayilan': UNVAN_KISALTMALARI, 'tur': 'metin', 'moduller': ['cari'],
     'aciklama': "Aranan=Yazilacak ciftleri. Arama Turkce harf/nokta/buyuk-kucuk harf duyarsizdir. "
                 "Sonuna $ konan ifade sadece unvanin sonunda aranir.",
     'formul': 'ARANAN=YAZILACAK; ...', 'ornek': 'LIMITED SIRKETI=LTD. ŞTİ.'},
    {'ayar': 'vergi_dairesi_duzelt', 'varsayilan': True, 'tur': 'bool', 'moduller': ['cari'],
     'aciklama': "Vergi dairesinden 'VD', 'V.D.', 'Vergi Dairesi Müdürlüğü' ekleri atilir; il/ilce adi olanlarin "
                 "yazimi duzeltilir, buyuk harf yazilir.",
     'formul': 'ek sil -> il/ilce listesi -> vergi_dairesi_eslesme', 'ornek': 'Kadikoy V.D. -> KADIKÖY'},
    {'ayar': 'vergi_dairesi_eslesme', 'varsayilan': '', 'tur': 'metin', 'moduller': ['cari'],
     'aciklama': 'Il/ilce adi olmayan vergi daireleri icin kendi duzeltmeleriniz.',
     'formul': 'ARANAN=YAZILACAK; ...', 'ornek': 'KOZYATAGI=KOZYATAĞI; BUYUK MUKELLEFLER=BÜYÜK MÜKELLEFLER'},
    {'ayar': 'vkn_tc_tutarlilik', 'varsayilan': True, 'tur': 'bool', 'moduller': ['cari'],
     'aciklama': 'Unvani sirket (LTD/A.S.) oldugu halde sadece TC Kimlik No girilmis kayitlari isaretler.',
     'formul': 'unvanda LTD/STI/AS ve VKN bos, TC dolu -> uyari', 'ornek': 'ABC LTD STI + TC -> UYARI'},
    {'ayar': 'ad_soyad_ayir', 'varsayilan': False, 'tur': 'bool', 'moduller': ['cari'],
     'aciklama': "Sahis carilerinde unvandan ad ve soyad kolonlarini doldurur (e-Fatura). DIA aktarim "
                 "sablonunuzda bu kolonlar varsa acin.",
     'formul': 'son kelime = soyad, oncekiler = ad', 'ornek': 'Ahmet Ali Yılmaz -> Ahmet Ali | Yılmaz'},
    {'ayar': 'ad_sutunu', 'varsayilan': 'adi', 'tur': 'metin', 'moduller': ['cari'],
     'aciklama': 'Ad kolonunun adi.', 'formul': 'kolon adi', 'ornek': 'adi'},
    {'ayar': 'soyad_sutunu', 'varsayilan': 'soyadi', 'tur': 'metin', 'moduller': ['cari'],
     'aciklama': 'Soyad kolonunun adi.', 'formul': 'kolon adi', 'ornek': 'soyadi'},
    {'ayar': 'adresten_il_ilce', 'varsayilan': True, 'tur': 'bool', 'moduller': ['cari'],
     'aciklama': "Il/ilce bos ise adres1/adres2 icinde gecen il ve o ile ait ilce ile doldurur. "
                 "'Ankara Cad.' gibi cadde/sokak adlari dikkate alinmaz.",
     'formul': "adreste il var ve ilce o ilde -> doldur", 'ornek': "'Moda Cad. 5 Kadıköy/İstanbul' -> İSTANBUL / KADIKÖY"},
    {'ayar': 'adres_kisaltma', 'varsayilan': False, 'tur': 'bool', 'moduller': ['cari'],
     'aciklama': 'Adreslerdeki Mahallesi/Caddesi/Sokak kelimelerini kisaltir (adres uzunlugu sinirina takilmamak icin).',
     'formul': 'adres_kisaltmalari listesine gore', 'ornek': 'Moda Mahallesi Bahariye Caddesi -> Moda MAH. Bahariye CAD.'},
    {'ayar': 'adres_kisaltmalari', 'varsayilan': ADRES_KISALTMALARI, 'tur': 'metin', 'moduller': ['cari'],
     'aciklama': 'Aranan=Yazilacak ciftleri.', 'formul': 'ARANAN=YAZILACAK; ...', 'ornek': 'CADDESI=CAD.'},
    {'ayar': 'cep_ayir', 'varsayilan': True, 'tur': 'bool', 'moduller': ['cari'],
     'aciklama': 'Cep kolonu bossa, telefon kolonlarindaki ilk cep numarasi (5xx) cep kolonuna tasinir.',
     'formul': 'telefon 5 ile basliyorsa ve cep bos -> cep', 'ornek': 'telefon1=5321234567 -> ceptel'},
    {'ayar': 'cep_sutunu', 'varsayilan': 'ceptel', 'tur': 'metin', 'moduller': ['cari'],
     'aciklama': 'Cep telefonu kolonunun adi.', 'formul': 'kolon adi', 'ornek': 'ceptel'},
    {'ayar': 'posta_kodu_il_kontrol', 'varsayilan': True, 'tur': 'bool', 'moduller': ['cari', 'personel'],
     'aciklama': 'Posta kodunun ilk 2 hanesi ilin plaka koduyla uyusmuyorsa uyari verir.',
     'formul': 'postakodu[:2] == plaka(il)', 'ornek': '34710 + ANKARA -> UYARI'},

    # ---------------- Stok ----------------
    {'ayar': 'barkod_kontrol', 'varsayilan': True, 'tur': 'bool', 'moduller': ['stok'],
     'aciklama': 'EAN-8 / UPC-A / EAN-13 / GTIN-14 barkodlarin kontrol hanesini dogrular. Diger uzunluklar (ic barkod) atlanir.',
     'formul': 'GTIN mod-10 kontrol hanesi', 'ornek': '8690000000001 -> GEÇERSİZ'},
    {'ayar': 'barkod_sutunu', 'varsayilan': 'Barkod No', 'tur': 'metin', 'moduller': ['stok'],
     'aciklama': 'Barkod kolonunun adi.', 'formul': 'kolon adi', 'ornek': 'Barkod No'},
    {'ayar': 'birim_standart', 'varsayilan': True, 'tur': 'bool', 'moduller': ['stok'],
     'aciklama': 'Birimleri tek bicime getirir; listede olmayan birim icin uyari verir.',
     'formul': 'birim_eslesme listesine gore', 'ornek': 'Adt -> ADET, Kilo -> KG'},
    {'ayar': 'birim_sutunu', 'varsayilan': 'Ana Birim', 'tur': 'metin', 'moduller': ['stok'],
     'aciklama': 'Birim kolon(lar)i (virgulle).', 'formul': 'kolon1, kolon2', 'ornek': 'Ana Birim'},
    {'ayar': 'birim_eslesme', 'varsayilan': BIRIM_ESLESME, 'tur': 'metin', 'moduller': ['stok'],
     'aciklama': 'Aranan=Yazilacak birim ciftleri (Turkce harf/nokta/bosluk duyarsiz).',
     'formul': 'ARANAN=YAZILACAK; ...', 'ornek': 'ADT=ADET; KILO=KG'},
    {'ayar': 'kdv_eski_oranlar', 'varsayilan': '8, 18', 'tur': 'metin', 'moduller': ['stok'],
     'aciklama': "Artik kullanilmayan KDV oranlari; bu oranlar icin uyari verilir. Kapatmak icin YOK.",
     'formul': 'oran1, oran2', 'ornek': '%18 -> UYARI: ESKİ KDV ORANI'},
    {'ayar': 'fiyat_temizle', 'varsayilan': True, 'tur': 'bool', 'moduller': ['stok'],
     'aciklama': "Metin fiyatlari sayiya cevirir, para birimini ayirir (TL disindaysa ilgili Doviz kolonuna yazar).",
     'formul': "'1.250,50 TL' -> 1250.5 | '12,5 USD' -> 12.5 + Döviz=USD", 'ornek': '1.250,50 TL -> 1250.5'},
    {'ayar': 'alis_satis_kontrol', 'varsayilan': True, 'tur': 'bool', 'moduller': ['stok'],
     'aciklama': 'Alis fiyati satis fiyatindan buyukse uyari verir.',
     'formul': 'alis_fiyat_sutunu > satis_fiyat_sutunu -> uyari', 'ornek': 'alis 120, satis 100 -> UYARI'},
    {'ayar': 'satis_fiyat_sutunu', 'varsayilan': 'Fiyat 1', 'tur': 'metin', 'moduller': ['stok'],
     'aciklama': 'Satis fiyati kolonu.', 'formul': 'kolon adi', 'ornek': 'Fiyat 1'},
    {'ayar': 'alis_fiyat_sutunu', 'varsayilan': 'Fiyat 2', 'tur': 'metin', 'moduller': ['stok'],
     'aciklama': 'Alis fiyati kolonu.', 'formul': 'kolon adi', 'ornek': 'Fiyat 2'},
    {'ayar': 'stok_adi_birim_duzelt', 'varsayilan': True, 'tur': 'bool', 'moduller': ['stok'],
     'aciklama': 'Stok adindaki olcu yazimini duzeltir: sayi ile birim arasina bosluk, birim buyuk harf.',
     'formul': '500gr / 500 Gr -> 500 GR | 1l -> 1 LT', 'ornek': 'Süt 1l -> Süt 1 LT'},
    {'ayar': 'stok_adi_yazim', 'varsayilan': 'DOKUNMA', 'tur': 'secim', 'moduller': ['stok'],
     'secenekler': ['DOKUNMA', 'BUYUK', 'BASLIK'],
     'aciklama': 'Stok adinin harf yazimi.', 'formul': 'DOKUNMA | BUYUK = SÜT | BASLIK = Süt',
     'ornek': 'süt tam yağlı -> SÜT TAM YAĞLI'},

    # ---------------- Banka / Personel ----------------
    {'ayar': 'iban_banka_adi', 'varsayilan': True, 'tur': 'bool', 'moduller': ['banka', 'personel'],
     'aciklama': "IBAN'daki banka kodundan banka adini doldurur; dolu banka adi IBAN'la uyusmuyorsa uyari verir.",
     'formul': 'IBAN[5:9] banka kodu -> banka adi', 'ornek': 'TR..00062.. -> GARANTİ BBVA'},
    {'ayar': 'banka_adi_standart', 'varsayilan': True, 'tur': 'bool', 'moduller': ['banka', 'personel'],
     'aciklama': 'Banka adini tek bicime getirir.', 'formul': 'bilinen bankalar listesi',
     'ornek': 'T.C. Ziraat Bankası A.Ş. -> T.C. ZİRAAT BANKASI'},
    {'ayar': 'iban_sutunu', 'varsayilan': 'IBAN', 'tur': 'metin', 'moduller': ['banka', 'personel'],
     'varsayilan_modul': {'banka': 'IBAN No', 'personel': 'IBAN'},
     'aciklama': 'IBAN kolonu.', 'formul': 'kolon adi', 'ornek': 'IBAN No'},
    {'ayar': 'banka_adi_sutunu', 'varsayilan': 'Banka Adı', 'tur': 'metin', 'moduller': ['banka', 'personel'],
     'varsayilan_modul': {'banka': 'Banka adı', 'personel': 'Banka Adı'},
     'aciklama': 'Banka adi kolonu.', 'formul': 'kolon adi', 'ornek': 'Banka adı'},
    {'ayar': 'iban_hesap_no_doldur', 'varsayilan': True, 'tur': 'bool', 'moduller': ['banka'],
     'aciklama': "Hesap No bossa IBAN'in son 16 hanesinden (bastaki 0'lar atilarak) doldurur.",
     'formul': 'TR IBAN[10:26].lstrip(0)', 'ornek': 'TR33 0006 1005 1978 6457 8413 26 -> 519786457841326'},
    {'ayar': 'hesap_no_sutunu', 'varsayilan': 'Hesap NO', 'tur': 'metin', 'moduller': ['banka'],
     'aciklama': 'Hesap numarasi kolonu.', 'formul': 'kolon adi', 'ornek': 'Hesap NO'},
    {'ayar': 'doviz_hesap_adi_kontrol', 'varsayilan': True, 'tur': 'bool', 'moduller': ['banka'],
     'aciklama': "Hesap adinda USD/EURO gibi para birimi geciyor ama Doviz Turu farkliysa uyari verir.",
     'formul': "'Garanti USD' + Döviz=TL -> UYARI", 'ornek': 'Garanti USD / TL -> UYARI'},
    {'ayar': 'personel_standart', 'varsayilan': True, 'tur': 'bool', 'moduller': ['personel'],
     'aciklama': 'Cinsiyet, medeni hal ve kan grubu yazimini standartlastirir (kolon varsa).',
     'formul': 'E/Bay/Erkek -> ERKEK | B/Bekar -> BEKAR | a rh pozitif -> A RH+',
     'ornek': 'Bayan -> KADIN, 0 Rh(-) -> 0 RH-'},
    {'ayar': 'yas_min', 'varsayilan': '15', 'tur': 'metin', 'moduller': ['personel'],
     'aciklama': 'Ise baslama tarihindeki yas bundan kucukse uyari.', 'formul': 'yil', 'ornek': '15'},
    {'ayar': 'yas_max', 'varsayilan': '80', 'tur': 'metin', 'moduller': ['personel'],
     'aciklama': 'Bugunku yas bundan buyukse uyari (dogum tarihi hatali olabilir).', 'formul': 'yil', 'ornek': '80'},

    # ---------------- Demirbas ----------------
    {'ayar': 'demirbas_kontrol', 'varsayilan': True, 'tur': 'bool', 'moduller': ['demirbas'],
     'aciklama': "Alis tarihi gelecekte mi, tutar 0/negatif mi, amortisman orani suresiyle uyumlu mu "
                 "(normal: 100/yil, azalan: 2x, en fazla %50), birikmis amortisman tutardan buyuk mu.",
     'formul': 'oran ~ 100/yil veya min(200/yil, 50)', 'ornek': '5 yil + %25 -> UYARI'},
    {'ayar': 'birikmis_amortisman_sutunu', 'varsayilan': 'BİRİKMİŞ AMORTİSMAN', 'tur': 'metin',
     'moduller': ['demirbas'],
     'aciklama': 'Birikmis amortisman kolonu (yoksa bu kontrol atlanir).', 'formul': 'kolon adi',
     'ornek': 'BİRİKMİŞ AMORTİSMAN'},

    # ---------------- Hesap plani ----------------
    {'ayar': 'tdhp_kontrol', 'varsayilan': True, 'tur': 'bool', 'moduller': ['hesapplan'],
     'aciklama': 'Ana hesap kodunun (ilk 3 hane) Tek Duzen Hesap Planinda olup olmadigini kontrol eder. 9xx nazim serbesttir.',
     'formul': 'kod[:3] in TDHP', 'ornek': '125.01 -> UYARI'},
    {'ayar': 'ek_ana_hesaplar', 'varsayilan': '', 'tur': 'metin', 'moduller': ['hesapplan'],
     'aciklama': 'TDHP disinda kullandiginiz ana hesaplar (virgulle).', 'formul': 'kod1, kod2', 'ornek': '125, 175'},
    {'ayar': 'hesap_ayrac', 'varsayilan': '.', 'tur': 'metin', 'moduller': ['hesapplan'],
     'aciklama': "Hesap kodu seviye ayraci. '-', '_', '/', ',' ve bosluk bu ayraca cevrilir.",
     'formul': '120-01-001 -> 120.01.001', 'ornek': '.'},
    {'ayar': 'hesap_seviye_uzunluklari', 'varsayilan': '', 'tur': 'metin', 'moduller': ['hesapplan'],
     'aciklama': "Seviye hane sayilari. Doluysa ayracsiz kodlar bolunur ve uzunlugu uymayan seviyeler icin uyari verilir.",
     'formul': '3,2,3 -> 12001001 = 120.01.001', 'ornek': '3,2,3'},
    {'ayar': 'eksik_ust_hesap_ekle', 'varsayilan': False, 'tur': 'bool', 'moduller': ['hesapplan'],
     'aciklama': "120.01.001 var ama 120.01 yoksa 120.01 satiri '(OTOMATIK) 120.01' adiyla dosya sonuna eklenir.",
     'formul': 'her kodun ust seviyeleri yoksa ekle', 'ornek': '120.01.001 -> + 120.01, + 120'},

    # ---------------- Urun agaci (grup agaci) ----------------
    {'ayar': 'grup_agaci_tutarlilik', 'varsayilan': False, 'tur': 'bool', 'moduller': ['urunagaci'],
     'aciklama': "Ayni grup kodunun farkli ust gruplar altinda kullanilmasini isaretler. Grup kodlariniz "
                 "her ust grupta 01, 02 diye yeniden basliyorsa kapali tutun.",
     'formul': 'GRUPn kodu -> tek bir ust grup yolu olmali', 'ornek': 'GRUP2=10 hem A hem B altinda -> UYARI'},
]
