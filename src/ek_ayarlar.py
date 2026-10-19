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
     'aciklama': "Excel formül hataları ve anlamsız değerler (#YOK, #N/A, '-', '.') boş hücreye çevrilir.",
     'formul': 'hücre hata_degerleri listesindeyse -> boş', 'ornek': '#YOK -> (boş)'},
    {'ayar': 'hata_degerleri', 'varsayilan': HATA_DEGERLERI, 'tur': 'metin', 'moduller': TUM,
     'aciklama': 'Boş sayılacak değerler (virgülle). 0 bilerek yok: KDV 0 gibi gerçek değerler silinmesin.',
     'formul': 'değer1, değer2, ...', 'ornek': '#YOK, #N/A, -'},
    {'ayar': 'gorunmez_karakter_temizle', 'varsayilan': True, 'tur': 'bool', 'moduller': TUM,
     'aciklama': ("Sekme, satır sonu, Excel'in _x000D_ kalıntısı, sıfır genişlikli ve bölünmez boşluklar "
                  'temizlenir.'),
     'formul': '\\t \\r \\n _x000D_ NBSP -> bosluk | ZWSP BOM -> silinir', 'ornek': 'ABC_x000D_ LTD -> ABC LTD'},
    {'ayar': 'bozuk_turkce_duzelt', 'varsayilan': True, 'tur': 'bool', 'moduller': TUM,
     'aciklama': 'Yanlış kod sayfasıyla açılmış dosyalardaki bozuk Türkçe harfler düzeltilir.',
     'formul': 'Ã‡->Ç Ä°->İ Åž->Ş Ã¼->ü ... | Ý->İ Þ->Ş Ð->Ğ ý->ı þ->ş ð->ğ',
     'ornek': 'Ä°STANBUL -> İSTANBUL'},
    {'ayar': 'bilimsel_gosterim_kontrol', 'varsayilan': True, 'tur': 'bool', 'moduller': TUM,
     'aciklama': ('Kod/barkod/telefon kolonlarında 1,23457E+11 gibi bilimsel gösterim kurtarılır; hane kaybı varsa '
                  'GEÇERSİZ işaretlenir.'),
     'formul': 'hane sayısı yeterliyse sayıya çevir, değilse uyarı', 'ornek': '8,69E+12 -> GEÇERSİZ (veri kaybı) | 8.691234567890E+12 -> 8691234567890'},
    {'ayar': 'benzer_kayit_kontrol', 'varsayilan': True, 'tur': 'bool',
     'moduller': ['cari', 'stok', 'banka', 'personel', 'demirbas', 'hesapplan'],
     'aciklama': ('Yazımı farklı ama aynı olabilecek kayıtları (ABC LTD ŞTİ / ABC Limited Şirketi) benzerlik '
                  'puanıyla işaretler. Silmez.'),
     'formul': 'LTD/ŞTİ/AŞ/TİC gibi ekler atılır, kalan metinlerin benzerliği >= eşik ise işaretlenir',
     'ornek': "'ABC Gıda Ltd Şti' ~ 'ABC GIDA LİMİTED ŞİRKETİ' %100"},
    {'ayar': 'benzer_kayit_sutunu', 'varsayilan': '', 'tur': 'metin',
     'moduller': ['cari', 'stok', 'banka', 'personel', 'demirbas', 'hesapplan'],
     'varsayilan_modul': {'cari': 'unvan', 'stok': 'stok adı', 'banka': 'Hesap Adı',
                          'personel': 'Adı, Soyadı', 'demirbas': 'DEMİRBAŞ ADI', 'hesapplan': 'Hesap Adı'},
     'aciklama': ('Benzerliği karşılaştırılacak kolon(lar). Birden fazlaysa virgülle; birleştirilerek '
                  'karşılaştırılır.'),
     'formul': 'kolon1, kolon2', 'ornek': 'Adı, Soyadı'},
    {'ayar': 'benzer_kayit_esik', 'varsayilan': '88', 'tur': 'metin',
     'moduller': ['cari', 'stok', 'banka', 'personel', 'demirbas', 'hesapplan'],
     'aciklama': 'Benzerlik eşiği (yüzde). Düşürülürse daha çok, yükseltilirse daha az kayıt işaretlenir.',
     'formul': '0-100', 'ornek': '88'},
    {'ayar': 'degisiklik_raporu', 'varsayilan': True, 'tur': 'bool', 'moduller': TUM,
     'aciklama': ("Her hücre için 'eski değer -> yeni değer' listesi (_DEGISIKLIKLER.xlsx). İkinci sayfada temiz "
                  'veri; değişen hücreler turuncu, eski değer hücre notunda.'),
     'formul': 'okunan veri ile temizlenmiş veri hücre hücre karşılaştırılır',
     'ornek': 'satir 5 | telefon1 | 0532 123 45 67 | 5321234567'},
    {'ayar': 'zorunlu_alanlar', 'varsayilan': '', 'tur': 'metin', 'moduller': TUM,
     'varsayilan_modul': {'cari': 'carikartkodu, unvan', 'stok': 'stok kart kodu, stok adı',
                          'banka': 'Hesap Kodu, Hesap Adı', 'personel': 'Adı, Soyadı, TC Kimlik No',
                          'demirbas': 'DEMİRBAŞ.KODU, DEMİRBAŞ ADI', 'hesapplan': 'Hesap Kodu, Hesap Adı'},
     'aciklama': "DIA'ya gitmeden önce boş olmaması gereken kolonlar. Boşsa GEÇERSİZ işaretlenir.",
     'formul': 'kolon1, kolon2 | kapatmak için YOK', 'ornek': 'carikartkodu, unvan'},
    {'ayar': 'alan_uzunluklari', 'varsayilan': '', 'tur': 'metin', 'moduller': TUM,
     'varsayilan_modul': {'cari': 'carikartkodu=25; unvan=200; vergidairesi=50; eposta=100; adres1=200; adres2=200',
                          'stok': 'stok kart kodu=25; stok adı=200; Barkod No=50',
                          'banka': 'Hesap Kodu=25; Hesap Adı=100',
                          'personel': 'Adı=50; Soyadı=50; Sicil=25',
                          'demirbas': 'DEMİRBAŞ.KODU=25; DEMİRBAŞ ADI=200',
                          'hesapplan': 'Hesap Kodu=50; Hesap Adı=200'},
     'aciklama': ("Kolonların azami karakter sayısı. Uzun olan GEÇERSİZ işaretlenir (DIA'da kesilmesin / hata "
                  'vermesin). Değerler tahminidir, DIA ekranına göre düzeltin.'),
     'formul': 'kolon=uzunluk; kolon=uzunluk', 'ornek': 'carikartkodu=25; unvan=200'},
    {'ayar': 'esleme_hafizasi', 'varsayilan': True, 'tur': 'bool', 'moduller': TUM,
     'aciklama': ("Kurallar'a kaydedilen sütun eşlemeleri program klasöründeki esleme_hafizasi.json'a da yazılır; "
                  'başka müşteride aynı başlık gelirse sormadan uygulanır.'),
     'formul': 'müşterinin kendi eşlemesi > hafıza (müşterininki önceliklidir)', 'ornek': "'Firma Ünvanı' -> unvan"},
    {'ayar': 'capraz_kontrol', 'varsayilan': '', 'tur': 'metin', 'moduller': TUM,
     'varsayilan_modul': {'stok': 'Grup Kodu>urunagaci:' + GRUP_KOLONLARI,
                          'muhkod': 'GRUP KODU>urunagaci:' + GRUP_KOLONLARI},
     'aciklama': ('Bu modüldeki kodlar diğer modülün dosyasında var mı? (aynı müşteri klasörü). Yoksa uyarı verir.'),
     'formul': 'kolon>modul:kolon1,kolon2; ...', 'ornek': 'Grup Kodu>urunagaci:Grup1,GRUP2'},

    # ---------------- Cari ----------------
    {'ayar': 'unvan_standart', 'varsayilan': True, 'tur': 'bool', 'moduller': ['cari'],
     'aciklama': 'Ünvandaki şirket türü kısaltmalarını tek biçime getirir.',
     'formul': 'unvan_kisaltmalari listesine göre', 'ornek': 'ABC Ltd.Şti -> ABC LTD. ŞTİ.'},
    {'ayar': 'unvan_kisaltmalari', 'varsayilan': UNVAN_KISALTMALARI, 'tur': 'metin', 'moduller': ['cari'],
     'aciklama': ('Aranan=Yazılacak çiftleri. Arama Türkçe harf / nokta / büyük-küçük harf duyarsızdır. Sonuna $ '
                  'konan ifade sadece ünvanın sonunda aranır.'),
     'formul': 'ARANAN=YAZILACAK; ...', 'ornek': 'LIMITED SIRKETI=LTD. ŞTİ.'},
    {'ayar': 'vergi_dairesi_duzelt', 'varsayilan': True, 'tur': 'bool', 'moduller': ['cari'],
     'aciklama': ("Vergi dairesinden 'VD', 'V.D.', 'Vergi Dairesi Müdürlüğü' ekleri atılır; il/ilçe adı olanların "
                  'yazımı düzeltilir, büyük harf yazılır.'),
     'formul': "ekleri sil -> il/ilçe listesinde ara -> vergi_dairesi_eslesme'de ara", 'ornek': 'Kadikoy V.D. -> KADIKÖY'},
    {'ayar': 'vergi_dairesi_eslesme', 'varsayilan': '', 'tur': 'metin', 'moduller': ['cari'],
     'aciklama': 'İl/ilçe adı olmayan vergi daireleri için kendi düzeltmeleriniz.',
     'formul': 'ARANAN=YAZILACAK; ...', 'ornek': 'KOZYATAGI=KOZYATAĞI; BUYUK MUKELLEFLER=BÜYÜK MÜKELLEFLER'},
    {'ayar': 'vkn_tc_tutarlilik', 'varsayilan': True, 'tur': 'bool', 'moduller': ['cari'],
     'aciklama': 'Ünvanı şirket (LTD/A.Ş.) olduğu halde sadece TC Kimlik No girilmiş kayıtları işaretler.',
     'formul': 'ünvanda LTD/ŞTİ/AŞ var, VKN boş, TC dolu -> uyarı', 'ornek': 'ABC LTD STI + TC -> UYARI'},
    {'ayar': 'ad_soyad_ayir', 'varsayilan': False, 'tur': 'bool', 'moduller': ['cari'],
     'aciklama': ('Şahıs carilerinde ünvandan ad ve soyad kolonlarını doldurur (e-Fatura). DIA aktarım '
                  'şablonunuzda bu kolonlar varsa açın.'),
     'formul': 'son kelime = soyad, öncekiler = ad', 'ornek': 'Ahmet Ali Yılmaz -> Ahmet Ali | Yılmaz'},
    {'ayar': 'ad_sutunu', 'varsayilan': 'adi', 'tur': 'metin', 'moduller': ['cari'],
     'aciklama': 'Ad kolonunun adı.', 'formul': 'kolon adı', 'ornek': 'adi'},
    {'ayar': 'soyad_sutunu', 'varsayilan': 'soyadi', 'tur': 'metin', 'moduller': ['cari'],
     'aciklama': 'Soyad kolonunun adı.', 'formul': 'kolon adı', 'ornek': 'soyadi'},
    {'ayar': 'adresten_il_ilce', 'varsayilan': True, 'tur': 'bool', 'moduller': ['cari'],
     'aciklama': ("İl/ilçe boşsa adres1/adres2 içinde geçen il ve o ile ait ilçe ile doldurur. 'Ankara Cad.' gibi "
                  'cadde/sokak adları dikkate alınmaz.'),
     'formul': 'adreste il geçiyor ve ilçe o ile aitse -> doldur', 'ornek': "'Moda Cad. 5 Kadıköy/İstanbul' -> İSTANBUL / KADIKÖY"},
    {'ayar': 'adres_kisaltma', 'varsayilan': False, 'tur': 'bool', 'moduller': ['cari'],
     'aciklama': ('Adreslerdeki Mahallesi/Caddesi/Sokak kelimelerini kısaltır (adres uzunluğu sınırına takılmamak '
                  'için).'),
     'formul': 'adres_kisaltmalari listesine göre', 'ornek': 'Moda Mahallesi Bahariye Caddesi -> Moda MAH. Bahariye CAD.'},
    {'ayar': 'adres_kisaltmalari', 'varsayilan': ADRES_KISALTMALARI, 'tur': 'metin', 'moduller': ['cari'],
     'aciklama': 'Aranan=Yazılacak çiftleri.', 'formul': 'ARANAN=YAZILACAK; ...', 'ornek': 'CADDESI=CAD.'},
    {'ayar': 'cep_ayir', 'varsayilan': True, 'tur': 'bool', 'moduller': ['cari'],
     'aciklama': 'Cep kolonu boşsa, telefon kolonlarındaki ilk cep numarası (5xx) cep kolonuna taşınır.',
     'formul': 'telefon 5 ile başlıyor ve cep boşsa -> cep kolonuna taşı', 'ornek': 'telefon1=5321234567 -> ceptel'},
    {'ayar': 'cep_sutunu', 'varsayilan': 'ceptel', 'tur': 'metin', 'moduller': ['cari'],
     'aciklama': 'Cep telefonu kolonunun adı.', 'formul': 'kolon adı', 'ornek': 'ceptel'},
    {'ayar': 'posta_kodu_il_kontrol', 'varsayilan': True, 'tur': 'bool', 'moduller': ['cari', 'personel'],
     'aciklama': 'Posta kodunun ilk 2 hanesi ilin plaka koduyla uyuşmuyorsa uyarı verir.',
     'formul': 'posta kodunun ilk 2 hanesi = ilin plaka kodu', 'ornek': '34710 + ANKARA -> UYARI'},

    # ---------------- Stok ----------------
    {'ayar': 'barkod_kontrol', 'varsayilan': True, 'tur': 'bool', 'moduller': ['stok'],
     'aciklama': ('EAN-8 / UPC-A / EAN-13 / GTIN-14 barkodların kontrol hanesini doğrular. Diğer uzunluklar (iç '
                  'barkod) atlanır.'),
     'formul': 'GTIN mod-10 kontrol hanesi', 'ornek': '8690000000001 -> GEÇERSİZ'},
    {'ayar': 'barkod_sutunu', 'varsayilan': 'Barkod No', 'tur': 'metin', 'moduller': ['stok'],
     'aciklama': 'Barkod kolonunun adı.', 'formul': 'kolon adı', 'ornek': 'Barkod No'},
    {'ayar': 'birim_standart', 'varsayilan': True, 'tur': 'bool', 'moduller': ['stok'],
     'aciklama': 'Birimleri tek biçime getirir; listede olmayan birim için uyarı verir.',
     'formul': 'birim_eslesme listesine göre', 'ornek': 'Adt -> ADET, Kilo -> KG'},
    {'ayar': 'birim_sutunu', 'varsayilan': 'Ana Birim', 'tur': 'metin', 'moduller': ['stok'],
     'aciklama': 'Birim kolon(lar)ı (virgülle).', 'formul': 'kolon1, kolon2', 'ornek': 'Ana Birim'},
    {'ayar': 'birim_eslesme', 'varsayilan': BIRIM_ESLESME, 'tur': 'metin', 'moduller': ['stok'],
     'aciklama': 'Aranan=Yazılacak birim çiftleri (Türkçe harf / nokta / boşluk duyarsız).',
     'formul': 'ARANAN=YAZILACAK; ...', 'ornek': 'ADT=ADET; KILO=KG'},
    {'ayar': 'kdv_eski_oranlar', 'varsayilan': '8, 18', 'tur': 'metin', 'moduller': ['stok'],
     'aciklama': 'Artık kullanılmayan KDV oranları; bu oranlar için uyarı verilir. Kapatmak için YOK.',
     'formul': 'oran1, oran2', 'ornek': '%18 -> UYARI: ESKİ KDV ORANI'},
    {'ayar': 'fiyat_temizle', 'varsayilan': True, 'tur': 'bool', 'moduller': ['stok'],
     'aciklama': ('Metin fiyatları sayıya çevirir, para birimini ayırır (TL dışındaysa ilgili Döviz kolonuna '
                  'yazar).'),
     'formul': "'1.250,50 TL' -> 1250.5 | '12,5 USD' -> 12.5 + Döviz=USD", 'ornek': '1.250,50 TL -> 1250.5'},
    {'ayar': 'alis_satis_kontrol', 'varsayilan': True, 'tur': 'bool', 'moduller': ['stok'],
     'aciklama': 'Alış fiyatı satış fiyatından büyükse uyarı verir.',
     'formul': 'alis_fiyat_sutunu > satis_fiyat_sutunu -> uyarı', 'ornek': 'alış 120, satış 100 -> UYARI'},
    {'ayar': 'satis_fiyat_sutunu', 'varsayilan': 'Fiyat 1', 'tur': 'metin', 'moduller': ['stok'],
     'aciklama': 'Satış fiyatı kolonu.', 'formul': 'kolon adı', 'ornek': 'Fiyat 1'},
    {'ayar': 'alis_fiyat_sutunu', 'varsayilan': 'Fiyat 2', 'tur': 'metin', 'moduller': ['stok'],
     'aciklama': 'Alış fiyatı kolonu.', 'formul': 'kolon adı', 'ornek': 'Fiyat 2'},
    {'ayar': 'stok_adi_birim_duzelt', 'varsayilan': True, 'tur': 'bool', 'moduller': ['stok'],
     'aciklama': 'Stok adındaki ölçü yazımını düzeltir: sayı ile birim arasına boşluk, birim büyük harf.',
     'formul': '500gr / 500 Gr -> 500 GR | 1l -> 1 LT', 'ornek': 'Süt 1l -> Süt 1 LT'},
    {'ayar': 'stok_adi_yazim', 'varsayilan': 'DOKUNMA', 'tur': 'secim', 'moduller': ['stok'],
     'secenekler': ['DOKUNMA', 'BUYUK', 'BASLIK'],
     'aciklama': 'Stok adının harf yazımı.', 'formul': 'DOKUNMA | BUYUK = SÜT | BASLIK = Süt',
     'ornek': 'süt tam yağlı -> SÜT TAM YAĞLI'},

    # ---------------- Banka / Personel ----------------
    {'ayar': 'iban_banka_adi', 'varsayilan': True, 'tur': 'bool', 'moduller': ['banka', 'personel'],
     'aciklama': ("IBAN'daki banka kodundan banka adını doldurur; dolu banka adı IBAN'la uyuşmuyorsa uyarı verir."),
     'formul': "IBAN'ın 5-9. haneleri (banka kodu) -> banka adı", 'ornek': 'TR..00062.. -> GARANTİ BBVA'},
    {'ayar': 'banka_adi_standart', 'varsayilan': True, 'tur': 'bool', 'moduller': ['banka', 'personel'],
     'aciklama': 'Banka adını tek biçime getirir.', 'formul': 'bilinen bankalar listesi',
     'ornek': 'T.C. Ziraat Bankası A.Ş. -> T.C. ZİRAAT BANKASI'},
    {'ayar': 'iban_sutunu', 'varsayilan': 'IBAN', 'tur': 'metin', 'moduller': ['banka', 'personel'],
     'varsayilan_modul': {'banka': 'IBAN No', 'personel': 'IBAN'},
     'aciklama': 'IBAN kolonu.', 'formul': 'kolon adı', 'ornek': 'IBAN No'},
    {'ayar': 'banka_adi_sutunu', 'varsayilan': 'Banka Adı', 'tur': 'metin', 'moduller': ['banka', 'personel'],
     'varsayilan_modul': {'banka': 'Banka adı', 'personel': 'Banka Adı'},
     'aciklama': 'Banka adı kolonu.', 'formul': 'kolon adı', 'ornek': 'Banka adı'},
    {'ayar': 'iban_hesap_no_doldur', 'varsayilan': True, 'tur': 'bool', 'moduller': ['banka'],
     'aciklama': "Hesap No boşsa IBAN'ın son 16 hanesinden (baştaki 0'lar atılarak) doldurur.",
     'formul': "TR IBAN'ın son 16 hanesi, baştaki 0'lar atılarak", 'ornek': 'TR12 0006 4000 0011 2345 6789 01 -> 112345678901'},
    {'ayar': 'hesap_no_sutunu', 'varsayilan': 'Hesap NO', 'tur': 'metin', 'moduller': ['banka'],
     'aciklama': 'Hesap numarası kolonu.', 'formul': 'kolon adı', 'ornek': 'Hesap NO'},
    {'ayar': 'doviz_hesap_adi_kontrol', 'varsayilan': True, 'tur': 'bool', 'moduller': ['banka'],
     'aciklama': 'Hesap adında USD/EURO gibi para birimi geçiyor ama Döviz Türü farklıysa uyarı verir.',
     'formul': "'Garanti USD' + Döviz=TL -> UYARI", 'ornek': 'Garanti USD / TL -> UYARI'},
    {'ayar': 'personel_standart', 'varsayilan': True, 'tur': 'bool', 'moduller': ['personel'],
     'aciklama': 'Cinsiyet, medeni hal ve kan grubu yazımını standartlaştırır (kolon varsa).',
     'formul': 'E/Bay/Erkek -> ERKEK | B/Bekar -> BEKAR | a rh pozitif -> A RH+',
     'ornek': 'Bayan -> KADIN, 0 Rh(-) -> 0 RH-'},
    {'ayar': 'yas_min', 'varsayilan': '15', 'tur': 'metin', 'moduller': ['personel'],
     'aciklama': 'İşe başlama tarihindeki yaş bundan küçükse uyarı verir.', 'formul': 'yıl', 'ornek': '15'},
    {'ayar': 'yas_max', 'varsayilan': '80', 'tur': 'metin', 'moduller': ['personel'],
     'aciklama': 'Bugünkü yaş bundan büyükse uyarı verir (doğum tarihi hatalı olabilir).', 'formul': 'yıl', 'ornek': '80'},

    # ---------------- Demirbas ----------------
    {'ayar': 'demirbas_kontrol', 'varsayilan': True, 'tur': 'bool', 'moduller': ['demirbas'],
     'aciklama': ('Alış tarihi gelecekte mi, tutar 0/negatif mi, amortisman oranı süresiyle uyumlu mu (normal: '
                  '100/yıl, azalan: 2 katı, en fazla %50), birikmiş amortisman tutardan büyük mü kontrol eder.'),
     'formul': 'oran ≈ 100 / yıl (normal) veya en fazla 50 olmak üzere 200 / yıl (azalan)', 'ornek': '5 yıl + %25 -> UYARI (normal %20, azalan %40 olmalı)'},
    {'ayar': 'birikmis_amortisman_sutunu', 'varsayilan': 'BİRİKMİŞ AMORTİSMAN', 'tur': 'metin',
     'moduller': ['demirbas'],
     'aciklama': 'Birikmiş amortisman kolonu (yoksa bu kontrol atlanır).', 'formul': 'kolon adı',
     'ornek': 'BİRİKMİŞ AMORTİSMAN'},

    # ---------------- Hesap plani ----------------
    {'ayar': 'tdhp_kontrol', 'varsayilan': True, 'tur': 'bool', 'moduller': ['hesapplan'],
     'aciklama': ('Ana hesap kodunun (ilk 3 hane) Tek Düzen Hesap Planında olup olmadığını kontrol eder. 9xx nazım '
                  'hesaplar serbesttir.'),
     'formul': "hesap kodunun ilk 3 hanesi TDHP'de olmalı", 'ornek': '125.01 -> UYARI'},
    {'ayar': 'ek_ana_hesaplar', 'varsayilan': '', 'tur': 'metin', 'moduller': ['hesapplan'],
     'aciklama': 'TDHP dışında kullandığınız ana hesaplar (virgülle).', 'formul': 'kod1, kod2', 'ornek': '125, 175'},
    {'ayar': 'hesap_ayrac', 'varsayilan': '.', 'tur': 'metin', 'moduller': ['hesapplan'],
     'aciklama': "Hesap kodu seviye ayracı. '-', '_', '/', ',' ve boşluk bu ayraca çevrilir.",
     'formul': '120-01-001 -> 120.01.001', 'ornek': '.'},
    {'ayar': 'hesap_seviye_uzunluklari', 'varsayilan': '', 'tur': 'metin', 'moduller': ['hesapplan'],
     'aciklama': ('Seviye hane sayıları. Doluysa ayraçsız kodlar bölünür ve uzunluğu uymayan seviyeler için uyarı '
                  'verilir.'),
     'formul': '3,2,3 -> 12001001 = 120.01.001', 'ornek': '3,2,3'},
    {'ayar': 'eksik_ust_hesap_ekle', 'varsayilan': False, 'tur': 'bool', 'moduller': ['hesapplan'],
     'aciklama': ("120.01.001 var ama 120.01 yoksa, 120.01 satırı '(OTOMATIK) 120.01' adıyla dosya sonuna eklenir."),
     'formul': 'her kodun üst seviyeleri yoksa ekle', 'ornek': '120.01.001 -> + 120.01, + 120'},

    # ---------------- Urun agaci (grup agaci) ----------------
    {'ayar': 'grup_agaci_tutarlilik', 'varsayilan': False, 'tur': 'bool', 'moduller': ['urunagaci'],
     'aciklama': ('Aynı grup kodunun farklı üst gruplar altında kullanılmasını işaretler. Grup kodlarınız her üst '
                  'grupta 01, 02 diye yeniden başlıyorsa kapalı tutun.'),
     'formul': 'her grup kodunun tek bir üst grubu olmalı', 'ornek': 'GRUP2=10 hem A hem B altında -> UYARI'},
]
