"""Kullanim rehberi.

Tek kaynak: program icindeki Rehber penceresi (gui.RehberPenceresi), HTML disa aktarimi ve
docs/KULLANIM_REHBERI.md ayni icerikten uretilir. Ayar referansi AYAR_TANIMLARI'ndan
otomatik olusur; yeni ayar eklenince rehber kendiliginden guncellenir.

Bolum: {'baslik': str, 'parcalar': [(tur, icerik), ...]}
  ('p', metin)                 paragraf
  ('h', metin)                 alt baslik
  ('madde', [metin, ...])      madde isaretli liste
  ('adim', [metin, ...])       numarali adimlar
  ('kod', metin)               aynen yazilacak deger / ornek (tek yazi tipi)
  ('tablo', [basliklar, [satir, ...]])
  ('not', metin)               dikkat kutusu
"""
import html

from _ortak import AYAR_TANIMLARI, modul_ayar_tanimlari

MODULLER = [('cari', 'Cari'), ('stok', 'Stok'), ('banka', 'Banka'), ('personel', 'Personel'),
            ('demirbas', 'Demirbaş'), ('hesapplan', 'Hesap Planı'), ('muhkod', 'Muhasebe Kodu'),
            ('urunagaci', 'Ürün Ağacı')]

# AYAR_TANIMLARI disindaki (eski) AYARLAR satirlari
DOSYA_AYARLARI = [
    ['girdi_dosya', 'Okunacak müşteri dosyası', 'cari_listesi.xlsx'],
    ['cikti_dosya', 'DIA\'ya yüklenecek temiz dosya', 'cari_dia_hazir.xlsx'],
    ['sayfa', 'Okunacak sayfa(lar). Boşsa ilk sayfa; birden fazlaysa virgülle, birleştirilir', 'Sayfa1, Sayfa2'],
    ['telefon_alanlari', 'Telefon olarak işlenecek kolonlar', 'telefon1, telefon2, ceptel, fax'],
    ['buyuk_harf_alanlari', 'BÜYÜK harfe çevrilecek kolonlar (Türkçe i/İ doğru çevrilir)', 'sehir, ilce'],
    ['kucuk_harf_alanlari', 'küçük harfe çevrilecek kolonlar', 'eposta'],
    ['kod_alanlari', 'Kod kolonları: Türkçe harf → ASCII, boşluk → nokta, büyük harf', 'carikartkodu, grupkodu'],
    ['kdv_alanlari', 'KDV oranı kolonları (%18, 0,18, 18 → 18)', 'KDV Satış Oran'],
    ['tarih_alanlari', 'Tarih kolonları (gg.aa.yyyy yazılır)', 'İşe Başlama Tarihi'],
    ['mail_kontrol', 'E-posta biçim kontrolü ve küçük harf', 'True'],
    ['iban_temizle', 'IBAN boşluk temizliği ve MOD-97 kontrolü', 'True'],
    ['doviz_standart', 'Döviz adları: Dolar → USD, Avro → EUR, TRY → TL', 'True'],
    ['tc_kimlik_11_hane', 'TC Kimlik No 11 hane kontrolü', 'True'],
]


def _deger_yazi(d):
    if isinstance(d, bool):
        return 'True' if d else 'False'
    return '(boş)' if d in ('', None) else str(d)


def _tur_yazi(a):
    tur = a.get('tur', 'metin')
    if tur == 'bool':
        return 'Açık/Kapalı (True/False)'
    if tur == 'secim':
        return 'Seçim: ' + ' / '.join(a.get('secenekler', []))
    return 'Metin'


def _ayar_satirlari(tanimlar, modul_adi=None):
    parcalar = []
    for a in tanimlar:
        vm = a.get('varsayilan_modul')
        if vm and modul_adi is None:
            varsayilan = '; '.join(m + ': ' + _deger_yazi(v) for m, v in vm.items()) + \
                '  (diğer modüller: ' + _deger_yazi(a['varsayilan']) + ')'
        else:
            varsayilan = _deger_yazi(a['varsayilan'])
        parcalar.append(('h', a['ayar']))
        parcalar.append(('p', a['aciklama']))
        parcalar.append(('tablo', [['', ''], [
            ['Tür', _tur_yazi(a)],
            ['Varsayılan', varsayilan],
            ['Formül / yazım', a.get('formul', '')],
            ['Örnek', a.get('ornek', '')],
        ]]))
    return parcalar


def bolumler():
    B = []

    # ------------------------------------------------------------------
    B.append({'baslik': '1. Başlarken', 'parcalar': [
        ('p', 'Program, müşteriden gelen Excel listelerini (cari, stok, banka, personel, demirbaş, hesap planı, '
              'muhasebe kodu, ürün ağacı) temizleyip DIA\'ya yüklenebilir hale getirir. Hiçbir zaman müşterinin '
              'orijinal dosyasını değiştirmez; sonuçları yeni dosyalara yazar, eskisini .bak olarak yedekler.'),
        ('h', 'Klasör düzeni'),
        ('p', 'Her müşteri için ayrı bir klasör açın. Müşterinin dosyalarını şu adlarla koyun '
              '(adlar Ayarlar > girdi_dosya ile değiştirilebilir):'),
        ('tablo', [['Modül', 'Girdi dosyası', 'Çıktı (DIA\'ya yüklenecek)'], [
            [ad, kod + '_listesi.xlsx', kod + '_dia_hazir.xlsx'] for kod, ad in MODULLER]]),
        ('p', 'Birden fazla müşteriyi tek seferde işlemek için program klasörünün yanında "musteriler" klasörü '
              'açıp her müşteriyi alt klasör olarak koyun ve modülde HEPSİ seçin.'),
        ('h', 'Çalıştırma adımları'),
        ('adim', [
            'Klasör Seç ile müşteri klasörünü seçin.',
            'Modülü seçin (Cari, Stok ... veya HEPSİ).',
            'İşlem modunu seçin (aşağıdaki tablo).',
            '▶ ÇALIŞTIR\'a basın. Alttaki pencerede ne yapıldığı satır satır yazar.',
            'Çıkan dosyaları Klasörü Aç ile açıp kontrol edin.',
        ]),
        ('tablo', [['İşlem modu', 'Ne yapar', 'Ne zaman'], [
            ['İnceleme', 'Hiçbir satırı silmez. _INCELEME.xlsx üretir; mükerrer satırlar sarı, SIL_ONAY kolonu yeşil.',
             'İlk çalıştırma; müşteriye sormak istediğiniz satırlar varsa.'],
            ['Uygula', '_INCELEME.xlsx\'te SIL_ONAY kolonuna E yazdığınız satırları siler, _dia_hazir.xlsx yazar.',
             'İnceleme dosyasını düzenledikten sonra.'],
            ['Otomatik Sil', 'Mükerrer sekmesinde aksiyonu "sil" olan kurallara uyan satırları kendisi siler.',
             'Kurallara güveniyorsanız, hızlı sonuç için.'],
            ['Kuru Çalıştır', 'Hiç dosya yazmaz, sadece ne yapılacağını raporlar.', 'Ayar denerken.'],
        ]]),
        ('not', 'İnceleme modunda _dia_hazir.xlsx oluşmaz. DIA\'ya gönderim için önce Uygula veya Otomatik Sil '
                'çalıştırılmalıdır.'),
    ]})

    # ------------------------------------------------------------------
    B.append({'baslik': '2. Ayarlar Penceresi', 'parcalar': [
        ('p', 'Ana ekranda ⚙ Ayarlar. Ayarlar müşteri klasöründeki kurallar.xlsx dosyasına yazılır; her müşterinin '
              'ayarları ayrıdır. Klasörde kurallar.xlsx yoksa program klasöründeki kopyalanır, o da yoksa '
              'varsayılanlardan oluşturulur.'),
        ('adim', [
            'Üstteki MODÜL kutusundan modülü seçin (ayarlar modül bazındadır).',
            'Sekmeyi seçin: Mükerrer, Eşleşme, Varsayılan veya Ayarlar.',
            'Satırı seçip ✎ Düzenle (veya + Yeni / Sil).',
            'Bitince KAYDET. KAYDET\'e basmadan pencereyi kapatırsanız değişiklikler kaybolur.',
        ]),
        ('h', 'Sekmeler'),
        ('tablo', [['Sekme', 'Ne içindir', 'Örnek satır'], [
            ['Mükerrer', 'Hangi kolon(lar) aynıysa kayıt mükerrer sayılır, ne yapılır (isaretle / sil).',
             'sutunlar: vergidairesi, verginumarasi | aksiyon: sil'],
            ['Eşleşme', 'Müşterinin kolon başlığı → DIA kolon adı.', 'excel_baslik: Firma Ünvanı | dia_alan: unvan'],
            ['Varsayılan', 'Kolon yoksa eklenir, boş hücreler bu değerle dolar.', 'alan: doviz | deger: TL'],
            ['Ayarlar', 'Temizlik ve kontrollerin açık/kapalı ve parametreleri.', 'telefon_format | 10'],
        ]]),
        ('h', 'Ayarlar sekmesi butonları'),
        ('madde', [
            'Satıra çift tık veya ℹ Detay: ayarın açıklaması, formülü ve örneği büyük pencerede.',
            '＋ Eksik Ayarları Ekle: eski bir kurallar.xlsx\'te olmayan yeni ayarları varsayılan değerleriyle ekler. '
            'Program güncellenince bir kez basıp KAYDET deyin.',
            'Ayarlar sekmesinde satırı olmayan ayar için de varsayılan değer kullanılır; satırı silmek ayarı '
            'kapatmaz, varsayılana döndürür. Kapatmak için değeri False (liste ayarlarında YOK) yapın.',
        ]),
    ]})

    # ------------------------------------------------------------------
    B.append({'baslik': '3. Değer Yazım Kuralları (Formül Dili)', 'parcalar': [
        ('p', 'Ayarlar sekmesindeki "deger" kolonuna yazılan değerler aşağıdaki biçimlerden birindedir. '
              'Her ayarın biçimi Ayar Referansı bölümünde "Formül / yazım" satırında yazar.'),
        ('tablo', [['Biçim', 'Nasıl yazılır', 'Örnek'], [
            ['Açık/Kapalı', 'True veya False. Evet, E, 1, X da açık sayılır.', 'True'],
            ['Seçim', 'Seçeneklerden biri, aynen.', 'BUYUK'],
            ['Sayı', 'Sadece rakam.', '88'],
            ['Kolon adı', 'DIA kolonunun adı, büyük/küçük harf ve Türkçe karakter AYNEN.', 'stok adı'],
            ['Liste', 'Virgülle ayrılmış değerler.', 'carikartkodu, unvan, verginumarasi'],
            ['Eşleme çiftleri', 'ARANAN=YAZILACAK, çiftler noktalı virgülle.', 'ADT=ADET; KILO=KG'],
            ['Uzunluk', 'kolon=karakter sayısı, noktalı virgülle.', 'carikartkodu=25; unvan=200'],
            ['Çapraz kontrol', 'kolon>modül:kolon1,kolon2 ; birden fazla kural noktalı virgülle.',
             'Grup Kodu>urunagaci:Grup1,GRUP2'],
        ]]),
        ('h', 'Eşleme çiftlerinde arama nasıl yapılır'),
        ('madde', [
            'Büyük/küçük harf fark etmez: "adt", "ADT", "Adt" hepsi ADT=… ile eşleşir.',
            'Türkçe karakter fark etmez: "KILO" yazmanız "Kilo", "KİLO", "kılo"yu da yakalar. Aranan tarafı '
            'Türkçe karaktersiz yazmanız yeterli; yazılacak tarafı istediğiniz gibi (ÇİFT, KOLİ) yazın.',
            'Ünvan ve adres kısaltmalarında nokta ve boşluk esnektir: "LTD STI" aranırsa "Ltd.Şti", "LTD. ŞTİ.", '
            '"Ltd Şti" hepsi bulunur. Sadece tam kelime bulunur: "SAN" aranırken "SANAT" değişmez.',
            'Birim eşlemesinde hücrenin tamamı karşılaştırılır (nokta ve boşluk yok sayılır): "A D T" = ADT.',
            'Aranan ifadenin sonuna $ koyarsanız sadece metnin SONUNDA aranır: "AS$=A.Ş." → "DENİZ AS" değişir, '
            '"AS YAPI" değişmez.',
            'Uzun ifadeler önce uygulanır: "LIMITED SIRKETI" ve "SIRKETI" ikisi de varsa önce uzun olan çalışır.',
        ]),
        ('h', 'Boş ve YOK'),
        ('madde', [
            'Değer boş bırakılırsa varsayılan kullanılır.',
            'Liste ve eşleme ayarlarını tamamen kapatmak için YOK yazın (örn. zorunlu_alanlar = YOK).',
            'Açık/Kapalı ayarlar için False yazın.',
        ]),
        ('not', 'Kolon adları birebir eşleşmelidir. "Stok Adı" ile "stok adı" farklı kolonlardır. Doğru adı '
                'görmek için _dia_hazir.xlsx\'in başlık satırına bakın.'),
    ]})

    # ------------------------------------------------------------------
    tarifler = [
        ('Listede olmayan birimi tanıtmak (TANIMSIZ BİRİM uyarısı)', [
            'Ayarlar > MODÜL: stok > Ayarlar sekmesi.',
            'birim_eslesme satırını seçip ✎ Düzenle.',
            'Değerin SONUNA noktalı virgülle yeni çiftleri ekleyin (mevcutları silmeyin):',
        ], '…; TOP=TOP; CUVAL=ÇUVAL; CVL=ÇUVAL; BIDON=BİDON',
            'KAYDET, sonra modülü yeniden ÇALIŞTIR. Satır yoksa önce ＋ Eksik Ayarları Ekle.'),
        ('Yeni ünvan kısaltması eklemek', [
            'MODÜL: cari > Ayarlar > unvan_kisaltmalari > Düzenle.',
            'Sona ekleyin:',
        ], '…; INSAAT=İNŞ.; PAZARLAMA=PAZ.; ITHALAT IHRACAT=İTH. İHR.',
            'Bir kısaltmayı istemiyorsanız (örn. TİCARET açık yazılsın) o çifti silin: "TICARET=TİC.; TIC=TİC." '
            'kısmını kaldırın. Hiç kısaltma istemiyorsanız unvan_standart = False.'),
        ('İl/ilçe adı olmayan vergi dairelerini düzeltmek', [
            'MODÜL: cari > vergi_dairesi_eslesme > Düzenle (varsayılan boştur).',
            'Müşterinin yazımı = doğru yazım çiftlerini girin:',
        ], 'KOZYATAGI=KOZYATAĞI; BUYUK MUKELLEFLER=BÜYÜK MÜKELLEFLER; ULUS=ULUS; GOKSU=GÖKSU',
            'Aranan tarafta Türkçe karakter ve boşluk/nokta önemli değildir. "V.D.", "Vergi Dairesi Müdürlüğü" '
            'ekleri zaten silinir, eklemenize gerek yok.'),
        ('Zorunlu alan eklemek veya kaldırmak', [
            'İlgili modül > zorunlu_alanlar > Düzenle.',
            'Kolon adlarını virgülle yazın:',
        ], 'carikartkodu, unvan, verginumarasi, sehir',
            'Boş olan satırlar KONTROL_UYARI\'da "GEÇERSİZ: ZORUNLU ALAN BOŞ" olarak kırmızı görünür. '
            'Kontrolü kapatmak için YOK yazın.'),
        ('Alan uzunluğunu DIA\'ya göre ayarlamak', [
            'DIA ekranında alanın kaç karakter aldığına bakın.',
            'İlgili modül > alan_uzunluklari > Düzenle:',
        ], 'carikartkodu=20; unvan=250; adres1=150; vergidairesi=40',
            'Daha uzun değerler "GEÇERSİZ: UZUNLUK 130 > 100 KARAKTER" uyarısı alır (değer kesilmez).'),
        ('Telefon biçimini değiştirmek', [
            'İlgili modül > telefon_format > Düzenle.',
            'Seçeneklerden birini yazın:',
        ], '10   → 5321234567\n11   → 05321234567\n90   → 905321234567\n+90  → +905321234567',
            'Hangi kolonların telefon sayılacağı telefon_alanlari ayarındadır.'),
        ('Cep numarasının ceptel\'e taşınmasını kapatmak', [
            'MODÜL: cari > cep_ayir > Düzenle > False.',
        ], 'False', 'Başka bir cep kolonu kullanıyorsanız cep_sutunu ayarına o kolonun adını yazın.'),
        ('Benzer kayıt (olası mükerrer) hassasiyeti', [
            'İlgili modül > benzer_kayit_esik > Düzenle.',
        ], '92', 'Yükseltirseniz (95) sadece çok benzerler, düşürürseniz (80) daha fazla kayıt işaretlenir. '
                 'Karşılaştırılan kolonları benzer_kayit_sutunu belirler; örn. personelde "Adı, Soyadı".'),
        ('Yeni mükerrer kuralı eklemek', [
            'İlgili modül > Mükerrer sekmesi > + Yeni.',
            'sutunlar: virgülle kolonlar; aksiyon: isaretle veya sil.',
        ], 'sutunlar: unvan, sehir\naksiyon : isaretle',
            '"sil" aksiyonu sadece 2 veya daha fazla kolonlu kurallarda ve Otomatik Sil modunda uygulanır; '
            'tek kolonlu kurallar sadece işaretler (yanlışlıkla toplu silme olmasın diye).'),
        ('Müşterinin kolon başlığını tanıtmak', [
            'İlgili modül > Eşleşme sekmesi > + Yeni.',
            'excel_baslik: müşterinin dosyasındaki başlık (aynen); dia_alan: DIA kolonu.',
        ], 'excel_baslik: Firma Ünvanı\ndia_alan    : unvan',
            'Daha kolayı: ana ekranda Eşleme Öğren dosyayı tarar, tanımadığı başlıkları listeler; açılır listeden '
            'seçip kaydedin. esleme_hafizasi açıksa bu eşlemeler sonraki müşterilerde de kendiliğinden uygulanır.'),
        ('Boş hücrelere sabit değer yazmak', [
            'İlgili modül > Varsayılan sekmesi > + Yeni.',
        ], 'alan : doviz\ndeger: TL', 'Kolon yoksa eklenir ve tüm satırlara yazılır; varsa sadece boş hücreler dolar.'),
        ('TDHP dışı ana hesap kullanan müşteri', [
            'MODÜL: hesapplan > ek_ana_hesaplar > Düzenle.',
        ], '125, 175, 850', 'Bu hesaplar artık "TEK DÜZEN HESAP PLANINDA YOK" uyarısı almaz.'),
        ('Ayraçsız hesap kodlarını bölmek', [
            'MODÜL: hesapplan > hesap_seviye_uzunluklari > Düzenle.',
        ], '3,2,3', '"12001001" → "120.01.001". Ayrıca 120.1 gibi seviyesi kısa kalan kodlar uyarı alır '
                    '(Excel sayı olarak okuyup 120.10\'un sonundaki 0\'ı silmiş olabilir).'),
        ('Modüller arası yeni çapraz kontrol', [
            'Kontrol edilecek modül > capraz_kontrol > Düzenle.',
            'Mevcut kuralın sonuna noktalı virgülle ekleyin:',
        ], 'Grup Kodu>urunagaci:Grup1,GRUP2,GRUP3; Özel Kod 1>muhkod:STOK ÖZEL KOD 1',
            'Hedef modülün dosyası (önce _dia_hazir, yoksa _listesi) aynı müşteri klasöründe olmalıdır.'),
        ('Bir kontrolü tamamen kapatmak', [
            'Ayarın değerini False yapın (liste ayarlarında YOK).',
        ], 'barkod_kontrol = False\nzorunlu_alanlar = YOK', ''),
        ('Ayarları tüm yeni müşterilere varsayılan yapmak', [
            'Bir müşteride ayarları istediğiniz gibi yapıp KAYDET.',
            'O müşterinin klasöründeki kurallar.xlsx\'i program (EXE) klasörüne kopyalayın.',
        ], '', 'Bundan sonra kurallar.xlsx\'i olmayan her yeni müşteri klasörüne bu dosya kopyalanır. '
               'Mevcut müşterilerin dosyaları değişmez.'),
    ]
    p = [('p', 'Sık yapılan işlemler adım adım. Her işlemden sonra KAYDET\'e basıp modülü yeniden çalıştırın.')]
    for baslik, adimlar, kod, sonra in tarifler:
        p.append(('h', baslik))
        p.append(('adim', adimlar))
        if kod:
            p.append(('kod', kod))
        if sonra:
            p.append(('p', sonra))
    B.append({'baslik': '4. Örneklerle Ayar Yapma (Tarifler)', 'parcalar': p})

    # ------------------------------------------------------------------
    B.append({'baslik': '5. Raporları Okumak', 'parcalar': [
        ('tablo', [['Dosya', 'İçerik'], [
            ['X_dia_hazir.xlsx / .xls', 'DIA\'ya yüklenecek temiz veri. KONTROL_UYARI kolonunda satırın uyarıları.'],
            ['X_dia_hazir_INCELEME.xlsx', 'İnceleme modu: SIL_ONAY kolonuna E yazılan satırlar Uygula\'da silinir.'],
            ['X_dia_hazir_DEGISIKLIKLER.xlsx', 'DEGISIKLIKLER: her hücre için eski → yeni. TEMIZ_VERI: değişen '
             'hücreler turuncu, eski değer hücre notunda (fareyle üzerine gelin).'],
            ['X_OZET.xlsx', 'OZET: toplam uyarı/düzeltilen/geçersiz sayısı. DETAY: satır, kolon, değer, durum.'],
            ['X_dia_hazir_SILINENLER.xlsx', 'Silinen mükerrer satırlar (geri almak için).'],
            ['X_dia_hazir_MUKERRER_GRUPLAR.xlsx', 'Mükerrer kayıtlar grup grup.'],
            ['musteriler/TOPLU_RAPOR_tarih.xlsx', 'Tüm müşteriler × modüller: satır, uyarı, düzeltilen, geçersiz.'],
            ['X_DIA_KURU_RAPOR.xlsx / X_DIA_AKTARIM_RAPORU.xlsx', 'DIA gönderiminin satır satır sonucu.'],
            ['*.bak', 'Üzerine yazılan dosyanın bir önceki hali.'],
        ]]),
        ('h', 'KONTROL_UYARI kelimeleri ve renkler'),
        ('tablo', [['Durum', 'Anlamı', 'Renk'], [
            ['DÜZELTİLDİ (eski -> yeni)', 'Program düzeltti, bilgi amaçlı.', 'Turuncu'],
            ['UYARI: …', 'Düzeltilmedi, kontrol edin (olası hata).', '—'],
            ['GEÇERSİZ / HATALI', 'DIA\'ya bu haliyle gitmemeli, düzeltin.', 'Kırmızı'],
            ['MÜKERRER / OLASI MÜKERRER', 'Başka satırla aynı/benzer; satır numarası yazar.', 'Turuncu'],
            ['KONTROL_MUKERRER kolonu', 'Mükerrer kurallarına takılan satırlar.', 'Sarı'],
            ['SIL_ONAY = E', 'Uygula\'da silinecek.', 'Yeşil'],
        ]]),
        ('p', 'Uyarıdaki "satir 5" Excel\'deki satır numarasıdır (1. satır başlık). Excel\'de Veri > Filtre ile '
              'KONTROL_UYARI kolonunu "GEÇERSİZ" içerenlere süzerek önce kırmızıları düzeltin.'),
    ]})

    # ------------------------------------------------------------------
    B.append({'baslik': '6. Logo\'dan Çekme ve DIA\'ya Gönderme', 'parcalar': [
        ('h', 'Logo\'dan Çek'),
        ('adim', [
            'SQL sunucu, veritabanı, kullanıcı/şifre (veya Windows oturumu) ve firma numarasını girin.',
            'Bağlantı Testi: firmalar listelenir.',
            'Cari Çek / Stok Çek: müşteri klasörüne cari_listesi.xlsx / stok_listesi.xlsx yazılır.',
            'Sonra normal temizlik (ÇALIŞTIR). Logo verisine yazılmaz, sadece okunur.',
        ]),
        ('not', 'Bilgisayarda "ODBC Driver 17 (veya 18) for SQL Server" kurulu olmalıdır.'),
        ('h', '☁ DIA\'ya Gönder'),
        ('adim', [
            'Sunucu kodu (örn. racabilisim), kullanıcı, şifre, API key, Firma Kodu, Dönem Kodu.',
            'Bağlantı Testi: kalan kontör ve yetkili firma/dönem listesi gelir. Girdiğiniz firma listede yoksa '
            '"!! DIKKAT" uyarısı çıkar (401 INSUFFICIENT_PRIVILEGES hatasının sebebi budur).',
            'Kuru çalıştır işaretliyken GÖNDER: hiçbir şey yazılmaz, X_DIA_KURU_RAPOR.xlsx üretilir.',
            'Raporu kontrol edip işareti kaldırın ve GÖNDER. Kod DIA\'da yoksa eklenir, varsa güncellenir.',
        ]),
        ('p', 'Her çağrı 0,0125 kontör harcar. Kontör biterse (406) gönderim kendiliğinden durur. Şifreler hiçbir '
              'yere kaydedilmez; API key sadece "hatırla" işaretliyse kaydedilir.'),
        ('h', 'DIA alan eşlemesi (dia_eslesme.json)'),
        ('p', 'DIA penceresindeki Alan Eşlemesi butonu müşteri klasöründeki dia_eslesme.json\'u açar. "kart" '
              'bölümünde sol taraf DIA alanı, sağ taraf Excel kolonudur. Başında ? olan satırlar doğrulanmamıştır '
              've gönderilmez; DIA dokümanından kontrol edip ?\'i silin.'),
        ('kod', '"kart": {\n'
                '    "carikartkodu": "carikartkodu",\n'
                '    "unvan": "unvan",\n'
                '    "eposta": "eposta",            <- ? silindi, artık gönderilir\n'
                '    "?webadresi": "web adresi"     <- gönderilmez\n'
                '},\n'
                '"sabit": {"carikarttipi": "AL"},   <- her karta aynı değer\n'
                '"mevcut_ise": "guncelle"           <- veya "atla"'),
        ('not', 'JSON yazımında virgül ve tırnaklara dikkat edin; son satırdan sonra virgül olmaz. Bozulursa dosyayı '
                'silin, program varsayılanını yeniden oluşturur.'),
    ]})

    # ------------------------------------------------------------------
    B.append({'baslik': '7. Sık Karşılaşılan Sorunlar', 'parcalar': [
        ('tablo', [['Sorun', 'Çözüm'], [
            ['"cari_listesi.xlsx bulunamadi"', 'Dosya adını kontrol edin veya Ayarlar > girdi_dosya\'yı değiştirin.'],
            ['"cari_dia_hazir.xlsx bulunamadi" (DIA)', 'İnceleme modunda hazır dosya oluşmaz; Uygula/Otomatik Sil çalıştırın.'],
            ['Dosya yazılamadı / Permission denied', 'Dosya Excel\'de açık; kapatıp tekrar çalıştırın.'],
            ['VKN/telefonun başındaki 0 gitmiş', 'Program VKN\'yi 10 haneye tamamlar. Müşteriden metin biçimli dosya isteyin.'],
            ['"BİLİMSEL GÖSTERİM, HANE KAYBI"', 'Kaynak Excel\'de kolonu Metin yapıp değerleri yeniden alın; kayıp geri getirilemez.'],
            ['Yeni ayar Ayarlar\'da görünmüyor', '＋ Eksik Ayarları Ekle, sonra KAYDET.'],
            ['Ayar değiştirdim ama etkisi yok', 'KAYDET\'e bastınız mı? Doğru MODÜL seçili mi? Kolon adı aynen mi?'],
            ['Çok fazla OLASI MÜKERRER', 'benzer_kayit_esik\'i yükseltin (95) veya benzer_kayit_kontrol = False.'],
            ['DIA 401 INSUFFICIENT_PRIVILEGES', 'Firma kodu yanlış; Bağlantı Testi\'ndeki yetkili firma listesine bakın.'],
            ['DIA 406', 'Kontör bitmiş; DIA\'dan kontör yükletin.'],
            ['Windows "bilinmeyen yayıncı"', 'Ek bilgi > Yine de çalıştır (imzasız EXE).'],
        ]]),
    ]})

    # ------------------------------------------------------------------
    ortak = [a for a in AYAR_TANIMLARI if not a.get('moduller')]
    B.append({'baslik': '8. Ayar Referansı: Tüm Modüller', 'parcalar': [
        ('p', 'Bu ayarlar her modülde ayrı ayrı bulunur (her modül için farklı değer verilebilir).'),
    ] + _ayar_satirlari(ortak)})
    for kod, ad in MODULLER:
        ozel = [a for a in modul_ayar_tanimlari(kod) if a.get('moduller')]
        if ozel:
            B.append({'baslik': '8. Ayar Referansı: ' + ad, 'parcalar': _ayar_satirlari(ozel, kod)})
    B.append({'baslik': '8. Ayar Referansı: Dosya ve Kolon Ayarları', 'parcalar': [
        ('p', 'Her modülde bulunan temel ayarlar.'),
        ('tablo', [['Ayar', 'Ne yapar', 'Örnek değer'], DOSYA_AYARLARI]),
    ]})

    # ------------------------------------------------------------------
    from teklif_veri import AYARLAR as TEKLIF_AYARLARI
    B.append({'baslik': '9. Teklif Motoru', 'parcalar': [
        ('p', 'DIA lisans + hizmet teklifi hazırlar; RACA_DIA_Teklif_Motoru.xlsx\'in program halidir. Ana ekranda '
              'Teklif Motoru butonuyla veya tek başına RACA_Teklif.exe ile açılır (aynı veri kullanılır).'),
        ('h', 'İlk açılış'),
        ('adim', [
            'Program "paket listesi yok" der; Evet deyip mevcut teklif Excel\'inizi seçin.',
            'PaketListesi, ADAM SAAT, AYARLAR, KAMPANYALAR, RAKİP ve TEKLİF GEÇMİŞİ sayfaları içe aktarılır.',
            'Fiyatlar programın içine değil, sizin bilgisayarınızdaki veri dosyasına yazılır.',
        ]),
        ('h', 'Teklif hazırlama (sekmeler)'),
        ('adim', [
            'Müşteri: ad, kullanıcı sayıları, yetkili / telefon / e-posta / VKN. Paket tipi eşiğe göre seçilir. '
            'Etiket (Ekonomik / Önerilen / Kapsamlı) alternatif tekliflerde kullanılır.',
            '1) Paketler: + Paket Ekle ile arayın, ekleyin. Adet, İnd. %, İnd. Tutar, A/S kolonlarına çift tıklayıp '
            'değiştirin. Sabit kalemler, kampanya, maliyet ve bu teklife özel saat ücreti de burada.',
            'Kodla hızlı ekleme: "Kod:" kutusuna paket kodunu yazıp/yapıştırıp Enter. Birden çok kodu Excel\'den '
            'veya bir listeden kopyalayıp "Panodan Yapıştır" (ya da paket listesinde Ctrl+V) ile tek seferde '
            'ekleyin. Alt alta, virgül/noktalı virgül/boşluk ile ayrılmış kodlar kabul edilir; büyük/küçük harf, '
            'Türkçe karakter ve - _ . farkı önemsizdir. "KOD 3" veya Excel\'de kodun yanındaki hücre tam sayıysa '
            'adet olarak alınır. Bulunamayan, pasif ve teklifte zaten olan kodlar listelenir, eklenmez.',
            'Uyarlama bedeli: ekranın altındaki turuncu bantta (tıklayınca Keşif & Uyarlama dökümü açılır) ve '
            'PDF/Excel\'de ayrı bölümde gösterilir. PDF\'te bölüm başında açıklama metni (Ayarlar: '
            'uyarlama_aciklama), her hizmet kaleminin kapsamı, dayanağı (ör. "7 kullanıcı → 1 grup × 8 saat") ve '
            'hesabı (saat × saat ücreti), süre saat ve iş günü olarak (Ayarlar: is_gunu_saat), sonunda da vurgulu '
            '"UYARLAMA HİZMET BEDELİ" kutusu yer alır. Elle değiştirilen saat ve kalibrasyon gerekçeleri müşteriye '
            'gösterilmez.',
            '2) Keşif & Uyarlama: soruları cevaplayın. Sağdaki hizmet dökümü her kalemin saatini ve "neye göre" '
            'hesaplandığını gösterir. Saat kolonuna çift tıklayıp elle değiştirebilirsiniz (boş = öneri). Müşteri '
            'verisinden doldur: veri temizleme klasöründen cari/stok sayısını ve hata oranını alır. Şehir dışıysa '
            'yerinde gün ve kişi sayısını girin.',
            '3) Ödeme / Taksit: her seçenekte müşterinin ödeyeceği, bankanın keseceği komisyon ve size kalan tutar. '
            'PDF kolonuna çift tıklayarak müşteriye hangi seçeneklerin sunulacağını belirleyin.',
            '4) Bakım Anlaşması: önerilen paket ve hesabın dökümü; teklife eklenecek paketi seçin.',
            '5) Satış Asistanı & Kontroller: tahmini kazanma ihtimali, öneriler ve kritik hatalar.',
            'Kaydet: numara verilir. Revizyon: aynı teklifin -R1, -R2 sürümü (eskisi korunur). PDF / Excel: '
            'kritik hata varsa yönetici şifresi (tanımlıysa) veya onay istenir.',
        ]),
        ('h', 'Uyarlama nasıl hesaplanır (hizmet dökümü)'),
        ('tablo', [['Hesap türü', 'Formül', 'Örnek hizmet'], [
            ['PAKET_SAAT', 'Paketlerin ADAM SAAT tablosundaki saatleri toplamı', 'Kurulum ve parametre'],
            ['KAYIT', 'kayıt sayısı ÷ P1 (kayıt/saat) + P2 (hazırlık saati)', 'Veri aktarımı'],
            ['KULLANICI', 'tavan(kullanıcı ÷ P1 kişi) × P2 saat', 'Kullanıcı eğitimi'],
            ['ADET', 'miktar × P1 saat', 'Özel rapor, entegrasyon'],
            ['SABIT', 'P1 saat', 'Canlıya geçiş desteği'],
            ['YUZDE', 'diğer hizmetlerin P1 %\'si', 'Proje yönetimi'],
        ]]),
        ('p', 'Sonra keşif cevapları uygulanır: EKLE (saat ekler; SAYI sorularında eşiği aşan her birim için), CARP '
              '(katsayı, * = tüm hizmetler; örn. IT personeli yoksa ×1,15), CARP_YUZDE (veri hata oranı %18 ise '
              '×1,18), MIKTAR (kayıt/rapor/entegrasyon sayısı). En son kalibrasyon çarpanı uygulanır. Uyarlama = '
              'toplam saat × saat ücreti + lisans × lisans_yuzdesi; alt/üst sınır uygulanır. Hizmet kalemleri ve '
              'sorular Veriler > Hizmet Kalemleri / Keşif Soruları\'ndan değiştirilir, yeni soru eklenebilir.'),
        ('h', 'Tahminleri gerçekle düzeltmek (kalibrasyon)'),
        ('adim', [
            'Proje bitince Geçmiş > teklifi seçin > Gerçekleşen Saat Gir: her hizmet için gerçek saati yazın.',
            'Geçmiş > Tahmin / Gerçek Raporu: hizmet bazında sapma (örn. eğitim %+30) ve önerilen çarpan.',
            'Satırı seçip "önerilen çarpanı uygula": sonraki tekliflerde o hizmetin saati otomatik düzelir.',
        ]),
        ('h', 'Kontroller'),
        ('madde', [
            'KRİTİK: paketin kullanıcı sınırı aşıldı (açıklamada "Max 3 kullanıcı"), "modül eklenemez" paketine '
            'modül eklendi, modülün gerektirdiği ana modüller (DIA listesindeki Zorunlu kolonu) teklifte yok — '
            'eksikleri içeren paketler önerilir; toplam indirim max_toplam_indirim\'i, kâr marjı min_kar_marji\'nı '
            'aştı; saat ücreti taban altında veya tanımsız.',
            'UYARI: bir paketin modüllerinin yarıdan fazlası başka pakette zaten var (çift ücret), keşif formu boş, '
            'geçiş var ama kayıt sayısı yok, yetkili bilgisi eksik, fiyat listesi eski, seçilen kampanyanın koşulu '
            'sağlanmıyor, uyarlama/lisans oranı sınır dışında.',
            'Veriler > Ayarlar > yonetici_sifresi doluysa KRİTİK hatalı teklifin PDF/Excel çıktısı şifre ister '
            '(şifre açık yazılmaz, özeti saklanır).',
        ]),
        ('h', 'Satış asistanı neye göre öneri yapar'),
        ('madde', [
            'Rakip: aynı saat × rakip saat ücreti ortalamasıyla karşılaştırır; pahalıysanız ortalamaya inmek için '
            'gereken indirimi TL ve % olarak söyler.',
            'En yüksek indirim: kâr marjı alt sınırını ve toplam indirim üst sınırını bozmadan verilebilecek kampanya %.',
            'Kampanya uygunluğu: kampanyanın Koşul alanı (kullanici>=10, soru:gecis_programi=Logo|Netsis, '
            'soru:yeni_musteri, tarih:12-15..12-31; " ve " ile birleşir; "elle" = otomatik önerilmez).',
            'Geçmiş: sonuçlanan tekliflerin kazanma oranı, benzer büyüklükteki (±%40) tekliflerin oranı, kazanılan ve '
            'kaybedilenlerdeki ortalama indirim, kayıp nedenlerinde "fiyat" ağırlığı.',
            'Ödeme: peşin indirimin müşteriye kazandırdığı tutar; kredi kartında komisyonun size kalan kısmı ve onu '
            'karşılayan vade farkı.',
            'Tahmini kazanma ihtimali: 50 puandan başlar; rakip konumu (±), uygun kampanya (+5), kritik hata (-10), '
            'ihtiyaç-paket uyumsuzluğu (-5), keşif formu %60\'tan az dolu (-8) / eksiksiz (+3), karar verici ve '
            'iletişim bilgisi yok (-3), bakım anlaşması teklifte (+3); yeterli geçmiş varsa benzer tekliflerin '
            'kazanma oranıyla harmanlanır. Her puanın nedeni skorun altında yazar. Kesin değil, yön göstericidir.',
            'İhtiyaç-paket uyumu: keşifte "üretim var" denip üretim modülü yoksa, "e-dönüşüm kurulacak" denip '
            'e-Fatura paketi yoksa, mobil kullanıcı girilip mobil paket yoksa uyarır.',
            'Pazarlık aralığı (KDV hariç): Liste (indirimsiz) → Teklif → Hedef → Taban. Taban = en fazla indirim '
            '(max_toplam_indirim) ve en az kâr marjından (min_kar_marji, maliyet girildiyse) yüksek olanı; hedef '
            'teklif ile tabanın ortası.',
            'Taviz sırası: fiyattan önce verilebilecek tavizler (ek eğitim saati, bakımın ilk ayları, vade farksız '
            'taksit), en son fiyat indirimi ve inilebilecek en fazla tutar.',
            'İtiraz karşılama: "Fiyat yüksek", "Rakip daha ucuz", "Uyarlamayı sonra yaparız", "Düşünüp döneceğiz" '
            'için teklifin kendi rakamlarıyla hazır cevaplar.',
            'Takip planı: teklif tarihinden 2 ve 7 gün sonra takip, geçerlilik bitmeden 2 gün önce son hatırlatma.',
            'Hazırlık kontrol listesi (sol alt): müşteri, yetkili, iletişim, keşif %80+, kritik hata, ödeme, bakım, '
            'rakip verisi.',
            '"Satış brifingini kopyala": tüm bunları düz metin olarak panoya alır; CRM\'e, e-postaya veya ekip '
            'mesajına yapıştırın (müşteriye gönderilmez).',
        ]),
        ('h', 'Keşif formunu özelleştirme'),
        ('madde', [
            'Keşif & Uyarlama sekmesinde "Soruları Düzenle": soru metnini değiştirin (yeniden adlandırma), grup, cevap '
            'türü (Evet/Hayır, Sayı, Seçim listesi), seçenekler, standart cevap, açıklama, etkilediği hizmet ve ne '
            'yaptığı (saat ekle, çarp, % artır, miktar, sadece bilgi). Sağda "Bu soru: ..." satırı etkisini düz '
            'cümleyle gösterir. Yeni soru, kopyala, sırala (▲▼), pasif yap, sil.',
            'Standart cevap: yeni teklif keşif formu bu cevaplarla açılır (ör. şube 1, e-dönüşüm EVET). "Standart '
            'cevaplar" düğmesi formu bu değerlere döndürür.',
            'Profil: hazır cevap setleri (Perakende mağaza, Üretim firması, Toptan / dağıtım, Hizmet / ofis). '
            'Profili seçip "Uygula": standart cevaplar + profil cevapları yazılır. Mevcut cevapları "Profil kaydet" '
            'ile yeni profil yapabilirsiniz (cari / stok sayıları gibi müşteriye özel sayılar profile alınmaz).',
            'Form soruları grup başlıklarıyla gösterilir; soru üzerine gelince açıklama ve etkisi görünür; '
            'cevabın yanında o cevabın etkisi (+8 saat, × 1,15, 1.200 adet) yazar; üstte form doluluğu ve boş '
            'kalan sorular gösterilir.',
            'Kod alanı değişmez: kampanya koşulları (soru:kod) bu kodla çalışır.',
        ]),
        ('h', 'Taksit, vade farkı ve kredi kartı komisyonu'),
        ('tablo', [['Tür', 'Hesap'], [
            ['PESIN', 'net × (1 - indirim %)'],
            ['TAKSIT (havale, çek/senet)', 'net × (1 + vade farkı % + aylık vade % × (taksit+1)/2)'],
            ['KREDI_KARTI', 'müşteri: net × (1 + vade farkı %); banka komisyonu: KDV dahil tahsilat × komisyon %; '
                            'size kalan = KDV hariç tutar - komisyon'],
        ]]),
        ('p', 'Komisyonu tam karşılayan vade farkı = 1 / (1 - komisyon × (1 + KDV)) - 1. Örn. %4,5 komisyon, %20 KDV '
              'için yaklaşık %5,7. Bankanızın oranlarını Veriler > Ödeme Seçenekleri\'ne girin.'),
        ('h', 'Bakım anlaşması neye göre belirlenir'),
        ('p', 'Yıllık destek saati = taban saat + kullanıcı × destek_kullanici_saat + modül sayısı × destek_modul_saat '
              '+ entegrasyon × destek_entegrasyon_saat + ek şube × destek_sube_saat + uyarlama saatinin '
              'destek_uyarlama_orani %\'si; IT personeli yok gibi genel katsayılar da uygulanır. Paket fiyatı = '
              '(yıllık saat × paket çarpanı × destek saat ücreti + yerinde gün × gün ücreti) × (1 - indirim). '
              'Önerilen paket yıllık saate göre seçilir (destek_esikleri: örn. 40 saatin altı Temel, 100\'ün altı '
              'Standart, üstü Premium). Paketler (saat çarpanı, yanıt süresi, yerinde gün, indirim, kapsam) '
              'Veriler > Bakım Paketleri\'nden değiştirilir. DIA\'nın yıllık bakımı (YBH) ayrıca hesaplanır ve '
              '3 yıllık toplam maliyet tablosunda gösterilir.'),
        ('h', 'Alternatif teklif ve revizyon'),
        ('madde', [
            'Aynı müşteri için Ekonomik / Önerilen / Kapsamlı teklifleri ayrı kaydedin (Etiket alanı). Geçmiş\'te '
            'Ctrl ile 2-4 tanesini seçip "Seçilenleri Karşılaştır": yan yana karşılaştırma PDF\'i.',
            'Revizyon: RT-2026-0001 → RT-2026-0001-R1; PDF başlığında "Revizyon 1" yazar, kayıtta önceki '
            'sürümle genel toplam farkı gösterilir.',
        ]),
        ('h', 'DİA paket indirimleri ve DİA ekranıyla karşılaştırma'),
        ('madde', [
            'Paket indirimi paketin kendisindedir: Veriler > Paketler / Adam-Saat sekmesinde Kamp. İnd. %, Kamp. İnd. '
            'Tutar ve Kamp. Açıklama kolonları (çift tıklayıp elle de girebilirsiniz).',
            'Her ay DİA paket listesi Excel\'ini "DIA Fiyat Listesini Güncelle" ile aktarın. Listede indirim kolonları '
            'varsa (başlık örnekleri: "İndirim %", "İndirim Tutarı", "Kampanya" veya "Kamp.İnd.%", "Kamp.İnd.Tutar", '
            '"Kamp.Açıklama") paketlerin indirimi o aya göre yenilenir; 0 olanların indirimi kalkar. Listede indirim '
            'kolonu yoksa paketlere elle girdiğiniz indirimler korunur. Adam/saatler her durumda korunur.',
            'Teklife paket eklenince indirimi kendiliğinden satırın İnd. % / İnd. Tutar kolonlarına yazılır ve '
            '"Kamp. Açıklama" kolonuna kampanya adı gelir (DİA üye ekranındaki gibi tek yerde). Tutarı satırda '
            'değiştirebilirsiniz. Kamp. Açıklama doluysa o indirim DİA kampanyası sayılır ve sizin "toplam indirim" '
            'sınırınıza girmez; açıklamayı silerseniz sizin indiriminiz sayılır.',
            'Ekranın altındaki mavi satır DİA üye ekranıyla aynı rakamları verir: Üyelik Giriş Tutarı (lisans net), '
            'Ek +1 Masaüstü / Mobil (birim × Kull. %), Yıllık Bulut Hizmet Bedeli (liste × YBH %). Örnek: Perakende '
            'Pro + Genel Muhasebe + POS 1 + POS 1+14 + E-Fatura (−7.800) → 297.900 / 17.100 / 99.300.',
            'Tutar yazarken 7.800, 7 800, 7.800,00 veya ₺7.800 yazabilirsiniz; hepsi 7800 olarak okunur '
            '(eskiden 7.800 → 7,8 okunuyordu).',
            'Adam/saati olmayan bir Excel içe aktarılırsa önceki adam/saatler korunur; A/S\'si 0 olan paket teklife '
            'eklenirse Kontroller\'de uyarı çıkar (kurulum saati eksik hesaplanmasın).',
        ]),
        ('h', 'PDF / Excel tasarımı (Tasarım düğmesi)'),
        ('madde', [
            'Hazır tasarımlar: Kurumsal (logo solda, firma bilgileri sağda), Kurumsal Klasik (ortalanmış logo, çift '
            'çizgi), Kurumsal Kapaklı (ilk sayfa kapak: logo, müşteri, teklif no, hazırlayan), Kurumsal Şeritli (her '
            'sayfada solda renkli şerit), Modern (renkli bant başlık), Sade (siyah-beyaz). Kopyalayıp kendi '
            'tasarımınızı oluşturabilirsiniz.',
            'Görünüm: şablon, ana ve vurgu rengi, yazı boyutu, teklif başlığı (Türkçe / İngilizce), logo (PNG/JPG).',
            'Uyarlama gösterimi: Detaylı (kalem + kapsam + dayanak + hesap), Kalem listesi (kalem + kapsam), Özet '
            '(tek kutu: toplam tutar, saat ve kalem adları). Teklif özeti: Detaylı, Kısa (net / KDV / genel toplam) '
            'veya Gizle. Bakım bedeli: Aylık, Aylık + Yıllık veya Yıllık. Bakım bölümünü tamamen kapatmak için '
            'Gösterilecek Alanlar > Bakım.',
            'Gösterilecek Alanlar: ön yazı, Teklifi Veren / Teklif Edilen kutuları, hazırlayan, paket kodları, paket '
            'içeriği, liste fiyatı / indirim kolonları, ek kullanıcı ve yıllık bulut bedeli, uyarlama dayanağı ve '
            'hesabı, ödeme seçenekleri, bakım, toplam maliyet, ticari şartlar, banka, notlar, imza, alt bilgi.',
            'Metinler (Türkçe ve English sekmeleri): ön yazı, ticari şartlar (her satır bir madde; tasarımla '
            'sabitlenir), kapanış. Sabit notlar her teklifte yazılır; teklif ekranındaki notu "Bu notu sabit '
            'notlara ekle" ile de sabitleyebilirsiniz. {musteri} {yetkili} {tarih} {gecerlilik} {teklif_no} {firma} '
            '{hazirlayan} alanları otomatik doldurulur.',
            'Firma ve Hazırlayan: adres, telefon, e-posta, web, vergi no, banka/IBAN, teklifi hazırlayan kişi. '
            'Bunlar tüm tasarımlarda ortaktır.',
            'Önizle ile PDF/Excel\'i hemen görün; "Kaydet ve Varsayılan Yap" ile tasarım sabitlenir, yeni teklifler '
            'bu tasarımla açılır. Teklif ekranındaki "PDF tasarımı" kutusundan teklife özel başka tasarım seçilebilir '
            've teklifle birlikte kaydedilir.',
            'PDF\'te alt bilgi: firma iletişim bilgileri ve "Sayfa x / y". Excel\'de logo, renkler ve yazdırma alt '
            'bilgisi de uygulanır.',
        ]),
        ('h', 'Uyarlama kalem adları, döviz ve İngilizce teklif'),
        ('madde', [
            'Keşif & Uyarlama sekmesinde hizmet adına çift tıklayın: o teklif için adı ve kapsamı değiştirin '
            '(PDF/Excel\'e bu yazılır). "Sonraki tekliflerde de kullanılsın" işaretlenirse Veriler > Hizmet '
            'Kalemleri\'ne kaydedilir. "Tanımlı ada dön" değişikliği geri alır.',
            'Para birimi TL / USD / EUR. Kur elle girilir (1 USD = ? TL); "Kuru varsayılan yap" ile kaydedilir ve '
            'sonraki dövizli tekliflerde hazır gelir. Hesap TL\'de yapılır, PDF/Excel tutarları kura bölünerek '
            'basılır; ticari şartlara kur ve faturalama notu eklenir. Ekranda dövizli net ve KDV dahil tutar görünür.',
            'Teklif dili Türkçe / English: başlık, bölüm adları, tablo başlıkları, varsayılan hizmet adları, ödeme '
            'seçenekleri ve tasarımın İngilizce metinleri kullanılır. Paket adları ve elle yazılan notlar olduğu '
            'gibi kalır.',
        ]),
        ('h', 'Satış süreci, takip ve "Bugün" listesi'),
        ('madde', [
            'Kaydedilen teklif "Teklif verildi" aşamasında başlar ve takip tarihi atanır (takip_ilk_gun, varsayılan 2 '
            'gün). Aşamalar: Fırsat, Keşif, Teklif verildi, Müzakere, Sözlü onay, Kazandı, Kaybetti.',
            'Satış Asistanı sekmesi > Satış süreci: aşama, sonraki takip tarihi (+2 / +7 gün), görüşme / aktivite '
            'ekle (Arama, Toplantı, Demo, E-posta, Ziyaret, Not; ne konuşuldu ve sonraki takip). Aktivite geçmişi '
            'altta listelenir.',
            'Üstteki "Bugün (N)" düğmesi: gecikmiş ve bugünkü takipler, geçerliliği 3 gün içinde bitecek teklifler, 7 '
            'gündür hareketsiz teklifler, kazanılan müşteride memnuniyet / bakım görüşmesi (kazanilan_takip_gun, '
            'varsayılan 30), kaybedilen müşteride tekrar temas (kaybedilen_takip_gun, varsayılan 120). Gecikmiş iş '
            'varsa düğme kırmızı olur ve program açılışta listeyi gösterir (acilista_gunluk).',
            'Geçmiş listesinde aşama, takip tarihi, kazanma olasılığı ve hazırlayan kolonları vardır.',
        ]),
        ('h', 'Satış asistanı: sonraki adımlar ve geçmişten öğrenme'),
        ('madde', [
            'Sonraki adımlar: teklifin aşamasına göre ne yapılacağını sırayla yazar (Veriler > Satış Senaryoları > '
            'ASAMA). Karar verici bilgisi yoksa, keşif eksikse, kritik hata varsa bunları en üste ekler.',
            'Satış Senaryoları tablosunu kendi satış yönteminize göre değiştirin: ASAMA (aşama adımları), ITIRAZ '
            '(itiraz ve cevabı), TAVIZ (fiyattan önce verilecek tavizler). Metinlerde {saat} {gun} {dokum} {hedef} '
            '{taban} {gecerlilik} {bakim} {rakip_ort} {biz} teklifin rakamlarıyla doldurulur.',
            'Geçmişten öğrenilenler (sonuçlanan teklif sayısı arttıkça güçlenir): bu müşteriye önceki teklifler, '
            'kazanma oranı en yüksek indirim aralığı, bakım anlaşmalı / anlaşmasız kazanma oranı, paket tipine göre '
            'kazanma, kazanılan tekliflerin ortalama kapanma süresi (bu teklif uzadıysa uyarır), en sık kayıp '
            'nedeni. Bunun için teklifleri KAZANDI / KAYBETTİ işaretleyin ve kayıp nedenini yazın.',
        ]),
        ('h', 'Satış raporu (Rapor düğmesi)'),
        ('madde', [
            'Dönem (son 3/6/12 ay, bu yıl, tümü) ve hazırlayan seçilir.',
            'Kartlar: teklif sayısı ve tutarı, kazanılan, kaybedilen, bekleyen, kazanma oranı (adet ve tutar), '
            'ortalama karar süresi, ağırlıklı bekleyen (bekleyen tutar × kazanma olasılığı).',
            'Aylık grafik (teklif ve kazanılan tutar, kazanma %), aşama hunisi, kayıp nedenleri, aktivite sayıları; '
            'Aylık, Satışçı, İndirim etkisi ve Paketler tabloları. "Excel\'e aktar" tüm tabloları ayrı sayfalara yazar.',
        ]),
        ('h', 'Excel\'de düzeltilen hatalar'),
        ('madde', [
            'K24 hücresindeki bozuk formül (#NAME?) lisans neti, ara toplam, net toplam ve ödeme planını bozuyordu.',
            'İnd. Tutar kolonu KDV hariç toplamı ve KDV alt toplamını etkilemiyordu.',
            'Ödeme planı ve şablondaki NET TOPLAM, net tutar yerine kampanya indirim tutarından hesaplanıyordu.',
            'Adam/saat toplamı SUM(O17:O24) idi; ilk 7 paketin saati uyarlamaya girmiyordu.',
            'Satır indirimi birim fiyattan (toplamdan değil) hesaplanıyordu.',
            'Kampanya seçimi indirim yüzdesine bağlı değildi.',
            'ADAM SAAT satırları paket listesine satır sırasıyla bağlıydı; artık paket koduyla tutulur, DIA fiyat '
            'listesi güncellense de saatler korunur.',
        ]),
        ('h', 'Teklif geçmişi'),
        ('p', 'Teklif Geçmişi penceresinde arama, Aç/Düzenle, Kopyala (aynı teklifi yeni numarayla), KAZANDI / '
              'KAYBETTİ (red nedeni sorulur) / BEKLİYOR işaretleme ve listeyi Excel\'e aktarma. Altta kazanma oranı '
              've kazanılan tutar görünür.'),
        ('h', 'Veriler / Ayarlar'),
        ('madde', [
            'Ayarlar, Sabit Kalemler, Kampanyalar, Rakipler ve Paketler / Adam-Saat sekmelerinde çift tık ile '
            'düzenleme; değişiklik hemen kaydedilir.',
            'DIA Fiyat Listesini Güncelle: DIA\'nın yeni fiyat Excel\'ini seçin ("Kodu" ve "Std. Fiyat" başlıkları '
            'olmalı). Fiyatı değişen, yeni gelen ve listeden kalkan paketler raporlanır; kalkanlar silinmez, '
            'pasif yapılır (eski teklifler bozulmaz). Adam/saatler korunur.',
        ]),
        ('tablo', [['Ayar', 'Varsayılan', 'Açıklama'],
                   [[a, '(boş)' if v == '' else str(v).replace('\n', ' '), ac] for a, v, ac in TEKLIF_AYARLARI]]),
        ('h', 'Veri, yedek ve kayıp kayıt sorunu'),
        ('p', 'Veriler EXE\'nin yanına değil Belgeler\\RACA_Teklif_Veri\\teklif.db dosyasına yazılır. Böylece EXE '
              'İndirilenler\'den, zip\'in içinden çalıştırılsa ya da yeni sürüm başka klasöre konsa da kayıtlar '
              'kaybolmaz.'),
        ('madde', [
            'Otomatik yedek: her gün ilk açılışta, teklif kaydında (en fazla 30 dakikada bir), içe aktarma ve '
            'fiyat güncellemesinden önce ve sonra, program kapanırken. Son 30 yedek Yedekler klasöründe tutulur.',
            'Veriler > Yedek Al: elle yedek. Yedekten Geri Yükle: seçilen yedeğe döner (mevcut hal önce yedeklenir).',
            'Veriler > Veri Klasörü: klasörü açar veya değiştirir (örn. satış ekibi için ortak ağ klasörü; mevcut '
            'veri oraya kopyalanır).',
            'Bilgisayar değiştirirken Belgeler\\RACA_Teklif_Veri klasörünü yeni bilgisayara kopyalamanız yeterli.',
        ]),
    ]})

    # ------------------------------------------------------------------
    B.append({'baslik': '10. Geliştirici: Yeni Ayar ve Kontrol Eklemek', 'parcalar': [
        ('p', 'Kaynak kod: github.com/yusufkaraca096/yusufkaraca096. Değişiklik gönderilince testler çalışır ve '
              'yeni EXE aynı indirme linkine konur.'),
        ('h', '1) Ayarı tanımlayın: src/ek_ayarlar.py > EK_AYAR_TANIMLARI'),
        ('kod', "{'ayar': 'stok_kodu_on_ek', 'varsayilan': '', 'tur': 'metin', 'moduller': ['stok'],\n"
                " 'aciklama': 'Stok kodlarının başına eklenecek ön ek.',\n"
                " 'formul': 'on_ek + kod', 'ornek': 'ST- + 001 -> ST-001'},"),
        ('madde', [
            "tur: 'bool' (True/False), 'secim' (+ 'secenekler': [...]), 'metin'.",
            "moduller: None = tüm modüller; liste = sadece o modüller.",
            "varsayilan_modul: {'cari': ..., 'stok': ...} modül bazında farklı varsayılan.",
            'Ayarlar penceresi, kurallar.xlsx ve bu rehber ayarı kendiliğinden gösterir.',
        ]),
        ('h', '2) Kodu yazın: src/ek_kontroller.py'),
        ('kod', "def stok_on_islem(df, modul, rapor):\n"
                "    on_ek = ayar(modul, 'stok_kodu_on_ek', 'stok')\n"
                "    if on_ek:\n"
                "        df['stok kart kodu'] = [on_ek + str(k) for k in df['stok kart kodu']]\n"
                "    ..."),
        ('madde', [
            'on_islem: değer düzelten işlemler (il/ilçe düzeltmesinden önce çalışır).',
            'son_kontroller: sadece uyarı üreten kontroller (en sonda).',
            "Uyarı: _ekle(rapor, satir_index, 'kolon', deger, 'UYARI: ...') — KONTROL_UYARI ve OZET'e yazılır.",
        ]),
        ('h', '3) Test yazın: tests/test_ek_kontroller.py'),
        ('kod', 'python -m unittest discover -s tests'),
        ('p', 'Rehberi (docs/KULLANIM_REHBERI.md) güncellemek için: python src/rehber.py'),
    ]})
    return B


# ======================================================================
# Ciktilar
# ======================================================================
def markdown():
    satirlar = ['# DIA Veri Temizleme - Kullanım Rehberi', '',
                '> Bu dosya `src/rehber.py`\'den üretilir (`python src/rehber.py`). Aynı rehber programda '
                'Rehber butonundadır.', '']
    for b in bolumler():
        satirlar += ['## ' + b['baslik'], '']
        for tur, icerik in b['parcalar']:
            if tur == 'p':
                satirlar += [icerik, '']
            elif tur == 'h':
                satirlar += ['### ' + icerik, '']
            elif tur == 'madde':
                satirlar += ['- ' + m for m in icerik] + ['']
            elif tur == 'adim':
                satirlar += [str(n) + '. ' + m for n, m in enumerate(icerik, 1)] + ['']
            elif tur == 'kod':
                satirlar += ['```', icerik, '```', '']
            elif tur == 'not':
                satirlar += ['> **Dikkat:** ' + icerik, '']
            elif tur == 'tablo':
                baslik, govde = icerik
                bas = [h or ' ' for h in baslik]
                satirlar += ['| ' + ' | '.join(bas) + ' |', '|' + '---|' * len(bas)]
                for s in govde:
                    satirlar.append('| ' + ' | '.join(str(x).replace('|', '\\|').replace('\n', ' ') for x in s) + ' |')
                satirlar.append('')
    return '\n'.join(satirlar)


def html_sayfa():
    e = html.escape
    govde = []
    icindekiler = []
    for n, b in enumerate(bolumler()):
        icindekiler.append('<li><a href="#b%d">%s</a></li>' % (n, e(b['baslik'])))
        govde.append('<h2 id="b%d">%s</h2>' % (n, e(b['baslik'])))
        for tur, icerik in b['parcalar']:
            if tur == 'p':
                govde.append('<p>' + e(icerik) + '</p>')
            elif tur == 'h':
                govde.append('<h3>' + e(icerik) + '</h3>')
            elif tur in ('madde', 'adim'):
                et = 'ul' if tur == 'madde' else 'ol'
                govde.append('<%s>%s</%s>' % (et, ''.join('<li>' + e(m) + '</li>' for m in icerik), et))
            elif tur == 'kod':
                govde.append('<pre>' + e(icerik) + '</pre>')
            elif tur == 'not':
                govde.append('<div class="not"><b>Dikkat:</b> ' + e(icerik) + '</div>')
            elif tur == 'tablo':
                baslik, satirlar = icerik
                t = ['<table>']
                if any(baslik):
                    t.append('<tr>' + ''.join('<th>' + e(h) + '</th>' for h in baslik) + '</tr>')
                for s in satirlar:
                    t.append('<tr>' + ''.join('<td>' + e(str(x)).replace('\n', '<br>') + '</td>' for x in s) + '</tr>')
                t.append('</table>')
                govde.append(''.join(t))
    stil = ('body{font-family:Segoe UI,Arial,sans-serif;max-width:1000px;margin:24px auto;padding:0 16px;'
            'color:#222;line-height:1.5}h1{color:#2c3e50}h2{color:#fff;background:#2c3e50;padding:6px 10px;'
            'margin-top:36px}h3{color:#2c3e50;border-bottom:1px solid #ccc}table{border-collapse:collapse;'
            'width:100%;margin:8px 0}td,th{border:1px solid #ccc;padding:4px 8px;vertical-align:top}'
            'th{background:#ecf0f1}pre{background:#f4f6f7;border:1px solid #ddd;padding:8px;white-space:pre-wrap}'
            '.not{background:#fff3cd;border-left:4px solid #f39c12;padding:6px 10px;margin:8px 0}'
            '@media print{h2{break-before:page}}')
    return ('<!doctype html><html lang="tr"><head><meta charset="utf-8"><title>DIA Veri Temizleme - Rehber</title>'
            '<style>' + stil + '</style></head><body><h1>DIA Veri Temizleme - Kullanım Rehberi</h1>'
            '<ul>' + ''.join(icindekiler) + '</ul>' + '\n'.join(govde) + '</body></html>')


if __name__ == '__main__':
    import os
    yol = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'docs', 'KULLANIM_REHBERI.md')
    with open(yol, 'w', encoding='utf-8', newline='\n') as f:
        f.write(markdown())
    print('Yazildi: ' + os.path.normpath(yol))
