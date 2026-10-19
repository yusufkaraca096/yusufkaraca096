# DIA Veri Temizleme - Kullanım Rehberi

> Bu dosya `src/rehber.py`'den üretilir (`python src/rehber.py`). Aynı rehber programda Rehber butonundadır.

## 1. Başlarken

Program, müşteriden gelen Excel listelerini (cari, stok, banka, personel, demirbaş, hesap planı, muhasebe kodu, ürün ağacı) temizleyip DIA'ya yüklenebilir hale getirir. Hiçbir zaman müşterinin orijinal dosyasını değiştirmez; sonuçları yeni dosyalara yazar, eskisini .bak olarak yedekler.

### Klasör düzeni

Her müşteri için ayrı bir klasör açın. Müşterinin dosyalarını şu adlarla koyun (adlar Ayarlar > girdi_dosya ile değiştirilebilir):

| Modül | Girdi dosyası | Çıktı (DIA'ya yüklenecek) |
|---|---|---|
| Cari | cari_listesi.xlsx | cari_dia_hazir.xlsx |
| Stok | stok_listesi.xlsx | stok_dia_hazir.xlsx |
| Banka | banka_listesi.xlsx | banka_dia_hazir.xlsx |
| Personel | personel_listesi.xlsx | personel_dia_hazir.xlsx |
| Demirbaş | demirbas_listesi.xlsx | demirbas_dia_hazir.xlsx |
| Hesap Planı | hesapplan_listesi.xlsx | hesapplan_dia_hazir.xlsx |
| Muhasebe Kodu | muhkod_listesi.xlsx | muhkod_dia_hazir.xlsx |
| Ürün Ağacı | urunagaci_listesi.xlsx | urunagaci_dia_hazir.xlsx |

Birden fazla müşteriyi tek seferde işlemek için program klasörünün yanında "musteriler" klasörü açıp her müşteriyi alt klasör olarak koyun ve modülde HEPSİ seçin.

### Çalıştırma adımları

1. Klasör Seç ile müşteri klasörünü seçin.
2. Modülü seçin (Cari, Stok ... veya HEPSİ).
3. İşlem modunu seçin (aşağıdaki tablo).
4. ▶ ÇALIŞTIR'a basın. Alttaki pencerede ne yapıldığı satır satır yazar.
5. Çıkan dosyaları Klasörü Aç ile açıp kontrol edin.

| İşlem modu | Ne yapar | Ne zaman |
|---|---|---|
| İnceleme | Hiçbir satırı silmez. _INCELEME.xlsx üretir; mükerrer satırlar sarı, SIL_ONAY kolonu yeşil. | İlk çalıştırma; müşteriye sormak istediğiniz satırlar varsa. |
| Uygula | _INCELEME.xlsx'te SIL_ONAY kolonuna E yazdığınız satırları siler, _dia_hazir.xlsx yazar. | İnceleme dosyasını düzenledikten sonra. |
| Otomatik Sil | Mükerrer sekmesinde aksiyonu "sil" olan kurallara uyan satırları kendisi siler. | Kurallara güveniyorsanız, hızlı sonuç için. |
| Kuru Çalıştır | Hiç dosya yazmaz, sadece ne yapılacağını raporlar. | Ayar denerken. |

> **Dikkat:** İnceleme modunda _dia_hazir.xlsx oluşmaz. DIA'ya gönderim için önce Uygula veya Otomatik Sil çalıştırılmalıdır.

## 2. Ayarlar Penceresi

Ana ekranda ⚙ Ayarlar. Ayarlar müşteri klasöründeki kurallar.xlsx dosyasına yazılır; her müşterinin ayarları ayrıdır. Klasörde kurallar.xlsx yoksa program klasöründeki kopyalanır, o da yoksa varsayılanlardan oluşturulur.

1. Üstteki MODÜL kutusundan modülü seçin (ayarlar modül bazındadır).
2. Sekmeyi seçin: Mükerrer, Eşleşme, Varsayılan veya Ayarlar.
3. Satırı seçip ✎ Düzenle (veya + Yeni / Sil).
4. Bitince KAYDET. KAYDET'e basmadan pencereyi kapatırsanız değişiklikler kaybolur.

### Sekmeler

| Sekme | Ne içindir | Örnek satır |
|---|---|---|
| Mükerrer | Hangi kolon(lar) aynıysa kayıt mükerrer sayılır, ne yapılır (isaretle / sil). | sutunlar: vergidairesi, verginumarasi \| aksiyon: sil |
| Eşleşme | Müşterinin kolon başlığı → DIA kolon adı. | excel_baslik: Firma Ünvanı \| dia_alan: unvan |
| Varsayılan | Kolon yoksa eklenir, boş hücreler bu değerle dolar. | alan: doviz \| deger: TL |
| Ayarlar | Temizlik ve kontrollerin açık/kapalı ve parametreleri. | telefon_format \| 10 |

### Ayarlar sekmesi butonları

- Satıra çift tık veya ℹ Detay: ayarın açıklaması, formülü ve örneği büyük pencerede.
- ＋ Eksik Ayarları Ekle: eski bir kurallar.xlsx'te olmayan yeni ayarları varsayılan değerleriyle ekler. Program güncellenince bir kez basıp KAYDET deyin.
- Ayarlar sekmesinde satırı olmayan ayar için de varsayılan değer kullanılır; satırı silmek ayarı kapatmaz, varsayılana döndürür. Kapatmak için değeri False (liste ayarlarında YOK) yapın.

## 3. Değer Yazım Kuralları (Formül Dili)

Ayarlar sekmesindeki "deger" kolonuna yazılan değerler aşağıdaki biçimlerden birindedir. Her ayarın biçimi Ayar Referansı bölümünde "Formül / yazım" satırında yazar.

| Biçim | Nasıl yazılır | Örnek |
|---|---|---|
| Açık/Kapalı | True veya False. Evet, E, 1, X da açık sayılır. | True |
| Seçim | Seçeneklerden biri, aynen. | BUYUK |
| Sayı | Sadece rakam. | 88 |
| Kolon adı | DIA kolonunun adı, büyük/küçük harf ve Türkçe karakter AYNEN. | stok adı |
| Liste | Virgülle ayrılmış değerler. | carikartkodu, unvan, verginumarasi |
| Eşleme çiftleri | ARANAN=YAZILACAK, çiftler noktalı virgülle. | ADT=ADET; KILO=KG |
| Uzunluk | kolon=karakter sayısı, noktalı virgülle. | carikartkodu=25; unvan=200 |
| Çapraz kontrol | kolon>modül:kolon1,kolon2 ; birden fazla kural noktalı virgülle. | Grup Kodu>urunagaci:Grup1,GRUP2 |

### Eşleme çiftlerinde arama nasıl yapılır

- Büyük/küçük harf fark etmez: "adt", "ADT", "Adt" hepsi ADT=… ile eşleşir.
- Türkçe karakter fark etmez: "KILO" yazmanız "Kilo", "KİLO", "kılo"yu da yakalar. Aranan tarafı Türkçe karaktersiz yazmanız yeterli; yazılacak tarafı istediğiniz gibi (ÇİFT, KOLİ) yazın.
- Ünvan ve adres kısaltmalarında nokta ve boşluk esnektir: "LTD STI" aranırsa "Ltd.Şti", "LTD. ŞTİ.", "Ltd Şti" hepsi bulunur. Sadece tam kelime bulunur: "SAN" aranırken "SANAT" değişmez.
- Birim eşlemesinde hücrenin tamamı karşılaştırılır (nokta ve boşluk yok sayılır): "A D T" = ADT.
- Aranan ifadenin sonuna $ koyarsanız sadece metnin SONUNDA aranır: "AS$=A.Ş." → "DENİZ AS" değişir, "AS YAPI" değişmez.
- Uzun ifadeler önce uygulanır: "LIMITED SIRKETI" ve "SIRKETI" ikisi de varsa önce uzun olan çalışır.

### Boş ve YOK

- Değer boş bırakılırsa varsayılan kullanılır.
- Liste ve eşleme ayarlarını tamamen kapatmak için YOK yazın (örn. zorunlu_alanlar = YOK).
- Açık/Kapalı ayarlar için False yazın.

> **Dikkat:** Kolon adları birebir eşleşmelidir. "Stok Adı" ile "stok adı" farklı kolonlardır. Doğru adı görmek için _dia_hazir.xlsx'in başlık satırına bakın.

## 4. Örneklerle Ayar Yapma (Tarifler)

Sık yapılan işlemler adım adım. Her işlemden sonra KAYDET'e basıp modülü yeniden çalıştırın.

### Listede olmayan birimi tanıtmak (TANIMSIZ BİRİM uyarısı)

1. Ayarlar > MODÜL: stok > Ayarlar sekmesi.
2. birim_eslesme satırını seçip ✎ Düzenle.
3. Değerin SONUNA noktalı virgülle yeni çiftleri ekleyin (mevcutları silmeyin):

```
…; TOP=TOP; CUVAL=ÇUVAL; CVL=ÇUVAL; BIDON=BİDON
```

KAYDET, sonra modülü yeniden ÇALIŞTIR. Satır yoksa önce ＋ Eksik Ayarları Ekle.

### Yeni ünvan kısaltması eklemek

1. MODÜL: cari > Ayarlar > unvan_kisaltmalari > Düzenle.
2. Sona ekleyin:

```
…; INSAAT=İNŞ.; PAZARLAMA=PAZ.; ITHALAT IHRACAT=İTH. İHR.
```

Bir kısaltmayı istemiyorsanız (örn. TİCARET açık yazılsın) o çifti silin: "TICARET=TİC.; TIC=TİC." kısmını kaldırın. Hiç kısaltma istemiyorsanız unvan_standart = False.

### İl/ilçe adı olmayan vergi dairelerini düzeltmek

1. MODÜL: cari > vergi_dairesi_eslesme > Düzenle (varsayılan boştur).
2. Müşterinin yazımı = doğru yazım çiftlerini girin:

```
KOZYATAGI=KOZYATAĞI; BUYUK MUKELLEFLER=BÜYÜK MÜKELLEFLER; ULUS=ULUS; GOKSU=GÖKSU
```

Aranan tarafta Türkçe karakter ve boşluk/nokta önemli değildir. "V.D.", "Vergi Dairesi Müdürlüğü" ekleri zaten silinir, eklemenize gerek yok.

### Zorunlu alan eklemek veya kaldırmak

1. İlgili modül > zorunlu_alanlar > Düzenle.
2. Kolon adlarını virgülle yazın:

```
carikartkodu, unvan, verginumarasi, sehir
```

Boş olan satırlar KONTROL_UYARI'da "GEÇERSİZ: ZORUNLU ALAN BOŞ" olarak kırmızı görünür. Kontrolü kapatmak için YOK yazın.

### Alan uzunluğunu DIA'ya göre ayarlamak

1. DIA ekranında alanın kaç karakter aldığına bakın.
2. İlgili modül > alan_uzunluklari > Düzenle:

```
carikartkodu=20; unvan=250; adres1=150; vergidairesi=40
```

Daha uzun değerler "GEÇERSİZ: UZUNLUK 130 > 100 KARAKTER" uyarısı alır (değer kesilmez).

### Telefon biçimini değiştirmek

1. İlgili modül > telefon_format > Düzenle.
2. Seçeneklerden birini yazın:

```
10   → 5321234567
11   → 05321234567
90   → 905321234567
+90  → +905321234567
```

Hangi kolonların telefon sayılacağı telefon_alanlari ayarındadır.

### Cep numarasının ceptel'e taşınmasını kapatmak

1. MODÜL: cari > cep_ayir > Düzenle > False.

```
False
```

Başka bir cep kolonu kullanıyorsanız cep_sutunu ayarına o kolonun adını yazın.

### Benzer kayıt (olası mükerrer) hassasiyeti

1. İlgili modül > benzer_kayit_esik > Düzenle.

```
92
```

Yükseltirseniz (95) sadece çok benzerler, düşürürseniz (80) daha fazla kayıt işaretlenir. Karşılaştırılan kolonları benzer_kayit_sutunu belirler; örn. personelde "Adı, Soyadı".

### Yeni mükerrer kuralı eklemek

1. İlgili modül > Mükerrer sekmesi > + Yeni.
2. sutunlar: virgülle kolonlar; aksiyon: isaretle veya sil.

```
sutunlar: unvan, sehir
aksiyon : isaretle
```

"sil" aksiyonu sadece 2 veya daha fazla kolonlu kurallarda ve Otomatik Sil modunda uygulanır; tek kolonlu kurallar sadece işaretler (yanlışlıkla toplu silme olmasın diye).

### Müşterinin kolon başlığını tanıtmak

1. İlgili modül > Eşleşme sekmesi > + Yeni.
2. excel_baslik: müşterinin dosyasındaki başlık (aynen); dia_alan: DIA kolonu.

```
excel_baslik: Firma Ünvanı
dia_alan    : unvan
```

Daha kolayı: ana ekranda Eşleme Öğren dosyayı tarar, tanımadığı başlıkları listeler; açılır listeden seçip kaydedin. esleme_hafizasi açıksa bu eşlemeler sonraki müşterilerde de kendiliğinden uygulanır.

### Boş hücrelere sabit değer yazmak

1. İlgili modül > Varsayılan sekmesi > + Yeni.

```
alan : doviz
deger: TL
```

Kolon yoksa eklenir ve tüm satırlara yazılır; varsa sadece boş hücreler dolar.

### TDHP dışı ana hesap kullanan müşteri

1. MODÜL: hesapplan > ek_ana_hesaplar > Düzenle.

```
125, 175, 850
```

Bu hesaplar artık "TEK DÜZEN HESAP PLANINDA YOK" uyarısı almaz.

### Ayraçsız hesap kodlarını bölmek

1. MODÜL: hesapplan > hesap_seviye_uzunluklari > Düzenle.

```
3,2,3
```

"12001001" → "120.01.001". Ayrıca 120.1 gibi seviyesi kısa kalan kodlar uyarı alır (Excel sayı olarak okuyup 120.10'un sonundaki 0'ı silmiş olabilir).

### Modüller arası yeni çapraz kontrol

1. Kontrol edilecek modül > capraz_kontrol > Düzenle.
2. Mevcut kuralın sonuna noktalı virgülle ekleyin:

```
Grup Kodu>urunagaci:Grup1,GRUP2,GRUP3; Özel Kod 1>muhkod:STOK ÖZEL KOD 1
```

Hedef modülün dosyası (önce _dia_hazir, yoksa _listesi) aynı müşteri klasöründe olmalıdır.

### Bir kontrolü tamamen kapatmak

1. Ayarın değerini False yapın (liste ayarlarında YOK).

```
barkod_kontrol = False
zorunlu_alanlar = YOK
```

### Ayarları tüm yeni müşterilere varsayılan yapmak

1. Bir müşteride ayarları istediğiniz gibi yapıp KAYDET.
2. O müşterinin klasöründeki kurallar.xlsx'i program (EXE) klasörüne kopyalayın.

Bundan sonra kurallar.xlsx'i olmayan her yeni müşteri klasörüne bu dosya kopyalanır. Mevcut müşterilerin dosyaları değişmez.

## 5. Raporları Okumak

| Dosya | İçerik |
|---|---|
| X_dia_hazir.xlsx / .xls | DIA'ya yüklenecek temiz veri. KONTROL_UYARI kolonunda satırın uyarıları. |
| X_dia_hazir_INCELEME.xlsx | İnceleme modu: SIL_ONAY kolonuna E yazılan satırlar Uygula'da silinir. |
| X_dia_hazir_DEGISIKLIKLER.xlsx | DEGISIKLIKLER: her hücre için eski → yeni. TEMIZ_VERI: değişen hücreler turuncu, eski değer hücre notunda (fareyle üzerine gelin). |
| X_OZET.xlsx | OZET: toplam uyarı/düzeltilen/geçersiz sayısı. DETAY: satır, kolon, değer, durum. |
| X_dia_hazir_SILINENLER.xlsx | Silinen mükerrer satırlar (geri almak için). |
| X_dia_hazir_MUKERRER_GRUPLAR.xlsx | Mükerrer kayıtlar grup grup. |
| musteriler/TOPLU_RAPOR_tarih.xlsx | Tüm müşteriler × modüller: satır, uyarı, düzeltilen, geçersiz. |
| X_DIA_KURU_RAPOR.xlsx / X_DIA_AKTARIM_RAPORU.xlsx | DIA gönderiminin satır satır sonucu. |
| *.bak | Üzerine yazılan dosyanın bir önceki hali. |

### KONTROL_UYARI kelimeleri ve renkler

| Durum | Anlamı | Renk |
|---|---|---|
| DÜZELTİLDİ (eski -> yeni) | Program düzeltti, bilgi amaçlı. | Turuncu |
| UYARI: … | Düzeltilmedi, kontrol edin (olası hata). | — |
| GEÇERSİZ / HATALI | DIA'ya bu haliyle gitmemeli, düzeltin. | Kırmızı |
| MÜKERRER / OLASI MÜKERRER | Başka satırla aynı/benzer; satır numarası yazar. | Turuncu |
| KONTROL_MUKERRER kolonu | Mükerrer kurallarına takılan satırlar. | Sarı |
| SIL_ONAY = E | Uygula'da silinecek. | Yeşil |

Uyarıdaki "satir 5" Excel'deki satır numarasıdır (1. satır başlık). Excel'de Veri > Filtre ile KONTROL_UYARI kolonunu "GEÇERSİZ" içerenlere süzerek önce kırmızıları düzeltin.

## 6. Logo'dan Çekme ve DIA'ya Gönderme

### Logo'dan Çek

1. SQL sunucu, veritabanı, kullanıcı/şifre (veya Windows oturumu) ve firma numarasını girin.
2. Bağlantı Testi: firmalar listelenir.
3. Cari Çek / Stok Çek: müşteri klasörüne cari_listesi.xlsx / stok_listesi.xlsx yazılır.
4. Sonra normal temizlik (ÇALIŞTIR). Logo verisine yazılmaz, sadece okunur.

> **Dikkat:** Bilgisayarda "ODBC Driver 17 (veya 18) for SQL Server" kurulu olmalıdır.

### ☁ DIA'ya Gönder

1. Sunucu kodu (örn. racabilisim), kullanıcı, şifre, API key, Firma Kodu, Dönem Kodu.
2. Bağlantı Testi: kalan kontör ve yetkili firma/dönem listesi gelir. Girdiğiniz firma listede yoksa "!! DIKKAT" uyarısı çıkar (401 INSUFFICIENT_PRIVILEGES hatasının sebebi budur).
3. Kuru çalıştır işaretliyken GÖNDER: hiçbir şey yazılmaz, X_DIA_KURU_RAPOR.xlsx üretilir.
4. Raporu kontrol edip işareti kaldırın ve GÖNDER. Kod DIA'da yoksa eklenir, varsa güncellenir.

Her çağrı 0,0125 kontör harcar. Kontör biterse (406) gönderim kendiliğinden durur. Şifreler hiçbir yere kaydedilmez; API key sadece "hatırla" işaretliyse kaydedilir.

### DIA alan eşlemesi (dia_eslesme.json)

DIA penceresindeki Alan Eşlemesi butonu müşteri klasöründeki dia_eslesme.json'u açar. "kart" bölümünde sol taraf DIA alanı, sağ taraf Excel kolonudur. Başında ? olan satırlar doğrulanmamıştır ve gönderilmez; DIA dokümanından kontrol edip ?'i silin.

```
"kart": {
    "carikartkodu": "carikartkodu",
    "unvan": "unvan",
    "eposta": "eposta",            <- ? silindi, artık gönderilir
    "?webadresi": "web adresi"     <- gönderilmez
},
"sabit": {"carikarttipi": "AL"},   <- her karta aynı değer
"mevcut_ise": "guncelle"           <- veya "atla"
```

> **Dikkat:** JSON yazımında virgül ve tırnaklara dikkat edin; son satırdan sonra virgül olmaz. Bozulursa dosyayı silin, program varsayılanını yeniden oluşturur.

## 7. Sık Karşılaşılan Sorunlar

| Sorun | Çözüm |
|---|---|
| "cari_listesi.xlsx bulunamadi" | Dosya adını kontrol edin veya Ayarlar > girdi_dosya'yı değiştirin. |
| "cari_dia_hazir.xlsx bulunamadi" (DIA) | İnceleme modunda hazır dosya oluşmaz; Uygula/Otomatik Sil çalıştırın. |
| Dosya yazılamadı / Permission denied | Dosya Excel'de açık; kapatıp tekrar çalıştırın. |
| VKN/telefonun başındaki 0 gitmiş | Program VKN'yi 10 haneye tamamlar. Müşteriden metin biçimli dosya isteyin. |
| "BİLİMSEL GÖSTERİM, HANE KAYBI" | Kaynak Excel'de kolonu Metin yapıp değerleri yeniden alın; kayıp geri getirilemez. |
| Yeni ayar Ayarlar'da görünmüyor | ＋ Eksik Ayarları Ekle, sonra KAYDET. |
| Ayar değiştirdim ama etkisi yok | KAYDET'e bastınız mı? Doğru MODÜL seçili mi? Kolon adı aynen mi? |
| Çok fazla OLASI MÜKERRER | benzer_kayit_esik'i yükseltin (95) veya benzer_kayit_kontrol = False. |
| DIA 401 INSUFFICIENT_PRIVILEGES | Firma kodu yanlış; Bağlantı Testi'ndeki yetkili firma listesine bakın. |
| DIA 406 | Kontör bitmiş; DIA'dan kontör yükletin. |
| Windows "bilinmeyen yayıncı" | Ek bilgi > Yine de çalıştır (imzasız EXE). |

## 8. Ayar Referansı: Tüm Modüller

Bu ayarlar her modülde ayrı ayrı bulunur (her modül için farklı değer verilebilir).

### telefon_format

Telefon / cep / faks numaralarının yazım biçimi.

|   |   |
|---|---|
| Tür | Seçim: 10 / 11 / 90 / +90 |
| Varsayılan | 10 |
| Formül / yazım | 10 = başta 0 yok \| 11 = başta 0 var \| 90 = ülke kodlu \| +90 = + ile ülke kodlu |
| Örnek | 0532 123 45 67 -> 10: 5321234567 \| 11: 05321234567 \| +90: +905321234567 |

### vkn_sifir_tamamla

Excel VKN'yi sayı olarak saklayınca baştaki 0 silinir. 8-9 haneli VKN soldan 0 ile 10 haneye tamamlanır.

|   |   |
|---|---|
| Tür | Açık/Kapalı (True/False) |
| Varsayılan | True |
| Formül / yazım | VKN 8 veya 9 haneyse -> soldan 0 ile 10 haneye tamamla |
| Örnek | 123456789 -> 0123456789 |

### vkn_algoritma_kontrol

VKN'nin 10. hanesini (kontrol hanesi) GİB algoritmasına göre doğrular; hatalıysa uyarı yazar, değeri değiştirmez.

|   |   |
|---|---|
| Tür | Açık/Kapalı (True/False) |
| Varsayılan | True |
| Formül / yazım | t=(d[i]+9-i)%10; p=(t*2^(9-i))%9 (t!=0 ve p=0 ise p=9); kontrol=(10-toplam%10)%10 |
| Örnek | 0123456789 -> OK |

### tc_algoritma_kontrol

TC Kimlik No algoritma kontrolü; hatalıysa uyarı yazar, değeri değiştirmez.

|   |   |
|---|---|
| Tür | Açık/Kapalı (True/False) |
| Varsayılan | True |
| Formül / yazım | d10=((d1+d3+d5+d7+d9)*7-(d2+d4+d6+d8))%10; d11=(d1+..+d10)%10 |
| Örnek | 10000000146 -> OK |

### il_ilce_duzelt

İl/ilçeyi 81 il, 973 ilçe listesine göre düzeltir: aynı hücredeki il/ilçeyi ayırır, yer değiştirmişse düzeltir.

|   |   |
|---|---|
| Tür | Açık/Kapalı (True/False) |
| Varsayılan | True |
| Formül / yazım | sehir='Kadıköy/İstanbul' veya sehir='Kadıköy', ilce='İstanbul' -> sehir=İSTANBUL, ilce=KADIKÖY |
| Örnek | Kadıköy / İstanbul -> İSTANBUL / KADIKÖY |

### il_bos_ise_doldur

İl boş, ilçe tek bir ile aitse ili doldurur (Gölbaşı gibi birden fazla ilde olan ilçeler doldurulmaz, uyarı verilir).

|   |   |
|---|---|
| Tür | Açık/Kapalı (True/False) |
| Varsayılan | True |
| Formül / yazım | ilçe yalnız 1 ilde varsa -> sehir = o il |
| Örnek | ilce='Siverek' -> sehir='ŞANLIURFA' |

### il_ilce_yazim

Düzeltilen il/ilçe adlarının yazımı.

|   |   |
|---|---|
| Tür | Seçim: BUYUK / BASLIK |
| Varsayılan | BUYUK |
| Formül / yazım | BUYUK = İSTANBUL \| BASLIK = İstanbul |
| Örnek | kadikoy -> KADIKÖY veya Kadıköy |

### il_sutunu

İl bilgisinin bulunduğu DIA kolonu.

|   |   |
|---|---|
| Tür | Metin |
| Varsayılan | sehir |
| Formül / yazım | kolon adı |
| Örnek | sehir |

### ilce_sutunu

İlçe bilgisinin bulunduğu DIA kolonu.

|   |   |
|---|---|
| Tür | Metin |
| Varsayılan | ilce |
| Formül / yazım | kolon adı |
| Örnek | ilce |

### hata_degerleri_temizle

Excel formül hataları ve anlamsız değerler (#YOK, #N/A, '-', '.') boş hücreye çevrilir.

|   |   |
|---|---|
| Tür | Açık/Kapalı (True/False) |
| Varsayılan | True |
| Formül / yazım | hücre hata_degerleri listesindeyse -> boş |
| Örnek | #YOK -> (boş) |

### hata_degerleri

Boş sayılacak değerler (virgülle). 0 bilerek yok: KDV 0 gibi gerçek değerler silinmesin.

|   |   |
|---|---|
| Tür | Metin |
| Varsayılan | #YOK, #N/A, #DEĞER!, #VALUE!, #REF!, #BAŞV!, #SAYI/0!, #DIV/0!, #AD?, #NAME?, #SAYI!, #NUM!, #NULL!, #BOŞ!, -, --, ., ?, N/A, NULL |
| Formül / yazım | değer1, değer2, ... |
| Örnek | #YOK, #N/A, - |

### gorunmez_karakter_temizle

Sekme, satır sonu, Excel'in _x000D_ kalıntısı, sıfır genişlikli ve bölünmez boşluklar temizlenir.

|   |   |
|---|---|
| Tür | Açık/Kapalı (True/False) |
| Varsayılan | True |
| Formül / yazım | \t \r \n _x000D_ NBSP -> bosluk \| ZWSP BOM -> silinir |
| Örnek | ABC_x000D_ LTD -> ABC LTD |

### bozuk_turkce_duzelt

Yanlış kod sayfasıyla açılmış dosyalardaki bozuk Türkçe harfler düzeltilir.

|   |   |
|---|---|
| Tür | Açık/Kapalı (True/False) |
| Varsayılan | True |
| Formül / yazım | Ã‡->Ç Ä°->İ Åž->Ş Ã¼->ü ... \| Ý->İ Þ->Ş Ð->Ğ ý->ı þ->ş ð->ğ |
| Örnek | Ä°STANBUL -> İSTANBUL |

### bilimsel_gosterim_kontrol

Kod/barkod/telefon kolonlarında 1,23457E+11 gibi bilimsel gösterim kurtarılır; hane kaybı varsa GEÇERSİZ işaretlenir.

|   |   |
|---|---|
| Tür | Açık/Kapalı (True/False) |
| Varsayılan | True |
| Formül / yazım | hane sayısı yeterliyse sayıya çevir, değilse uyarı |
| Örnek | 8,69E+12 -> GEÇERSİZ (veri kaybı) \| 8.691234567890E+12 -> 8691234567890 |

### degisiklik_raporu

Her hücre için 'eski değer -> yeni değer' listesi (_DEGISIKLIKLER.xlsx). İkinci sayfada temiz veri; değişen hücreler turuncu, eski değer hücre notunda.

|   |   |
|---|---|
| Tür | Açık/Kapalı (True/False) |
| Varsayılan | True |
| Formül / yazım | okunan veri ile temizlenmiş veri hücre hücre karşılaştırılır |
| Örnek | satir 5 \| telefon1 \| 0532 123 45 67 \| 5321234567 |

### zorunlu_alanlar

DIA'ya gitmeden önce boş olmaması gereken kolonlar. Boşsa GEÇERSİZ işaretlenir.

|   |   |
|---|---|
| Tür | Metin |
| Varsayılan | cari: carikartkodu, unvan; stok: stok kart kodu, stok adı; banka: Hesap Kodu, Hesap Adı; personel: Adı, Soyadı, TC Kimlik No; demirbas: DEMİRBAŞ.KODU, DEMİRBAŞ ADI; hesapplan: Hesap Kodu, Hesap Adı  (diğer modüller: (boş)) |
| Formül / yazım | kolon1, kolon2 \| kapatmak için YOK |
| Örnek | carikartkodu, unvan |

### alan_uzunluklari

Kolonların azami karakter sayısı. Uzun olan GEÇERSİZ işaretlenir (DIA'da kesilmesin / hata vermesin). Değerler tahminidir, DIA ekranına göre düzeltin.

|   |   |
|---|---|
| Tür | Metin |
| Varsayılan | cari: carikartkodu=25; unvan=200; vergidairesi=50; eposta=100; adres1=200; adres2=200; stok: stok kart kodu=25; stok adı=200; Barkod No=50; banka: Hesap Kodu=25; Hesap Adı=100; personel: Adı=50; Soyadı=50; Sicil=25; demirbas: DEMİRBAŞ.KODU=25; DEMİRBAŞ ADI=200; hesapplan: Hesap Kodu=50; Hesap Adı=200  (diğer modüller: (boş)) |
| Formül / yazım | kolon=uzunluk; kolon=uzunluk |
| Örnek | carikartkodu=25; unvan=200 |

### esleme_hafizasi

Kurallar'a kaydedilen sütun eşlemeleri %APPDATA%\RACA_DIA\esleme_hafizasi.json'a da yazılır; başka müşteride aynı başlık gelirse sormadan uygulanır.

|   |   |
|---|---|
| Tür | Açık/Kapalı (True/False) |
| Varsayılan | True |
| Formül / yazım | müşterinin kendi eşlemesi > hafıza (müşterininki önceliklidir) |
| Örnek | 'Firma Ünvanı' -> unvan |

### capraz_kontrol

Bu modüldeki kodlar diğer modülün dosyasında var mı? (aynı müşteri klasörü). Yoksa uyarı verir.

|   |   |
|---|---|
| Tür | Metin |
| Varsayılan | stok: Grup Kodu>urunagaci:Grup1,GRUP2,GRUP3,GRUP4,GRUP5,GRUP6; muhkod: GRUP KODU>urunagaci:Grup1,GRUP2,GRUP3,GRUP4,GRUP5,GRUP6  (diğer modüller: (boş)) |
| Formül / yazım | kolon>modul:kolon1,kolon2; ... |
| Örnek | Grup Kodu>urunagaci:Grup1,GRUP2 |

## 8. Ayar Referansı: Cari

### benzer_kayit_kontrol

Yazımı farklı ama aynı olabilecek kayıtları (ABC LTD ŞTİ / ABC Limited Şirketi) benzerlik puanıyla işaretler. Silmez.

|   |   |
|---|---|
| Tür | Açık/Kapalı (True/False) |
| Varsayılan | True |
| Formül / yazım | LTD/ŞTİ/AŞ/TİC gibi ekler atılır, kalan metinlerin benzerliği >= eşik ise işaretlenir |
| Örnek | 'ABC Gıda Ltd Şti' ~ 'ABC GIDA LİMİTED ŞİRKETİ' %100 |

### benzer_kayit_sutunu

Benzerliği karşılaştırılacak kolon(lar). Birden fazlaysa virgülle; birleştirilerek karşılaştırılır.

|   |   |
|---|---|
| Tür | Metin |
| Varsayılan | unvan |
| Formül / yazım | kolon1, kolon2 |
| Örnek | Adı, Soyadı |

### benzer_kayit_esik

Benzerlik eşiği (yüzde). Düşürülürse daha çok, yükseltilirse daha az kayıt işaretlenir.

|   |   |
|---|---|
| Tür | Metin |
| Varsayılan | 88 |
| Formül / yazım | 0-100 |
| Örnek | 88 |

### unvan_standart

Ünvandaki şirket türü kısaltmalarını tek biçime getirir.

|   |   |
|---|---|
| Tür | Açık/Kapalı (True/False) |
| Varsayılan | True |
| Formül / yazım | unvan_kisaltmalari listesine göre |
| Örnek | ABC Ltd.Şti -> ABC LTD. ŞTİ. |

### unvan_kisaltmalari

Aranan=Yazılacak çiftleri. Arama Türkçe harf / nokta / büyük-küçük harf duyarsızdır. Sonuna $ konan ifade sadece ünvanın sonunda aranır.

|   |   |
|---|---|
| Tür | Metin |
| Varsayılan | LIMITED SIRKETI=LTD. ŞTİ.; LIMITED STI=LTD. ŞTİ.; LTD SIRKETI=LTD. ŞTİ.; LTD STI=LTD. ŞTİ.; ANONIM SIRKETI=A.Ş.; ANONIM STI=A.Ş.; A S=A.Ş.; AS$=A.Ş.; TICARET=TİC.; TIC=TİC.; SANAYI=SAN.; SAN=SAN. |
| Formül / yazım | ARANAN=YAZILACAK; ... |
| Örnek | LIMITED SIRKETI=LTD. ŞTİ. |

### vergi_dairesi_duzelt

Vergi dairesinden 'VD', 'V.D.', 'Vergi Dairesi Müdürlüğü' ekleri atılır; il/ilçe adı olanların yazımı düzeltilir, büyük harf yazılır.

|   |   |
|---|---|
| Tür | Açık/Kapalı (True/False) |
| Varsayılan | True |
| Formül / yazım | ekleri sil -> il/ilçe listesinde ara -> vergi_dairesi_eslesme'de ara |
| Örnek | Kadikoy V.D. -> KADIKÖY |

### vergi_dairesi_eslesme

İl/ilçe adı olmayan vergi daireleri için kendi düzeltmeleriniz.

|   |   |
|---|---|
| Tür | Metin |
| Varsayılan | (boş) |
| Formül / yazım | ARANAN=YAZILACAK; ... |
| Örnek | KOZYATAGI=KOZYATAĞI; BUYUK MUKELLEFLER=BÜYÜK MÜKELLEFLER |

### vkn_tc_tutarlilik

Ünvanı şirket (LTD/A.Ş.) olduğu halde sadece TC Kimlik No girilmiş kayıtları işaretler.

|   |   |
|---|---|
| Tür | Açık/Kapalı (True/False) |
| Varsayılan | True |
| Formül / yazım | ünvanda LTD/ŞTİ/AŞ var, VKN boş, TC dolu -> uyarı |
| Örnek | ABC LTD STI + TC -> UYARI |

### ad_soyad_ayir

Şahıs carilerinde ünvandan ad ve soyad kolonlarını doldurur (e-Fatura). DIA aktarım şablonunuzda bu kolonlar varsa açın.

|   |   |
|---|---|
| Tür | Açık/Kapalı (True/False) |
| Varsayılan | False |
| Formül / yazım | son kelime = soyad, öncekiler = ad |
| Örnek | Ahmet Ali Yılmaz -> Ahmet Ali \| Yılmaz |

### ad_sutunu

Ad kolonunun adı.

|   |   |
|---|---|
| Tür | Metin |
| Varsayılan | adi |
| Formül / yazım | kolon adı |
| Örnek | adi |

### soyad_sutunu

Soyad kolonunun adı.

|   |   |
|---|---|
| Tür | Metin |
| Varsayılan | soyadi |
| Formül / yazım | kolon adı |
| Örnek | soyadi |

### adresten_il_ilce

İl/ilçe boşsa adres1/adres2 içinde geçen il ve o ile ait ilçe ile doldurur. 'Ankara Cad.' gibi cadde/sokak adları dikkate alınmaz.

|   |   |
|---|---|
| Tür | Açık/Kapalı (True/False) |
| Varsayılan | True |
| Formül / yazım | adreste il geçiyor ve ilçe o ile aitse -> doldur |
| Örnek | 'Moda Cad. 5 Kadıköy/İstanbul' -> İSTANBUL / KADIKÖY |

### adres_kisaltma

Adreslerdeki Mahallesi/Caddesi/Sokak kelimelerini kısaltır (adres uzunluğu sınırına takılmamak için).

|   |   |
|---|---|
| Tür | Açık/Kapalı (True/False) |
| Varsayılan | False |
| Formül / yazım | adres_kisaltmalari listesine göre |
| Örnek | Moda Mahallesi Bahariye Caddesi -> Moda MAH. Bahariye CAD. |

### adres_kisaltmalari

Aranan=Yazılacak çiftleri.

|   |   |
|---|---|
| Tür | Metin |
| Varsayılan | MAHALLESI=MAH.; MAHALLE=MAH.; MH=MAH.; CADDESI=CAD.; CADDE=CAD.; CD=CAD.; SOKAGI=SK.; SOKAK=SK.; SOK=SK.; BULVARI=BLV.; BULVAR=BLV.; NUMARA=NO |
| Formül / yazım | ARANAN=YAZILACAK; ... |
| Örnek | CADDESI=CAD. |

### cep_ayir

Cep kolonu boşsa, telefon kolonlarındaki ilk cep numarası (5xx) cep kolonuna taşınır.

|   |   |
|---|---|
| Tür | Açık/Kapalı (True/False) |
| Varsayılan | True |
| Formül / yazım | telefon 5 ile başlıyor ve cep boşsa -> cep kolonuna taşı |
| Örnek | telefon1=5321234567 -> ceptel |

### cep_sutunu

Cep telefonu kolonunun adı.

|   |   |
|---|---|
| Tür | Metin |
| Varsayılan | ceptel |
| Formül / yazım | kolon adı |
| Örnek | ceptel |

### posta_kodu_il_kontrol

Posta kodunun ilk 2 hanesi ilin plaka koduyla uyuşmuyorsa uyarı verir.

|   |   |
|---|---|
| Tür | Açık/Kapalı (True/False) |
| Varsayılan | True |
| Formül / yazım | posta kodunun ilk 2 hanesi = ilin plaka kodu |
| Örnek | 34710 + ANKARA -> UYARI |

## 8. Ayar Referansı: Stok

### benzer_kayit_kontrol

Yazımı farklı ama aynı olabilecek kayıtları (ABC LTD ŞTİ / ABC Limited Şirketi) benzerlik puanıyla işaretler. Silmez.

|   |   |
|---|---|
| Tür | Açık/Kapalı (True/False) |
| Varsayılan | True |
| Formül / yazım | LTD/ŞTİ/AŞ/TİC gibi ekler atılır, kalan metinlerin benzerliği >= eşik ise işaretlenir |
| Örnek | 'ABC Gıda Ltd Şti' ~ 'ABC GIDA LİMİTED ŞİRKETİ' %100 |

### benzer_kayit_sutunu

Benzerliği karşılaştırılacak kolon(lar). Birden fazlaysa virgülle; birleştirilerek karşılaştırılır.

|   |   |
|---|---|
| Tür | Metin |
| Varsayılan | stok adı |
| Formül / yazım | kolon1, kolon2 |
| Örnek | Adı, Soyadı |

### benzer_kayit_esik

Benzerlik eşiği (yüzde). Düşürülürse daha çok, yükseltilirse daha az kayıt işaretlenir.

|   |   |
|---|---|
| Tür | Metin |
| Varsayılan | 88 |
| Formül / yazım | 0-100 |
| Örnek | 88 |

### barkod_kontrol

EAN-8 / UPC-A / EAN-13 / GTIN-14 barkodların kontrol hanesini doğrular. Diğer uzunluklar (iç barkod) atlanır.

|   |   |
|---|---|
| Tür | Açık/Kapalı (True/False) |
| Varsayılan | True |
| Formül / yazım | GTIN mod-10 kontrol hanesi |
| Örnek | 8690000000001 -> GEÇERSİZ |

### barkod_sutunu

Barkod kolonunun adı.

|   |   |
|---|---|
| Tür | Metin |
| Varsayılan | Barkod No |
| Formül / yazım | kolon adı |
| Örnek | Barkod No |

### birim_standart

Birimleri tek biçime getirir; listede olmayan birim için uyarı verir.

|   |   |
|---|---|
| Tür | Açık/Kapalı (True/False) |
| Varsayılan | True |
| Formül / yazım | birim_eslesme listesine göre |
| Örnek | Adt -> ADET, Kilo -> KG |

### birim_sutunu

Birim kolon(lar)ı (virgülle).

|   |   |
|---|---|
| Tür | Metin |
| Varsayılan | Ana Birim |
| Formül / yazım | kolon1, kolon2 |
| Örnek | Ana Birim |

### birim_eslesme

Aranan=Yazılacak birim çiftleri (Türkçe harf / nokta / boşluk duyarsız).

|   |   |
|---|---|
| Tür | Metin |
| Varsayılan | AD=ADET; ADT=ADET; ADET=ADET; PCS=ADET; PC=ADET; PIECE=ADET; ST=ADET; KG=KG; KILO=KG; KILOGRAM=KG; KGM=KG; GR=GR; G=GR; GRAM=GR; LT=LT; L=LT; LITRE=LT; LTR=LT; ML=ML; MILILITRE=ML; MT=MT; M=MT; METRE=MT; MTR=MT; CM=CM; MM=MM; M2=M2; METREKARE=M2; M3=M3; METREKUP=M3; KOLI=KOLİ; KL=KOLİ; PAKET=PAKET; PK=PAKET; PKT=PAKET; KUTU=KUTU; KT=KUTU; TAKIM=TAKIM; TK=TAKIM; SET=TAKIM; CIFT=ÇİFT; DUZINE=DÜZİNE; DZ=DÜZİNE; TON=TON; TN=TON; RULO=RULO; TOP=TOP; SAAT=SAAT; GUN=GÜN |
| Formül / yazım | ARANAN=YAZILACAK; ... |
| Örnek | ADT=ADET; KILO=KG |

### kdv_eski_oranlar

Artık kullanılmayan KDV oranları; bu oranlar için uyarı verilir. Kapatmak için YOK.

|   |   |
|---|---|
| Tür | Metin |
| Varsayılan | 8, 18 |
| Formül / yazım | oran1, oran2 |
| Örnek | %18 -> UYARI: ESKİ KDV ORANI |

### fiyat_temizle

Metin fiyatları sayıya çevirir, para birimini ayırır (TL dışındaysa ilgili Döviz kolonuna yazar).

|   |   |
|---|---|
| Tür | Açık/Kapalı (True/False) |
| Varsayılan | True |
| Formül / yazım | '1.250,50 TL' -> 1250.5 \| '12,5 USD' -> 12.5 + Döviz=USD |
| Örnek | 1.250,50 TL -> 1250.5 |

### alis_satis_kontrol

Alış fiyatı satış fiyatından büyükse uyarı verir.

|   |   |
|---|---|
| Tür | Açık/Kapalı (True/False) |
| Varsayılan | True |
| Formül / yazım | alis_fiyat_sutunu > satis_fiyat_sutunu -> uyarı |
| Örnek | alış 120, satış 100 -> UYARI |

### satis_fiyat_sutunu

Satış fiyatı kolonu.

|   |   |
|---|---|
| Tür | Metin |
| Varsayılan | Fiyat 1 |
| Formül / yazım | kolon adı |
| Örnek | Fiyat 1 |

### alis_fiyat_sutunu

Alış fiyatı kolonu.

|   |   |
|---|---|
| Tür | Metin |
| Varsayılan | Fiyat 2 |
| Formül / yazım | kolon adı |
| Örnek | Fiyat 2 |

### stok_adi_birim_duzelt

Stok adındaki ölçü yazımını düzeltir: sayı ile birim arasına boşluk, birim büyük harf.

|   |   |
|---|---|
| Tür | Açık/Kapalı (True/False) |
| Varsayılan | True |
| Formül / yazım | 500gr / 500 Gr -> 500 GR \| 1l -> 1 LT |
| Örnek | Süt 1l -> Süt 1 LT |

### stok_adi_yazim

Stok adının harf yazımı.

|   |   |
|---|---|
| Tür | Seçim: DOKUNMA / BUYUK / BASLIK |
| Varsayılan | DOKUNMA |
| Formül / yazım | DOKUNMA \| BUYUK = SÜT \| BASLIK = Süt |
| Örnek | süt tam yağlı -> SÜT TAM YAĞLI |

## 8. Ayar Referansı: Banka

### benzer_kayit_kontrol

Yazımı farklı ama aynı olabilecek kayıtları (ABC LTD ŞTİ / ABC Limited Şirketi) benzerlik puanıyla işaretler. Silmez.

|   |   |
|---|---|
| Tür | Açık/Kapalı (True/False) |
| Varsayılan | True |
| Formül / yazım | LTD/ŞTİ/AŞ/TİC gibi ekler atılır, kalan metinlerin benzerliği >= eşik ise işaretlenir |
| Örnek | 'ABC Gıda Ltd Şti' ~ 'ABC GIDA LİMİTED ŞİRKETİ' %100 |

### benzer_kayit_sutunu

Benzerliği karşılaştırılacak kolon(lar). Birden fazlaysa virgülle; birleştirilerek karşılaştırılır.

|   |   |
|---|---|
| Tür | Metin |
| Varsayılan | Hesap Adı |
| Formül / yazım | kolon1, kolon2 |
| Örnek | Adı, Soyadı |

### benzer_kayit_esik

Benzerlik eşiği (yüzde). Düşürülürse daha çok, yükseltilirse daha az kayıt işaretlenir.

|   |   |
|---|---|
| Tür | Metin |
| Varsayılan | 88 |
| Formül / yazım | 0-100 |
| Örnek | 88 |

### iban_banka_adi

IBAN'daki banka kodundan banka adını doldurur; dolu banka adı IBAN'la uyuşmuyorsa uyarı verir.

|   |   |
|---|---|
| Tür | Açık/Kapalı (True/False) |
| Varsayılan | True |
| Formül / yazım | IBAN'ın 5-9. haneleri (banka kodu) -> banka adı |
| Örnek | TR..00062.. -> GARANTİ BBVA |

### banka_adi_standart

Banka adını tek biçime getirir.

|   |   |
|---|---|
| Tür | Açık/Kapalı (True/False) |
| Varsayılan | True |
| Formül / yazım | bilinen bankalar listesi |
| Örnek | T.C. Ziraat Bankası A.Ş. -> T.C. ZİRAAT BANKASI |

### iban_sutunu

IBAN kolonu.

|   |   |
|---|---|
| Tür | Metin |
| Varsayılan | IBAN No |
| Formül / yazım | kolon adı |
| Örnek | IBAN No |

### banka_adi_sutunu

Banka adı kolonu.

|   |   |
|---|---|
| Tür | Metin |
| Varsayılan | Banka adı |
| Formül / yazım | kolon adı |
| Örnek | Banka adı |

### iban_hesap_no_doldur

Hesap No boşsa IBAN'ın son 16 hanesinden (baştaki 0'lar atılarak) doldurur.

|   |   |
|---|---|
| Tür | Açık/Kapalı (True/False) |
| Varsayılan | True |
| Formül / yazım | TR IBAN'ın son 16 hanesi, baştaki 0'lar atılarak |
| Örnek | TR12 0006 4000 0011 2345 6789 01 -> 112345678901 |

### hesap_no_sutunu

Hesap numarası kolonu.

|   |   |
|---|---|
| Tür | Metin |
| Varsayılan | Hesap NO |
| Formül / yazım | kolon adı |
| Örnek | Hesap NO |

### doviz_hesap_adi_kontrol

Hesap adında USD/EURO gibi para birimi geçiyor ama Döviz Türü farklıysa uyarı verir.

|   |   |
|---|---|
| Tür | Açık/Kapalı (True/False) |
| Varsayılan | True |
| Formül / yazım | 'Garanti USD' + Döviz=TL -> UYARI |
| Örnek | Garanti USD / TL -> UYARI |

## 8. Ayar Referansı: Personel

### benzer_kayit_kontrol

Yazımı farklı ama aynı olabilecek kayıtları (ABC LTD ŞTİ / ABC Limited Şirketi) benzerlik puanıyla işaretler. Silmez.

|   |   |
|---|---|
| Tür | Açık/Kapalı (True/False) |
| Varsayılan | True |
| Formül / yazım | LTD/ŞTİ/AŞ/TİC gibi ekler atılır, kalan metinlerin benzerliği >= eşik ise işaretlenir |
| Örnek | 'ABC Gıda Ltd Şti' ~ 'ABC GIDA LİMİTED ŞİRKETİ' %100 |

### benzer_kayit_sutunu

Benzerliği karşılaştırılacak kolon(lar). Birden fazlaysa virgülle; birleştirilerek karşılaştırılır.

|   |   |
|---|---|
| Tür | Metin |
| Varsayılan | Adı, Soyadı |
| Formül / yazım | kolon1, kolon2 |
| Örnek | Adı, Soyadı |

### benzer_kayit_esik

Benzerlik eşiği (yüzde). Düşürülürse daha çok, yükseltilirse daha az kayıt işaretlenir.

|   |   |
|---|---|
| Tür | Metin |
| Varsayılan | 88 |
| Formül / yazım | 0-100 |
| Örnek | 88 |

### posta_kodu_il_kontrol

Posta kodunun ilk 2 hanesi ilin plaka koduyla uyuşmuyorsa uyarı verir.

|   |   |
|---|---|
| Tür | Açık/Kapalı (True/False) |
| Varsayılan | True |
| Formül / yazım | posta kodunun ilk 2 hanesi = ilin plaka kodu |
| Örnek | 34710 + ANKARA -> UYARI |

### iban_banka_adi

IBAN'daki banka kodundan banka adını doldurur; dolu banka adı IBAN'la uyuşmuyorsa uyarı verir.

|   |   |
|---|---|
| Tür | Açık/Kapalı (True/False) |
| Varsayılan | True |
| Formül / yazım | IBAN'ın 5-9. haneleri (banka kodu) -> banka adı |
| Örnek | TR..00062.. -> GARANTİ BBVA |

### banka_adi_standart

Banka adını tek biçime getirir.

|   |   |
|---|---|
| Tür | Açık/Kapalı (True/False) |
| Varsayılan | True |
| Formül / yazım | bilinen bankalar listesi |
| Örnek | T.C. Ziraat Bankası A.Ş. -> T.C. ZİRAAT BANKASI |

### iban_sutunu

IBAN kolonu.

|   |   |
|---|---|
| Tür | Metin |
| Varsayılan | IBAN |
| Formül / yazım | kolon adı |
| Örnek | IBAN No |

### banka_adi_sutunu

Banka adı kolonu.

|   |   |
|---|---|
| Tür | Metin |
| Varsayılan | Banka Adı |
| Formül / yazım | kolon adı |
| Örnek | Banka adı |

### personel_standart

Cinsiyet, medeni hal ve kan grubu yazımını standartlaştırır (kolon varsa).

|   |   |
|---|---|
| Tür | Açık/Kapalı (True/False) |
| Varsayılan | True |
| Formül / yazım | E/Bay/Erkek -> ERKEK \| B/Bekar -> BEKAR \| a rh pozitif -> A RH+ |
| Örnek | Bayan -> KADIN, 0 Rh(-) -> 0 RH- |

### yas_min

İşe başlama tarihindeki yaş bundan küçükse uyarı verir.

|   |   |
|---|---|
| Tür | Metin |
| Varsayılan | 15 |
| Formül / yazım | yıl |
| Örnek | 15 |

### yas_max

Bugünkü yaş bundan büyükse uyarı verir (doğum tarihi hatalı olabilir).

|   |   |
|---|---|
| Tür | Metin |
| Varsayılan | 80 |
| Formül / yazım | yıl |
| Örnek | 80 |

## 8. Ayar Referansı: Demirbaş

### benzer_kayit_kontrol

Yazımı farklı ama aynı olabilecek kayıtları (ABC LTD ŞTİ / ABC Limited Şirketi) benzerlik puanıyla işaretler. Silmez.

|   |   |
|---|---|
| Tür | Açık/Kapalı (True/False) |
| Varsayılan | True |
| Formül / yazım | LTD/ŞTİ/AŞ/TİC gibi ekler atılır, kalan metinlerin benzerliği >= eşik ise işaretlenir |
| Örnek | 'ABC Gıda Ltd Şti' ~ 'ABC GIDA LİMİTED ŞİRKETİ' %100 |

### benzer_kayit_sutunu

Benzerliği karşılaştırılacak kolon(lar). Birden fazlaysa virgülle; birleştirilerek karşılaştırılır.

|   |   |
|---|---|
| Tür | Metin |
| Varsayılan | DEMİRBAŞ ADI |
| Formül / yazım | kolon1, kolon2 |
| Örnek | Adı, Soyadı |

### benzer_kayit_esik

Benzerlik eşiği (yüzde). Düşürülürse daha çok, yükseltilirse daha az kayıt işaretlenir.

|   |   |
|---|---|
| Tür | Metin |
| Varsayılan | 88 |
| Formül / yazım | 0-100 |
| Örnek | 88 |

### demirbas_kontrol

Alış tarihi gelecekte mi, tutar 0/negatif mi, amortisman oranı süresiyle uyumlu mu (normal: 100/yıl, azalan: 2 katı, en fazla %50), birikmiş amortisman tutardan büyük mü kontrol eder.

|   |   |
|---|---|
| Tür | Açık/Kapalı (True/False) |
| Varsayılan | True |
| Formül / yazım | oran ≈ 100 / yıl (normal) veya en fazla 50 olmak üzere 200 / yıl (azalan) |
| Örnek | 5 yıl + %25 -> UYARI (normal %20, azalan %40 olmalı) |

### birikmis_amortisman_sutunu

Birikmiş amortisman kolonu (yoksa bu kontrol atlanır).

|   |   |
|---|---|
| Tür | Metin |
| Varsayılan | BİRİKMİŞ AMORTİSMAN |
| Formül / yazım | kolon adı |
| Örnek | BİRİKMİŞ AMORTİSMAN |

## 8. Ayar Referansı: Hesap Planı

### benzer_kayit_kontrol

Yazımı farklı ama aynı olabilecek kayıtları (ABC LTD ŞTİ / ABC Limited Şirketi) benzerlik puanıyla işaretler. Silmez.

|   |   |
|---|---|
| Tür | Açık/Kapalı (True/False) |
| Varsayılan | True |
| Formül / yazım | LTD/ŞTİ/AŞ/TİC gibi ekler atılır, kalan metinlerin benzerliği >= eşik ise işaretlenir |
| Örnek | 'ABC Gıda Ltd Şti' ~ 'ABC GIDA LİMİTED ŞİRKETİ' %100 |

### benzer_kayit_sutunu

Benzerliği karşılaştırılacak kolon(lar). Birden fazlaysa virgülle; birleştirilerek karşılaştırılır.

|   |   |
|---|---|
| Tür | Metin |
| Varsayılan | Hesap Adı |
| Formül / yazım | kolon1, kolon2 |
| Örnek | Adı, Soyadı |

### benzer_kayit_esik

Benzerlik eşiği (yüzde). Düşürülürse daha çok, yükseltilirse daha az kayıt işaretlenir.

|   |   |
|---|---|
| Tür | Metin |
| Varsayılan | 88 |
| Formül / yazım | 0-100 |
| Örnek | 88 |

### tdhp_kontrol

Ana hesap kodunun (ilk 3 hane) Tek Düzen Hesap Planında olup olmadığını kontrol eder. 9xx nazım hesaplar serbesttir.

|   |   |
|---|---|
| Tür | Açık/Kapalı (True/False) |
| Varsayılan | True |
| Formül / yazım | hesap kodunun ilk 3 hanesi TDHP'de olmalı |
| Örnek | 125.01 -> UYARI |

### ek_ana_hesaplar

TDHP dışında kullandığınız ana hesaplar (virgülle).

|   |   |
|---|---|
| Tür | Metin |
| Varsayılan | (boş) |
| Formül / yazım | kod1, kod2 |
| Örnek | 125, 175 |

### hesap_ayrac

Hesap kodu seviye ayracı. '-', '_', '/', ',' ve boşluk bu ayraca çevrilir.

|   |   |
|---|---|
| Tür | Metin |
| Varsayılan | . |
| Formül / yazım | 120-01-001 -> 120.01.001 |
| Örnek | . |

### hesap_seviye_uzunluklari

Seviye hane sayıları. Doluysa ayraçsız kodlar bölünür ve uzunluğu uymayan seviyeler için uyarı verilir.

|   |   |
|---|---|
| Tür | Metin |
| Varsayılan | (boş) |
| Formül / yazım | 3,2,3 -> 12001001 = 120.01.001 |
| Örnek | 3,2,3 |

### eksik_ust_hesap_ekle

120.01.001 var ama 120.01 yoksa, 120.01 satırı '(OTOMATIK) 120.01' adıyla dosya sonuna eklenir.

|   |   |
|---|---|
| Tür | Açık/Kapalı (True/False) |
| Varsayılan | False |
| Formül / yazım | her kodun üst seviyeleri yoksa ekle |
| Örnek | 120.01.001 -> + 120.01, + 120 |

## 8. Ayar Referansı: Ürün Ağacı

### grup_agaci_tutarlilik

Aynı grup kodunun farklı üst gruplar altında kullanılmasını işaretler. Grup kodlarınız her üst grupta 01, 02 diye yeniden başlıyorsa kapalı tutun.

|   |   |
|---|---|
| Tür | Açık/Kapalı (True/False) |
| Varsayılan | False |
| Formül / yazım | her grup kodunun tek bir üst grubu olmalı |
| Örnek | GRUP2=10 hem A hem B altında -> UYARI |

## 8. Ayar Referansı: Dosya ve Kolon Ayarları

Her modülde bulunan temel ayarlar.

| Ayar | Ne yapar | Örnek değer |
|---|---|---|
| girdi_dosya | Okunacak müşteri dosyası | cari_listesi.xlsx |
| cikti_dosya | DIA'ya yüklenecek temiz dosya | cari_dia_hazir.xlsx |
| sayfa | Okunacak sayfa(lar). Boşsa ilk sayfa; birden fazlaysa virgülle, birleştirilir | Sayfa1, Sayfa2 |
| telefon_alanlari | Telefon olarak işlenecek kolonlar | telefon1, telefon2, ceptel, fax |
| buyuk_harf_alanlari | BÜYÜK harfe çevrilecek kolonlar (Türkçe i/İ doğru çevrilir) | sehir, ilce |
| kucuk_harf_alanlari | küçük harfe çevrilecek kolonlar | eposta |
| kod_alanlari | Kod kolonları: Türkçe harf → ASCII, boşluk → nokta, büyük harf | carikartkodu, grupkodu |
| kdv_alanlari | KDV oranı kolonları (%18, 0,18, 18 → 18) | KDV Satış Oran |
| tarih_alanlari | Tarih kolonları (gg.aa.yyyy yazılır) | İşe Başlama Tarihi |
| mail_kontrol | E-posta biçim kontrolü ve küçük harf | True |
| iban_temizle | IBAN boşluk temizliği ve MOD-97 kontrolü | True |
| doviz_standart | Döviz adları: Dolar → USD, Avro → EUR, TRY → TL | True |
| tc_kimlik_11_hane | TC Kimlik No 11 hane kontrolü | True |

## 9. Teklif Motoru

DIA lisans + hizmet teklifi hazırlar; RACA_DIA_Teklif_Motoru.xlsx'in program halidir. Ana ekranda Teklif Motoru butonuyla veya tek başına RACA_Teklif.exe ile açılır (aynı veri kullanılır).

### İlk açılış

1. Program "paket listesi yok" der; Evet deyip mevcut teklif Excel'inizi seçin.
2. PaketListesi, ADAM SAAT, AYARLAR, KAMPANYALAR, RAKİP ve TEKLİF GEÇMİŞİ sayfaları içe aktarılır.
3. Fiyatlar programın içine değil, sizin bilgisayarınızdaki veri dosyasına yazılır.

### Teklif hazırlama (sekmeler)

1. Müşteri: ad, kullanıcı sayıları, yetkili / telefon / e-posta / VKN. Paket tipi eşiğe göre seçilir. Etiket (Ekonomik / Önerilen / Kapsamlı) alternatif tekliflerde kullanılır.
2. 1) Paketler: + Paket Ekle ile arayın, ekleyin. Adet, İnd. %, İnd. Tutar, A/S kolonlarına çift tıklayıp değiştirin. Sabit kalemler, kampanya, maliyet ve bu teklife özel saat ücreti de burada.
3. Kodla hızlı ekleme: "Kod:" kutusuna paket kodunu yazıp/yapıştırıp Enter. Birden çok kodu Excel'den veya bir listeden kopyalayıp "Panodan Yapıştır" (ya da paket listesinde Ctrl+V) ile tek seferde ekleyin. Alt alta, virgül/noktalı virgül/boşluk ile ayrılmış kodlar kabul edilir; büyük/küçük harf, Türkçe karakter ve - _ . farkı önemsizdir. "KOD 3" veya Excel'de kodun yanındaki hücre tam sayıysa adet olarak alınır. Bulunamayan, pasif ve teklifte zaten olan kodlar listelenir, eklenmez.
4. Uyarlama bedeli: ekranın altındaki turuncu bantta (tıklayınca Keşif & Uyarlama dökümü açılır) ve PDF/Excel'de ayrı bölümde gösterilir. PDF'te bölüm başında açıklama metni (Ayarlar: uyarlama_aciklama), her hizmet kaleminin kapsamı, dayanağı (ör. "7 kullanıcı → 1 grup × 8 saat") ve hesabı (saat × saat ücreti), süre saat ve iş günü olarak (Ayarlar: is_gunu_saat), sonunda da vurgulu "UYARLAMA HİZMET BEDELİ" kutusu yer alır. Elle değiştirilen saat ve kalibrasyon gerekçeleri müşteriye gösterilmez.
5. 2) Keşif & Uyarlama: soruları cevaplayın. Sağdaki hizmet dökümü her kalemin saatini ve "neye göre" hesaplandığını gösterir. Saat kolonuna çift tıklayıp elle değiştirebilirsiniz (boş = öneri). Müşteri verisinden doldur: veri temizleme klasöründen cari/stok sayısını ve hata oranını alır. Şehir dışıysa yerinde gün ve kişi sayısını girin.
6. 3) Ödeme / Taksit: her seçenekte müşterinin ödeyeceği, bankanın keseceği komisyon ve size kalan tutar. PDF kolonuna çift tıklayarak müşteriye hangi seçeneklerin sunulacağını belirleyin.
7. 4) Bakım Anlaşması: önerilen paket ve hesabın dökümü; teklife eklenecek paketi seçin.
8. 5) Satış Asistanı & Kontroller: tahmini kazanma ihtimali, öneriler ve kritik hatalar.
9. Kaydet: numara verilir. Revizyon Kaydet: aynı teklifin -R1, -R2 sürümü (eskisi korunur). PDF / Excel: kritik hata varsa yönetici şifresi (tanımlıysa) veya onay istenir.

### Uyarlama nasıl hesaplanır (hizmet dökümü)

| Hesap türü | Formül | Örnek hizmet |
|---|---|---|
| PAKET_SAAT | Paketlerin ADAM SAAT tablosundaki saatleri toplamı | Kurulum ve parametre |
| KAYIT | kayıt sayısı ÷ P1 (kayıt/saat) + P2 (hazırlık saati) | Veri aktarımı |
| KULLANICI | tavan(kullanıcı ÷ P1 kişi) × P2 saat | Kullanıcı eğitimi |
| ADET | miktar × P1 saat | Özel rapor, entegrasyon |
| SABIT | P1 saat | Canlıya geçiş desteği |
| YUZDE | diğer hizmetlerin P1 %'si | Proje yönetimi |

Sonra keşif cevapları uygulanır: EKLE (saat ekler; SAYI sorularında eşiği aşan her birim için), CARP (katsayı, * = tüm hizmetler; örn. IT personeli yoksa ×1,15), CARP_YUZDE (veri hata oranı %18 ise ×1,18), MIKTAR (kayıt/rapor/entegrasyon sayısı). En son kalibrasyon çarpanı uygulanır. Uyarlama = toplam saat × saat ücreti + lisans × lisans_yuzdesi; alt/üst sınır uygulanır. Hizmet kalemleri ve sorular Veriler > Hizmet Kalemleri / Keşif Soruları'ndan değiştirilir, yeni soru eklenebilir.

### Tahminleri gerçekle düzeltmek (kalibrasyon)

1. Proje bitince Geçmiş > teklifi seçin > Gerçekleşen Saat Gir: her hizmet için gerçek saati yazın.
2. Geçmiş > Tahmin / Gerçek Raporu: hizmet bazında sapma (örn. eğitim %+30) ve önerilen çarpan.
3. Satırı seçip "önerilen çarpanı uygula": sonraki tekliflerde o hizmetin saati otomatik düzelir.

### Kontroller

- KRİTİK: paketin kullanıcı sınırı aşıldı (açıklamada "Max 3 kullanıcı"), "modül eklenemez" paketine modül eklendi, modülün gerektirdiği ana modüller (DIA listesindeki Zorunlu kolonu) teklifte yok — eksikleri içeren paketler önerilir; toplam indirim max_toplam_indirim'i, kâr marjı min_kar_marji'nı aştı; saat ücreti taban altında veya tanımsız.
- UYARI: bir paketin modüllerinin yarıdan fazlası başka pakette zaten var (çift ücret), keşif formu boş, geçiş var ama kayıt sayısı yok, yetkili bilgisi eksik, fiyat listesi eski, seçilen kampanyanın koşulu sağlanmıyor, uyarlama/lisans oranı sınır dışında.
- Veriler > Ayarlar > yonetici_sifresi doluysa KRİTİK hatalı teklifin PDF/Excel çıktısı şifre ister (şifre açık yazılmaz, özeti saklanır).

### Satış asistanı neye göre öneri yapar

- Rakip: aynı saat × rakip saat ücreti ortalamasıyla karşılaştırır; pahalıysanız ortalamaya inmek için gereken indirimi TL ve % olarak söyler.
- En yüksek indirim: kâr marjı alt sınırını ve toplam indirim üst sınırını bozmadan verilebilecek kampanya %.
- Kampanya uygunluğu: kampanyanın Koşul alanı (kullanici>=10, soru:gecis_programi=Logo|Netsis, soru:yeni_musteri, tarih:12-15..12-31; " ve " ile birleşir; "elle" = otomatik önerilmez).
- Geçmiş: sonuçlanan tekliflerin kazanma oranı, benzer büyüklükteki (±%40) tekliflerin oranı, kazanılan ve kaybedilenlerdeki ortalama indirim, kayıp nedenlerinde "fiyat" ağırlığı.
- Ödeme: peşin indirimin müşteriye kazandırdığı tutar; kredi kartında komisyonun size kalan kısmı ve onu karşılayan vade farkı.
- Tahmini kazanma ihtimali: 50 puandan başlar; rakip konumu (±), uygun kampanya (+5), kritik hata (-10); yeterli geçmiş varsa benzer tekliflerin kazanma oranıyla harmanlanır. Kesin değil, yön göstericidir.

### Taksit, vade farkı ve kredi kartı komisyonu

| Tür | Hesap |
|---|---|
| PESIN | net × (1 - indirim %) |
| TAKSIT (havale, çek/senet) | net × (1 + vade farkı % + aylık vade % × (taksit+1)/2) |
| KREDI_KARTI | müşteri: net × (1 + vade farkı %); banka komisyonu: KDV dahil tahsilat × komisyon %; size kalan = KDV hariç tutar - komisyon |

Komisyonu tam karşılayan vade farkı = 1 / (1 - komisyon × (1 + KDV)) - 1. Örn. %4,5 komisyon, %20 KDV için yaklaşık %5,7. Bankanızın oranlarını Veriler > Ödeme Seçenekleri'ne girin.

### Bakım anlaşması neye göre belirlenir

Yıllık destek saati = taban saat + kullanıcı × destek_kullanici_saat + modül sayısı × destek_modul_saat + entegrasyon × destek_entegrasyon_saat + ek şube × destek_sube_saat + uyarlama saatinin destek_uyarlama_orani %'si; IT personeli yok gibi genel katsayılar da uygulanır. Paket fiyatı = (yıllık saat × paket çarpanı × destek saat ücreti + yerinde gün × gün ücreti) × (1 - indirim). Önerilen paket yıllık saate göre seçilir (destek_esikleri: örn. 40 saatin altı Temel, 100'ün altı Standart, üstü Premium). Paketler (saat çarpanı, yanıt süresi, yerinde gün, indirim, kapsam) Veriler > Bakım Paketleri'nden değiştirilir. DIA'nın yıllık bakımı (YBH) ayrıca hesaplanır ve 3 yıllık toplam maliyet tablosunda gösterilir.

### Alternatif teklif ve revizyon

- Aynı müşteri için Ekonomik / Önerilen / Kapsamlı teklifleri ayrı kaydedin (Etiket alanı). Geçmiş'te Ctrl ile 2-4 tanesini seçip "Seçilenleri Karşılaştır": yan yana karşılaştırma PDF'i.
- Revizyon Kaydet: RT-2026-0001 → RT-2026-0001-R1; PDF başlığında "Revizyon 1" yazar, kayıtta önceki sürümle genel toplam farkı gösterilir.

### Excel'de düzeltilen hatalar

- K24 hücresindeki bozuk formül (#NAME?) lisans neti, ara toplam, net toplam ve ödeme planını bozuyordu.
- İnd. Tutar kolonu KDV hariç toplamı ve KDV alt toplamını etkilemiyordu.
- Ödeme planı ve şablondaki NET TOPLAM, net tutar yerine kampanya indirim tutarından hesaplanıyordu.
- Adam/saat toplamı SUM(O17:O24) idi; ilk 7 paketin saati uyarlamaya girmiyordu.
- Satır indirimi birim fiyattan (toplamdan değil) hesaplanıyordu.
- Kampanya seçimi indirim yüzdesine bağlı değildi.
- ADAM SAAT satırları paket listesine satır sırasıyla bağlıydı; artık paket koduyla tutulur, DIA fiyat listesi güncellense de saatler korunur.

### Teklif geçmişi

Teklif Geçmişi penceresinde arama, Aç/Düzenle, Kopyala (aynı teklifi yeni numarayla), KAZANDI / KAYBETTİ (red nedeni sorulur) / BEKLİYOR işaretleme ve listeyi Excel'e aktarma. Altta kazanma oranı ve kazanılan tutar görünür.

### Veriler / Ayarlar

- Ayarlar, Sabit Kalemler, Kampanyalar, Rakipler ve Paketler / Adam-Saat sekmelerinde çift tık ile düzenleme; değişiklik hemen kaydedilir.
- DIA Fiyat Listesini Güncelle: DIA'nın yeni fiyat Excel'ini seçin ("Kodu" ve "Std. Fiyat" başlıkları olmalı). Fiyatı değişen, yeni gelen ve listeden kalkan paketler raporlanır; kalkanlar silinmez, pasif yapılır (eski teklifler bozulmaz). Adam/saatler korunur.

| Ayar | Varsayılan | Açıklama |
|---|---|---|
| ent_masaustu_esik | 10 | Masaüstü kullanıcı bu sayıyı geçerse paket ENTERPRISE olur |
| ent_mobil_esik | 10 | Mobil kullanıcı bu sayıyı geçerse paket ENTERPRISE olur |
| lisans_yuzdesi | 10 | Uyarlama = lisans net toplamı × bu % + adam/saat bedeli |
| saatlik_ucret | 0 | Saatlik danışman ücreti (KDV hariç) |
| min_uyarlama | 0 | Uyarlama alt sınırı (KDV hariç) |
| max_uyarlama | 0 | Uyarlama üst sınırı (0 = sınır yok) |
| kdv | 20 | KDV oranı (%) |
| rakip_ayni_saat | 1 | 1: rakip = bizim adam/saat × rakibin saatlik ücreti (aynı iş). 0: Rakipler tablosundaki saat |
| rakip_ucuz_esik | 10 | Rakip ortalamasından bu % kadar ucuzsak "Rekabetçisiniz" |
| rakip_pahali_esik | 15 | Rakip ortalamasından bu % kadar pahalıysak "Pahalısınız" |
| gecerlilik_gun | 15 | Teklifin geçerlilik süresi (gün) |
| teklif_no_onek | RT | Teklif numarası ön eki: RT-2026-0001 |
| firma_adi | RACA TEKNOLOJİ | PDF/Excel teklifin başlığındaki firma adı |
| firma_adres | (boş) | Firma adresi |
| firma_iletisim | (boş) | Telefon / e-posta / web (tek satır) |
| logo_yolu | (boş) | Logo dosyası (PNG/JPG). Boşsa logo basılmaz |
| teklif_notlari | Fiyatlar KDV hariç belirtilmiş olup KDV ayrıca gösterilmiştir. Uyarlama bedeli tahmini adam/saat üzerinden hesaplanmıştır. | Teklifin altına yazılacak şartlar / notlar |
| ybh_modul_kademe | 5:1; 10:1.1; 20:1.25; 30:1.45; *:1.65 | YBH modül sayısı katsayıları (en fazla:katsayı) |
| ybh_kullanici_kademe | 5:1; 10:1.1; 25:1.2; 50:1.35; *:1.55 | YBH kullanıcı sayısı katsayıları |
| max_toplam_indirim | 25 | Toplam indirim (satır + kampanya) liste fiyatına göre bu %'yi geçerse KRİTİK uyarı |
| min_kar_marji | 0 | Kâr marjı bu %'nin altına düşerse KRİTİK uyarı (maliyet girilen tekliflerde) |
| min_saat_ucreti | 0 | Teklife özel saatlik ücret bu tabanın altına inemez |
| uyarlama_lisans_min | 0 | Uyarlama / lisans oranı bu %'nin altındaysa "riskli düşük uyarlama" uyarısı |
| uyarlama_lisans_max | 0 | Uyarlama / lisans oranı bu %'nin üstündeyse "müşteri itiraz edebilir" uyarısı |
| yonetici_sifresi | (boş) | KRİTİK uyarılı teklifin PDF/Excel çıktısı için yönetici şifresi (boş = sadece uyarı) |
| fiyat_listesi_max_gun | 60 | Paket fiyat listesi bu günden eskiyse güncelleme uyarısı |
| fiyat_listesi_tarihi | (boş) | Son fiyat listesi içe aktarma tarihi (otomatik) |
| yol_ucreti | 0 | Şehir dışı: kişi başı gidiş-dönüş yol ücreti (ziyaret başına) |
| konaklama_ucreti | 0 | Şehir dışı: kişi başı gecelik konaklama |
| gundelik | 0 | Şehir dışı: kişi başı günlük harcırah |
| ziyaret_basina_gun | 3 | Bir ziyarette ortalama kaç gün kalınır (yol sayısı = gün / bu) |
| destek_saat_ucreti | 0 | Bakım anlaşması saat ücreti (0 = saatlik_ucret) |
| destek_taban_saat | 12 | Bakım: her müşteri için yıllık taban destek saati |
| destek_kullanici_saat | 0.5 | Bakım: kullanıcı başına yıllık ek saat |
| destek_modul_saat | 0.25 | Bakım: modül başına yıllık ek saat |
| destek_entegrasyon_saat | 4 | Bakım: entegrasyon başına yıllık ek saat |
| destek_sube_saat | 2 | Bakım: ek şube başına yıllık ek saat |
| destek_uyarlama_orani | 10 | Bakım: uyarlama saatinin bu %'si yıllık destek saatine eklenir |
| destek_yerinde_gun_ucreti | 0 | Bakım: yerinde gün ücreti (0 = 8 × destek saat ücreti) |
| destek_esikleri | 40; 100 | Önerilen bakım paketi: yıllık saat < 1. eşik Temel, < 2. eşik Standart, üstü Premium |
| tco_yil | 3 | Toplam sahip olma maliyeti tablosunun yıl sayısı |
| ybh_baslangic_yili | 2 | DIA yıllık bakım (YBH) kaçıncı yıldan başlar |
| kapsam_haric | Kapsam dışı talepler (özel rapor / ekran geliştirme, ek eğitim, ek yerinde destek vb.) saatlik ücret üzerinden ayrıca ücretlendirilir. | PDF'te uyarlama kapsamının altına yazılır |
| uyarlama_aciklama | Uyarlama hizmeti; DİA'nın firmanızın iş yapış şekline göre kurulup ayarlanması, verilerinizin aktarılması, kullanıcılarınızın eğitilmesi ve canlıya geçişte yanınızda olunması için gereken danışmanlık çalışmasıdır. Bedel, aşağıdaki iş kalemlerinin tahmini süresi ile danışman saat ücretinin çarpımı üzerinden hesaplanmıştır. | PDF'te uyarlama bölümünün başına yazılan açıklama (boş = yazılmaz) |
| is_gunu_saat | 8 | Uyarlama süresi PDF'te gün olarak da gösterilir: 1 iş günü = bu kadar saat (0 = gösterme) |

### Veri, yedek ve kayıp kayıt sorunu

Veriler EXE'nin yanına değil Belgeler\RACA_Teklif_Veri\teklif.db dosyasına yazılır. Böylece EXE İndirilenler'den, zip'in içinden çalıştırılsa ya da yeni sürüm başka klasöre konsa da kayıtlar kaybolmaz.

- Otomatik yedek: her gün ilk açılışta, teklif kaydında (en fazla 30 dakikada bir), içe aktarma ve fiyat güncellemesinden önce ve sonra, program kapanırken. Son 30 yedek Yedekler klasöründe tutulur.
- Veriler > Yedek Al: elle yedek. Yedekten Geri Yükle: seçilen yedeğe döner (mevcut hal önce yedeklenir).
- Veriler > Veri Klasörü: klasörü açar veya değiştirir (örn. satış ekibi için ortak ağ klasörü; mevcut veri oraya kopyalanır).
- Bilgisayar değiştirirken Belgeler\RACA_Teklif_Veri klasörünü yeni bilgisayara kopyalamanız yeterli.

## 10. Geliştirici: Yeni Ayar ve Kontrol Eklemek

Kaynak kod: github.com/yusufkaraca096/yusufkaraca096. Değişiklik gönderilince testler çalışır ve yeni EXE aynı indirme linkine konur.

### 1) Ayarı tanımlayın: src/ek_ayarlar.py > EK_AYAR_TANIMLARI

```
{'ayar': 'stok_kodu_on_ek', 'varsayilan': '', 'tur': 'metin', 'moduller': ['stok'],
 'aciklama': 'Stok kodlarının başına eklenecek ön ek.',
 'formul': 'on_ek + kod', 'ornek': 'ST- + 001 -> ST-001'},
```

- tur: 'bool' (True/False), 'secim' (+ 'secenekler': [...]), 'metin'.
- moduller: None = tüm modüller; liste = sadece o modüller.
- varsayilan_modul: {'cari': ..., 'stok': ...} modül bazında farklı varsayılan.
- Ayarlar penceresi, kurallar.xlsx ve bu rehber ayarı kendiliğinden gösterir.

### 2) Kodu yazın: src/ek_kontroller.py

```
def stok_on_islem(df, modul, rapor):
    on_ek = ayar(modul, 'stok_kodu_on_ek', 'stok')
    if on_ek:
        df['stok kart kodu'] = [on_ek + str(k) for k in df['stok kart kodu']]
    ...
```

- on_islem: değer düzelten işlemler (il/ilçe düzeltmesinden önce çalışır).
- son_kontroller: sadece uyarı üreten kontroller (en sonda).
- Uyarı: _ekle(rapor, satir_index, 'kolon', deger, 'UYARI: ...') — KONTROL_UYARI ve OZET'e yazılır.

### 3) Test yazın: tests/test_ek_kontroller.py

```
python -m unittest discover -s tests
```

Rehberi (docs/KULLANIM_REHBERI.md) güncellemek için: python src/rehber.py
