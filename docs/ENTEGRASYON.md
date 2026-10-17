# DIA Web Servis ve ETA / Logo / Netsis Veritabanı Araştırması

Bu doküman, DIA'ya web servisle veri yazmak ve ETA / Logo / Netsis'ten veri çekmek
için yapılan internet araştırmasının özetidir. Her bilgi için güven seviyesi verilmiştir:

- ✅ **Doğrulandı**: resmi doküman veya tablo dokümantasyonunda görüldü
- 🟡 **Muhtemel**: birden fazla topluluk kaynağında/SQL örneğinde geçiyor
- ❓ **Doğrulanmalı**: kaynak bulunamadı, kendi veritabanınızda `sql/01_sema_kesif.sql` ile kontrol edin

> Not: Araştırma sırasında `doc.dia.com.tr` sayfalarına doğrudan erişilemedi; DIA bilgileri
> arama sonuçlarındaki doküman özetlerinden derlendi. Kodu yazarken ilgili sayfayı açıp
> alan adlarını teyit edin (bağlantılar aşağıda).

---

## 1. DIA Web Servis (WS API v3)

### 1.1 Gerekenler ✅
DIA'nın "E-Ticaret entegrasyonu yol haritası" sayfasına göre:
1. Firma yetkilisinden **kullanıcı adı, şifre ve apikey** alınır.
2. Kullanıcı detayında **"Web Servis Çağırabilir"** işaretli ve yeterli yetkili olmalı.
3. DİA uygulamasında Firma Listesi'nden **firma_kodu**, firma detayındaki Dönem Listesi'nden **donem_kodu** not edilir.
4. Denemeler **DİA Web Servis Test Aracı** ile yapılabilir.

### 1.2 Adresler ✅
```
https://SUNUCUKODU.ws.dia.com.tr/api/v3/<modul>/json
```
`<modul>`: `sis` (sistem, login), `scf` (stok-cari-fatura), `prj`, `krg`, `rpr` (raporlama), `miy` (CRM) …
Tüm çağrılar **POST + JSON**. Örneklerde test sunucusu olarak `diademo` kullanılıyor.

### 1.3 Oturum (login) ✅
```json
POST https://SUNUCUKODU.ws.dia.com.tr/api/v3/sis/json
{
  "login": {
    "username": "KULLANICI",
    "password": "SIFRE",
    "disconnect_same_user": "True",
    "lang": "tr",
    "params": { "apikey": "APIKEY" }
  }
}
```
Başarılı cevap: `{"code": "200", "msg": "<session_id>"}`. Sonraki tüm çağrılarda bu `session_id` kullanılır.
- apikey **sadece login'de** gönderilir; oturum o anahtara bağlanır.
- Oturum zaman aşımı **1 saat**. `Invalid session` hatası → yeniden login.
- `disconnect_same_user`: aynı kullanıcının açık oturumunu düşürür ("True"/"False").

### 1.4 Servis türleri ✅
Servis adı = `<modul>_<tablo>_<işlem>`. Örn. `scf_carikart_ekle`.

| Son ek | İş | Not |
|---|---|---|
| `_listele` | Liste | `filters`, `sorts`, `params`, `limit`, `offset` (limit 0 = hepsi) |
| `_getir` | Tek kayıt | `_key` ile |
| `_ekle` | Ekle | `kart` gönderilir, cevapta yeni kaydın **`_key`** değeri döner |
| `_guncelle` | Güncelle | `kart` içinde **`_key` zorunlu** |
| `_sil` | Sil | |

**Ekleme isteği iskeleti** ✅
```json
POST https://SUNUCUKODU.ws.dia.com.tr/api/v3/scf/json
{
  "scf_carikart_ekle": {
    "session_id": "…",
    "firma_kodu": 1,
    "donem_kodu": 1,
    "kart": {
      "carikartkodu": "C001",
      "unvan": "ABC GIDA SAN. TİC. A.Ş.",
      "carikarttipi": "AL",
      "verginumarasi": "0123456789"
    }
  }
}
```
- Gönderilmeyen alanlara **varsayılan değer** yazılır (ör. `carikarttipi` gönderilmezse `AL`).
- Zorunlu alanlar gönderilmeli (ör. `carikartkodu`).
- **Combo alanlar kısa kodla** gönderilir (ör. `scf_stokkart.stokkartturu` için "Ticari Mal" = `TCR`).
- **`_key_…` alanları** başka tabloya bağlantıdır (ör. `_key_sis_doviz`, `_key_sis_vergidairesi`).
  Key bilinmiyorsa `{"kod": "…"}` gibi **arama nesnesi** verilebilir; tek eşleşme bulunursa kullanılır,
  hiç ya da birden fazla eşleşme varsa uyarı döner. 🟡 (Arama nesnesindeki alan adı tabloya göre değişir.)
- Alt kayıtlar aynı istekte gönderilebilir: stok kartında `m_birimler`, onun altında `m_barkodlar`. ✅
- Güncellemede alt kayıt silmek için `m_silinecek_birimler`, `m_silinecek_barkodlar` … etiketleri var. ✅
- Hata cevabı: `{"code": "400", "msg": "…açıklama…"}`.

**Listeleme** ✅
```json
{
  "scf_carikart_listele": {
    "session_id": "…", "firma_kodu": 1, "donem_kodu": 1,
    "filters": [{"field": "carikarttipi", "operator": "=", "value": "AL"}],
    "sorts":   [{"field": "carikartkodu", "sorttype": "ASC"}],
    "params":  {},
    "limit": 0, "offset": 0
  }
}
```
Operatörler: `=`, `!=`, `IN` (virgülle çoklu değer)… Değişenleri çekmek için `_date` alanına filtre.
Sadece bazı kolonları almak için `params` içinde kolon seçimi yapılabilir. Stok listesinde
`miktarlariHesapla` parametresi miktar hesabını kapatıp hızlandırır.

### 1.5 Kontör, hatalar ve diğer servisler ✅
- **Kontör:** her web servis çağrısı **0,0125 kontör** düşer; `login` / `logout` düşmez.
  Kalan bakiye: `sis_kontor_sorgula`. Toplu aktarımda (ör. 10.000 cari = ~125 kontör) önceden kontrol edin.
- **Hata kodları:** `406` CREDIT_ERROR (kontör yetersiz), `419` LOGIN_TIMEOUT (oturum düştü → yeniden login),
  `400` geçersiz veri (`msg` alanında açıklama). Detay: [Hata Takibi](https://doc.dia.com.tr/doku.php?id=gelistirici%3Awsapi%3Ahata_takibi).
- **Getir servisleri** (`*_getir`): `key` (Integer64, kaydın `_key`'i) zorunlu; silmede de `key` zorunlu.
- **Şehir anahtarı:** `sis_sehirler_listele` (sis modülü, `filters/sorts/limit/offset`) → il adından
  `_key_sis_sehirler` bulmak için. Adres eklerken bu key kullanılır.
- **Test aracı:** [DİA Web Service Tester](https://doc.dia.com.tr/doku.php?id=gelistirici:wsapi:wstester).
- Ana sayfa bölümleri: Genel Bakış (kontör, güvenlik, isimlendirme, input/response) · Test Aracı ·
  Session & API Key (apikey alma, login, logout) · Yetkili firma/dönem bilgilerinin alınması ·
  Servis türleri (listele, getir, ekle, güncelle) · Hata takibi · modül sayfaları (SIS, SCF, RPR, GTS, EFA…).

### 1.6 Bu proje için kullanılacak servisler
| İş | Servis | Model (alan listesi için) |
|---|---|---|
| Cari kart | `scf_carikart_ekle` / `_guncelle` / `_listele` | `ScfCarikartModel` |
| Cari adres | `scf_carikart_adresleri_ekle` / `_listele` | `ScfCarikartAdresleriModel` (`_key_scf_carikart`, `_key_sis_sehirler`, `_key_sis_ulkeler`) |
| Stok kart | `scf_stokkart_ekle` / `_guncelle` / `_listele` | `ScfStokkartModel` (+ `m_birimler`, `m_barkodlar`) |
| Stok birim | `scf_stokkart_birimleri_listele` | `ScfStokkartBirimleriModel` (`birimkod` ör. `AD`) |
| Fiyat kartı | `scf_fiyatkart_ekle` | |

Model alanlarından doğrulananlar: `carikartkodu`, `unvan`, `carikarttipi` Unicode(16) (varsayılan `AL`),
`verginumarasi` Unicode(16), `tckimlikno`, `_key_sis_doviz` Integer64; stokta `stokkartkodu`, `aciklama`,
`anabirimkey`, `kdvdurumu`, `fiyat1`, `fiyat2`, `ozelkod2`. ✅
> ⚠️ Adreste il/ilçe **`_key_sis_sehirler`** gibi bağlantı alanıyla tutulur; il adını DIA'daki şehir
> tablosundaki kayda eşlemek gerekir. Bizim il/ilçe düzeltmesi (resmi il adları) bu eşlemeyi kolaylaştırır.

### 1.7 Önerilen akış
```
Kaynak (Logo/Netsis/ETA SQL veya Excel)
   → DIA_Veri_Temizleme (temizle + mükerrer + il/ilçe + VKN)
   → cari_dia_hazir.xlsx
   → DIA WS: login → *_listele (mevcutları al) → yoksa *_ekle, varsa *_guncelle
   → sonuç raporu (code/msg) Excel'e
```
Öneriler: oturumu 1 saatten önce yenileyin; kodu aynı olan kayıtta `_ekle` yerine `_guncelle`;
her satırın `code`/`msg` sonucunu rapora yazın; önce `diademo`/test firmasında deneyin.

---

## 2. Logo (Tiger / Go) – MS SQL

### 2.1 Tablo isimlendirme ✅
- Kart tabloları: `LG_<firma>_<TABLO>` → firma 1 için `LG_001_CLCARD`
- Dönem tabloları (hareketler): `LG_<firma>_<dönem>_<TABLO>` → `LG_001_01_CLFLINE`
- Stok miktarı görünümü: `LV_001_01_STINVTOT` (`STOCKREF = ITEMS.LOGICALREF`, tüm ambarlar `INVENNO = -1`) 🟡

### 2.2 Cari – `LG_001_CLCARD` ✅
| Alan | Açıklama |
|---|---|
| `LOGICALREF` | Kayıt ref. |
| `ACTIVE` | 0 = aktif, 1 = pasif |
| `CARDTYPE` | 1 Alıcı, 2 Satıcı, 3 Alıcı+Satıcı 🟡 |
| `CODE` (17) | Cari kodu |
| `DEFINITION_` (51) | Ünvan |
| `SPECODE`, `CYPHCODE` | Özel kod, yetki kodu |
| `ADDR1`, `ADDR2` | Adres |
| `CITY`, `CITYCODE` | Şehir |
| `TOWN`, `TOWNCODE` | İlçe |
| `DISTRICT` | Semt |
| `COUNTRY`, `COUNTRYCODE`, `POSTCODE` | Ülke, posta kodu |
| `TELNRS1`, `TELNRS2`, `FAXNR`, `CELLPHONE` | Telefonlar |
| `TAXNR`, `TAXOFFICE`, `TAXOFFCODE` | Vergi no, vergi dairesi |
| `EMAILADDR`, `WEBADDR`, `INCHARGE` | E-posta, web, ilgili |
| `CCURRENCY` | Döviz türü |
| `TCKNO`, `ISPERSCOMP`, `NAME`, `SURNAME` | 🟡 Yeni sürümlerde: şahıs şirketi ise `ISPERSCOMP=1`, TC no `TCKNO` |

### 2.3 Stok – `LG_001_ITEMS` ✅
| Alan | Açıklama |
|---|---|
| `CODE`, `NAME` | Malzeme kodu, adı |
| `CARDTYPE` | 1 TM Ticari Mal, 2 KK, 3 DM, 4 SK, 10 HM, 11 YM, 12 MM, 13 TK, 20/21 Malzeme sınıfı |
| `ACTIVE` | 0 aktif |
| `STGRPCODE`, `SPECODE`, `PRODUCERCODE` | Grup, özel kod, üretici kodu |
| `VAT` | KDV oranı |
| `UNITSETREF` | → `LG_001_UNITSETF.LOGICALREF` |

Birimler: `LG_001_UNITSETL` (`CODE`, `NAME`, `UNITSETREF`, `MAINUNIT`=ana birim, `CONVFACT1/2`) ✅
Barkod: `LG_001_ITMUNITA.BARCODE` (`ITEMREF` → ITEMS, `UNITLINEREF` → UNITSETL) ✅ — yeni sürümlerde ayrıca `LG_001_UNITBARCODE` tablosu ❓
Fiyat: `LG_001_PRCLIST` (`CARDREF` → ITEMS, `PRICE`, `PTYPE` 1 alış / 2 satış, `CURRENCY`, `INCVAT`, `UOMREF`) ✅
Hizmet: `LG_001_SRVCARD` (`CODE`, `DEFINITION_`, `VAT`) ✅

### 2.4 Diğer kartlar ✅
| Modül | Tablo | Ana alanlar |
|---|---|---|
| Hesap planı | `LG_001_EMUHACC` | `CODE`, `DEFINITION_`, `LEVEL_`, `ACCTYPE`, `ACTIVE` |
| Banka | `LG_001_BNCARD` | `CODE`, `DEFINITION_`, `BRANCH`, `BRANCHNO`, `CITY`, `TOWN` |
| Banka hesabı | `LG_001_BANKACC` | `CODE`, `DEFINITION_`, `BANKREF`→BNCARD, `CURRENCY`, `ACCOUNTNO` |
| Özel kodlar | `LG_001_SPECODES` | `SPECODETYPE`, `SPECODE`, `DEFINITION_` |

Hazır sorgular: `sql/logo_cari.sql`, `sql/logo_stok.sql`.

---

## 3. Netsis – MS SQL

Her şirketin ayrı veritabanı vardır (örn. `SIRKET2024`). Tablolar `TBL` önekli.

| Tablo | Açıklama | Bilinen alanlar |
|---|---|---|
| `TBLCASABIT` ✅ | Cari sabit (kart) | `CARI_KOD`, `CARI_ISIM`, `CARI_TEL`, `CARI_IL`, `CARI_ILCE`, `VERGI_NUMARASI`, `VERGI_DAIRESI`, `CARI_ADRES`, `POSTAKODU`, `CARI_TIP`, `GRUP_KODU`, `FAX` 🟡 |
| `TBLCASABITEK` ✅ | Cari kullanıcı tanımlı sahalar | TC kimlik no burada veya `TBLCASABIT`'te olabilir ❓ |
| `TBLCARIEMAIL` ✅ | Cari e-posta tanımları | ❓ |
| `TBLCAHAR` ✅ | Cari hareket | |
| `TBLSTSABIT` ✅ | Stok kartı | `STOK_KODU`, `STOK_ADI`, `GRUP_KODU`, `KOD_1`…, `OLCU_BR1`, `SATIS_FIAT1` 🟡, `KDV_ORANI` ❓ |
| `TBLSTSABITEK` ✅ | Stok kullanıcı tanımlı sahalar | `STOK_KODU` ile bağlanır |
| `TBLSTHAR` ✅ | Stok hareket | |
| `TBLMUPLAN` 🟡 | Muhasebe hesap planı | ❓ |
| `TBLBNKHESSABIT` 🟡 | Banka hesapları | ❓ |

Hazır sorgular: `sql/netsis_cari.sql`, `sql/netsis_stok.sql` (❓ alanlar yorum satırında).

---

## 4. ETA (ETA:SQL / ETA V8 SQL) – MS SQL

| Tablo | Açıklama | Bilinen alanlar |
|---|---|---|
| `CARKART` ✅ | Cari kart | `CARKOD`, `CARUNVAN`, `CARBAKIYE` |
| `ADRESLER` ✅ | Cari adresleri | ❓ |
| `CARHAR` ✅ | Cari hareket | |
| `STKKART` ✅ | Stok kartı | `STKKOD`, `STKCINSI`, `STKBIRIM` 🟡 |
| `DEPSTKKART` ✅ | Depo bazında stok | |
| `STKFIYAT`, `STKFIYTIP` ✅ | Fiyat listeleri, fiyat tipleri | ❓ |
| `STKBARKOD` ✅ | Barkod | `STKBARSTKKOD`, `STKBARKOD` |

ETA için kamuya açık alan listesi çok az. **Önce `sql/01_sema_kesif.sql` çalıştırın**,
sonucu Excel'e kaydedip gönderin; sorgular gerçek alan adlarıyla tamamlanır.

---

## 5. DIA alanı ↔ kaynak alan eşlemesi

Sol kolon DIA_Veri_Temizleme'nin kullandığı DIA alan adlarıdır (`kurallar.py`).

### Cari
| DIA | Logo `CLCARD` | Netsis `TBLCASABIT` | ETA `CARKART` |
|---|---|---|---|
| carikartkodu | CODE | CARI_KOD | CARKOD |
| unvan | DEFINITION_ | CARI_ISIM | CARUNVAN |
| verginumarasi | TAXNR | VERGI_NUMARASI | ❓ |
| vergidairesi | TAXOFFICE | VERGI_DAIRESI | ❓ |
| tckimlikno2 | TCKNO 🟡 | ❓ | ❓ |
| adres1 / adres2 | ADDR1 / ADDR2 | CARI_ADRES | ADRESLER ❓ |
| sehir / ilce | CITY / TOWN | CARI_IL / CARI_ILCE | ❓ |
| postakodu | POSTCODE | POSTAKODU | ❓ |
| telefon1 / telefon2 | TELNRS1 / TELNRS2 | CARI_TEL / ❓ | ❓ |
| ceptel / fax | CELLPHONE / FAXNR | ❓ / FAX | ❓ |
| eposta / web adresi | EMAILADDR / WEBADDR | TBLCARIEMAIL ❓ | ❓ |
| grupkodu / ozelkod1 | ❓ / SPECODE | GRUP_KODU / ❓ | ❓ |

### Stok
| DIA | Logo | Netsis `TBLSTSABIT` | ETA `STKKART` |
|---|---|---|---|
| stok kart kodu | ITEMS.CODE | STOK_KODU | STKKOD |
| stok adı | ITEMS.NAME | STOK_ADI | STKCINSI |
| Ana Birim | UNITSETL.CODE (MAINUNIT=1) | OLCU_BR1 | STKBIRIM 🟡 |
| KDV Satış Oran | ITEMS.VAT | KDV_ORANI ❓ | ❓ |
| Barkod No | ITMUNITA.BARCODE | ❓ | STKBARKOD.STKBARKOD |
| Fiyat 1 | PRCLIST.PRICE (PTYPE=2) | SATIS_FIAT1 | STKFIYAT ❓ |
| Özel Kod 1 | ITEMS.SPECODE | KOD_1 | ❓ |
| Marka | ❓ | ❓ | ❓ |

---

## 6. Kaynaklar
- DIA: [Ana sayfa](https://doc.dia.com.tr/doku.php?id=gelistirici:wsapi:anasayfa) · [Genel Bakış](https://doc.dia.com.tr/doku.php?id=gelistirici%3Awsapi%3Agenel) ·
  [Session & API Key](https://doc.dia.com.tr/doku.php?id=gelistirici%3Awsapi%3Asession) ·
  [SCF Cari](https://doc.dia.com.tr/doku.php?id=gelistirici%3Awsapi%3Acari) ·
  [Ekleme servisleri](https://doc.dia.com.tr/doku.php?id=gelistirici%3Awsapi%3Aservis_turu_ekleme) ·
  [Güncelleme servisleri](https://doc.dia.com.tr/doku.php?id=gelistirici%3Awsapi%3Aservis_turu_guncelleme) ·
  [Listeleme servisleri](https://doc.dia.com.tr/doku.php?id=gelistirici%3Awsapi%3Aservis_turu_listeleme) ·
  [ScfCarikartModel](https://doc.dia.com.tr/doku.php?id=gelistirici%3Amodels%3Ascf_carikart_model) ·
  [ScfCarikartAdresleriModel](https://doc.dia.com.tr/doku.php?id=gelistirici%3Amodels%3Ascf_carikart_adresleri_model) ·
  [ScfStokkartBirimleriModel](https://doc.dia.com.tr/doku.php?id=gelistirici%3Amodels%3Ascf_stokkart_birimleri_model) ·
  [Stok listeleme scripti](https://doc.dia.com.tr/doku.php?id=gelistirici%3Awsapi%3Aornek_script) ·
  [Gereken bilgiler](https://doc.dia.com.tr/doku.php?id=gelistirici%3Awsapi%3Asss_ws_0) ·
  [E-Ticaret yol haritası](https://doc.dia.com.tr/doku.php?id=gelistirici%3Awsapi%3Asss_ws_28) ·
  [SSS](https://doc.dia.com.tr/doku.php?id=gelistirici%3Awsapi%3Asss) ·
  [Örnek 1 (Python)](https://doc.dia.com.tr/doku.php?id=gelistirici%3Awsapi%3Aornek_1)
- Logo: [ugurozpinar/logo – tablo açıklamaları](https://github.com/ugurozpinar/logo) ·
  [LG_CLCARD](https://ugurozpinar.github.io/Logo/Tablo%20A%C3%A7%C4%B1klamalar%C4%B1%20Yeni/LG_CLCARD/) ·
  [Logo Tiger 3 cari sevkiyat adresleri](https://docs.logo.com.tr/pages/viewpage.action?pageId=22281353) ·
  [Logo cari hesap bilgileri raporu](https://logoyazilimdestek.com/logo-cari-hesap-bilgileri-raporu/) ·
  [Logo Tiger depo bazında stok](https://www.kemalbayat.com.tr/2015/01/30/logo-tiger-depo-bazinda-malzemestok-durumlari/) ·
  [Logo j-Platform REST](https://docs.logo.com.tr/public/jct/integration-interfaces/j-platform-rest)
- Netsis: [Evoset tablo açıklamaları](https://www.evoset.com.tr/netsis-veritabani-tablo-aciklamalari/) ·
  [Piasoft Netsis tabloları](https://www.piasoft.com.tr/netsis-tablolari-ve-aciklamalari) ·
  [netsisnedir tablo açıklamaları](https://netsisnedir.wordpress.com/2011/05/25/netsis-tablo-aciklamalari/) ·
  [10em – şirketler arası aktarım](https://10em.net/knowledgebase.php?article=207) ·
  [Cari borç alacak SQL](https://ismuhendisi.wordpress.com/2015/07/11/netsis-cari-borc-alacak-listesi-raporu-ms-sql-kodu/) ·
  [erpne stok kartı](https://www.erpne.org/stok-karti-kayitlari_topic1808.html)
- ETA: [EtaSQL tabloları](http://delphisizasla.blogspot.com/2013/01/etasqlde-bulunan-tablolarn-anlamlar.html) ·
  [ETA SQL SSS (Elkon)](https://elkon.com.tr/eta-sql-ve-eta-v8-sql-sik-sorulan-sorular-genel-2009-ekim.html) ·
  [ETA müşteri SSS (PDF)](https://cdn.eta.com.tr/pdf/eta-musterilerimiz-icin-toplu-sorular-01-07-2025.pdf) ·
  [ETA Excel'den stok aktarımı](http://www.etayazilim.org/stokexcelden-stok-modulune-kart-aktarimi-import-and-export-data-ile/)
