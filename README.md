# DIA Kurulum Veri Temizleme

Müşteri Excel dosyalarını (cari, stok, banka, personel, demirbaş, hesap planı,
muhasebe kodu, ürün ağacı) DIA'ya aktarılabilir hale getiren Tkinter uygulaması.

**Kullanım rehberi:** programda **📖 Rehber** butonu (Ayarlar penceresinde de var; seçili ayarın
açıklama ve örneklerine gider) veya [docs/KULLANIM_REHBERI.md](docs/KULLANIM_REHBERI.md).

## Dosyalar

| Dosya | Görev |
|---|---|
| `src/gui.py` | Arayüz (ana pencere, Ayarlar, Eşleme Öğren, Hakkında) |
| `src/dia_isle.py` | Ana akış: oku → temizle → mükerrer → kaydet (komut satırından da çalışır) |
| `src/_ortak.py` | Temizlik fonksiyonları ve **parametrik ayar tanımları** (`AYAR_TANIMLARI`) |
| `src/ek_kontroller.py` | Ek kontroller (ünvan, vergi dairesi, barkod, birim, IBAN→banka, TDHP, benzer kayıt, değişiklik raporu ...) |
| `src/ek_ayarlar.py` | Ek kontrollerin parametrik ayar tanımları (modül bazında) |
| `src/rehber.py`, `src/rehber_gui.py` | Kullanım rehberi içeriği ve penceresi (`python src/rehber.py` → docs/KULLANIM_REHBERI.md) |
| `src/il_ilce.py` | İl / ilçe ayrıştırma ve düzeltme |
| `src/il_ilce_veri.py` | 81 il, 973 ilçe listesi |
| `src/kurallar.py` | `kurallar.xlsx` yoksa kullanılan varsayılan kurallar |
| `tests/` | Otomatik testler: `python -m unittest discover -s tests` |
| `src/logo_kaynak.py` | Logo (MS SQL) veritabanından cari/stok okuma (sadece okur) |
| `src/dia_ws.py`, `src/dia_aktar.py` | DIA web servis istemcisi ve aktarım (ekle/güncelle, rapor) |
| `src/entegrasyon_gui.py` | "Logo'dan Çek" ve "DIA'ya Gönder" pencereleri |
| `docs/ENTEGRASYON.md` | DIA web servis + ETA / Logo / Netsis veritabanı araştırması |
| `sql/` | Şema keşif ve Logo/Netsis/ETA'dan cari-stok çekme sorguları (sadece okur) |

## Parametrik ayarlar

Tüm yeni özellikler **⚙ Ayarlar → Ayarlar** sekmesinden modül bazında yönetilir
(`kurallar.xlsx` → `AYARLAR` sayfası). Eski bir `kurallar.xlsx` kullanıyorsanız
**＋ Eksik Ayarları Ekle** butonuna basıp **KAYDET** deyin. Excel'de satırı olmayan
ayar için varsayılan değer kullanılır.

| Ayar | Varsayılan | Açıklama |
|---|---|---|
| `telefon_format` | `10` | `10` → 5321234567, `11` → 05321234567, `90` → 905321234567, `+90` → +905321234567 |
| `vkn_sifir_tamamla` | `True` | Excel'in sildiği baştaki 0'ları VKN'ye geri ekler (8-9 hane → 10 hane) |
| `vkn_algoritma_kontrol` | `True` | VKN kontrol hanesi hatalıysa uyarı yazar |
| `tc_algoritma_kontrol` | `True` | TC Kimlik No algoritması hatalıysa uyarı yazar |
| `il_ilce_duzelt` | `True` | Aynı hücreye yazılmış / yer değiştirmiş il-ilçeyi düzeltir |
| `il_bos_ise_doldur` | `True` | İl boşsa ve ilçe tek bir ile aitse ili doldurur |
| `il_ilce_yazim` | `BUYUK` | `BUYUK` → İSTANBUL, `BASLIK` → İstanbul |
| `il_sutunu` / `ilce_sutunu` | `sehir` / `ilce` | İl ve ilçe kolon adları |

### Ek kontroller (hepsi Ayarlar'dan açılıp kapatılır)

Her ayarın açıklaması, formülü ve örneği Ayarlar penceresinde satıra çift tıklayınca görünür.
Liste ayarlarını boş bırakmak varsayılanı kullanır; tamamen kapatmak için `YOK` yazın.

| Modül | Ne yapar | Ayar |
|---|---|---|
| Hepsi | `#YOK`, `#N/A`, `-`, `.` gibi değerleri boşaltır (0'a dokunmaz) | `hata_degerleri_temizle`, `hata_degerleri` |
| Hepsi | Sekme, satır sonu, `_x000D_`, görünmez boşluk temizliği | `gorunmez_karakter_temizle` |
| Hepsi | Bozuk Türkçe onarımı (`Ä°STANBUL` → `İSTANBUL`, `ÝZMÝR` → `İZMİR`) | `bozuk_turkce_duzelt` |
| Hepsi | `1,23457E+11` bilimsel gösterimi kurtarır; hane kaybı varsa GEÇERSİZ | `bilimsel_gosterim_kontrol` |
| Hepsi | Benzer kayıt (olası mükerrer) – benzerlik puanıyla, silmez | `benzer_kayit_kontrol`, `_sutunu`, `_esik` |
| Hepsi | Öncesi/sonrası: `*_DEGISIKLIKLER.xlsx` (liste + değişen hücreler turuncu, eski değer notta) | `degisiklik_raporu` |
| Hepsi | DIA zorunlu alan ve azami uzunluk kontrolü | `zorunlu_alanlar`, `alan_uzunluklari` |
| Hepsi | Eşleme hafızası: onaylı sütun eşlemeleri diğer müşterilerde de uygulanır | `esleme_hafizasi` |
| Hepsi | Çapraz kontrol: kod diğer modülün dosyasında var mı (stok/muhkod grup kodu → ürün ağacı) | `capraz_kontrol` |
| Cari | Ünvan kısaltma standardı (`Ltd.Şti` → `LTD. ŞTİ.`, `AŞ` → `A.Ş.`) | `unvan_standart`, `unvan_kisaltmalari` |
| Cari | Vergi dairesi: `V.D.` eki silinir, il/ilçe adları düzeltilir | `vergi_dairesi_duzelt`, `vergi_dairesi_eslesme` |
| Cari | Şirket ünvanı ama sadece TC var uyarısı | `vkn_tc_tutarlilik` |
| Cari | Şahıs carilerinde ad/soyad ayırma (varsayılan kapalı) | `ad_soyad_ayir` |
| Cari | Adresten boş il/ilçeyi doldurma; adres kısaltma (kapalı) | `adresten_il_ilce`, `adres_kisaltma` |
| Cari | Cep numarasını (5xx) cep kolonuna taşıma | `cep_ayir`, `cep_sutunu` |
| Cari, Personel | Posta kodu ↔ il plaka uyumu | `posta_kodu_il_kontrol` |
| Stok | EAN-8/13, UPC-A, GTIN-14 kontrol hanesi | `barkod_kontrol` |
| Stok | Birim standardı (`Adt` → `ADET`), tanımsız birim uyarısı | `birim_standart`, `birim_eslesme` |
| Stok | Eski KDV oranı (%8, %18) uyarısı | `kdv_eski_oranlar` |
| Stok | `1.250,50 TL` → 1250.5, döviz `Döviz N` kolonuna; alış > satış uyarısı | `fiyat_temizle`, `alis_satis_kontrol` |
| Stok | Stok adında ölçü yazımı (`500gr` → `500 GR`), harf yazımı | `stok_adi_birim_duzelt`, `stok_adi_yazim` |
| Banka, Personel | IBAN'dan banka adı, banka adı standardı, uyuşmazlık uyarısı | `iban_banka_adi`, `banka_adi_standart` |
| Banka | IBAN'dan hesap no; hesap adında USD var ama döviz TL uyarısı | `iban_hesap_no_doldur`, `doviz_hesap_adi_kontrol` |
| Personel | Cinsiyet / medeni hal / kan grubu standardı; yaş kontrolü | `personel_standart`, `yas_min`, `yas_max` |
| Demirbaş | Gelecek tarih, 0 tutar, oran-süre uyumu, birikmiş > tutar | `demirbas_kontrol` |
| Hesap planı | Tek Düzen Hesap Planı ana hesap kontrolü | `tdhp_kontrol`, `ek_ana_hesaplar` |
| Hesap planı | Ayraç standardı (`120-01-001` → `120.01.001`), seviye uzunluğu; eksik üst hesap ekleme (kapalı) | `hesap_ayrac`, `hesap_seviye_uzunluklari`, `eksik_ust_hesap_ekle` |
| Ürün ağacı | Aynı grup kodu farklı üst gruplarda (kapalı) | `grup_agaci_tutarlilik` |

`musteriler/` altındaki tüm müşteriler işlenince `musteriler/TOPLU_RAPOR_<tarih>.xlsx`
(müşteri × modül × satır / uyarı / düzeltilen / geçersiz) oluşur.

Yeni bir ayar eklemek için `src/ek_ayarlar.py` içindeki `EK_AYAR_TANIMLARI` listesine bir satır
ekleyin ve kodda `ayar(modul, 'ayar_adi', modul_adi)` ile okuyun; Ayarlar penceresi onu otomatik gösterir.

## Logo → DIA aktarımı

1. **🔌 Logo'dan Çek**: SQL sunucu, veritabanı, firma no girin → *Bağlantı Testi* → *Cari Çek* / *Stok Çek*.
   Müşteri klasörüne `cari_listesi.xlsx` / `stok_listesi.xlsx` yazılır (eskisi `.bak` yedeklenir).
   Bilgisayarda **ODBC Driver 17 (veya 18) for SQL Server** kurulu olmalı. Logo verisi değiştirilmez.
2. Ana ekranda modülü seçip **▶ CALISTIR** (temizleme, il/ilçe, VKN, mükerrer).
3. **☁ DIA'ya Gönder**: sunucu kodu, kullanıcı, şifre, API key, firma/dönem kodu → *Bağlantı Testi* (kontörü gösterir)
   → önce **Kuru çalıştır** işaretliyken *GÖNDER* → `cari_DIA_KURU_RAPOR.xlsx`'i kontrol edin
   → işareti kaldırıp gerçek gönderim → `cari_DIA_AKTARIM_RAPORU.xlsx` (her satırın sonucu).
   Kod DIA'da yoksa eklenir, varsa güncellenir (veya atlanır).
4. **Alan Eslemesi** butonu `dia_eslesme.json`'u açar: hangi Excel kolonunun hangi DIA alanına gideceği.
   `?` ile başlayan satırlar doğrulanmamış alanlardır ve gönderilmez; DIA dokümanına göre kontrol edip
   başındaki `?` işaretini silin.

Bağlantı ayarları müşteri klasöründe `entegrasyon.json`'a yazılır; **şifreler hiçbir zaman kaydedilmez**,
API key sadece "hatırla" işaretlenirse kaydedilir.

## EXE oluşturma (Windows)

```bat
cd DIA_Veri_Temizleme_klasoru
python -m pip install -r requirements.txt pyinstaller
cd src
pyinstaller --onefile --windowed --name DIA_Veri_Temizleme gui.py
```

Oluşan `src\dist\DIA_Veri_Temizleme.exe` dosyasını eskisinin yerine koyun.
`kurallar.xlsx` dosyanızı EXE'nin yanında tutmaya devam edin.

## Komut satırı

```bat
cd src
python dia_isle.py cari            :: inceleme (silmez, _INCELEME.xlsx üretir)
python dia_isle.py cari --uygula   :: INCELEME'de SIL_ONAY=E olanları siler
python dia_isle.py hepsi --sil     :: tüm modüller, otomatik mükerrer silme
python dia_isle.py cari --dry-run  :: dosya yazmadan rapor
```
