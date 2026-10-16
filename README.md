# DIA Kurulum Veri Temizleme

Müşteri Excel dosyalarını (cari, stok, banka, personel, demirbaş, hesap planı,
muhasebe kodu, ürün ağacı) DIA'ya aktarılabilir hale getiren Tkinter uygulaması.

## Dosyalar

| Dosya | Görev |
|---|---|
| `src/gui.py` | Arayüz (ana pencere, Ayarlar, Eşleme Öğren, Hakkında) |
| `src/dia_isle.py` | Ana akış: oku → temizle → mükerrer → kaydet (komut satırından da çalışır) |
| `src/_ortak.py` | Temizlik fonksiyonları ve **parametrik ayar tanımları** (`AYAR_TANIMLARI`) |
| `src/il_ilce.py` | İl / ilçe ayrıştırma ve düzeltme |
| `src/il_ilce_veri.py` | 81 il, 973 ilçe listesi |
| `src/kurallar.py` | `kurallar.xlsx` yoksa kullanılan varsayılan kurallar |
| `tests/` | Otomatik testler: `python -m unittest discover -s tests` |

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

Yeni bir ayar eklemek için `src/_ortak.py` içindeki `AYAR_TANIMLARI` listesine bir satır
ekleyin ve kodda `ayar(modul, 'ayar_adi')` ile okuyun; Ayarlar penceresi onu otomatik gösterir.

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
