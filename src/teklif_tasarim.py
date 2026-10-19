"""Teklif PDF / Excel tasarimlari: sablon (Kurumsal / Modern / Sade), renkler, gosterilecek alanlar, metinler.

Tasarimlar teklif.db'deki 'tasarim' tablosunda JSON olarak tutulur; biri varsayilandir (sabitlenir).
Firma ve hazirlayan bilgileri Ayarlar'dadir (tum tasarimlarda ortak).
"""
import copy
import json
import re

SABLONLAR = ['Kurumsal', 'Modern', 'Sade']

# (anahtar, ekranda gorunen ad, varsayilan)
ALANLAR = [
    ('on_yazi', 'Ön yazı ("Sayın ...")', True),
    ('taraflar', 'Teklifi Veren / Teklif Edilen kutuları', True),
    ('hazirlayan', 'Hazırlayan kişi (ad, unvan, telefon)', True),
    ('paket_kodu', 'Paket kodları', True),
    ('paket_aciklama', 'Paket içeriği (kısa açıklama)', False),
    ('kullanici_kolonu', 'Kapsam kolonu (kullanıcı / adet)', True),
    ('liste_indirim', 'Liste fiyatı ve indirim kolonları', True),
    ('dia_ek_bilgi', 'Ek kullanıcı ve yıllık bulut hizmet bedeli', True),
    ('uyarlama_aciklama', 'Uyarlama açıklama metni', True),
    ('uyarlama_dayanak', 'Uyarlama: dayanak satırı', True),
    ('uyarlama_hesap', 'Uyarlama: saat × ücret hesabı', True),
    ('uyarlama_gun', 'Uyarlama: iş günü karşılığı', True),
    ('kapsam_haric', 'Kapsam dışı notu', True),
    ('ozet_kdv', 'Özet: KDV ve KDV dahil toplam', True),
    ('odeme', 'Ödeme seçenekleri', True),
    ('bakim', 'Bakım / destek anlaşması seçenekleri', True),
    ('tco', 'Çok yıllık toplam maliyet', True),
    ('sartlar', 'Ticari şartlar', True),
    ('banka', 'Banka / IBAN bilgileri', True),
    ('notlar', 'Notlar', True),
    ('imza', 'İmza / onay alanı', True),
    ('alt_bilgi', 'Alt bilgi (firma + sayfa no)', True),
]

VARSAYILAN = {
    'sablon': 'Kurumsal',
    'ana_renk': '#1F3A5F',
    'vurgu_renk': '#C9A227',
    'yazi_boyutu': 9,
    'baslik': 'DİA ERP LİSANS VE HİZMET TEKLİFİ',
    'on_yazi': 'Sayın {yetkili},\n\n{musteri} için ihtiyaçlarınız doğrultusunda hazırladığımız DİA ERP lisans ve '
               'uyarlama hizmeti teklifimizi bilgilerinize sunarız. Teklifimizle ilgili her türlü sorunuz için '
               'bizimle iletişime geçebilirsiniz.',
    'sartlar': 'Teklif {gecerlilik} tarihine kadar geçerlidir.\n'
               'Fiyatlara KDV dahil değildir; KDV ayrıca gösterilmiştir.\n'
               'Lisans bedelleri DİA Yazılım güncel fiyat listesine göredir.\n'
               'Uyarlama hizmeti sözleşme imzasını takiben karşılıklı planlanan takvimle yürütülür.',
    'kapanis': 'Saygılarımızla,',
    'goster': {k: v for k, _, v in ALANLAR},
}

HAZIR_TASARIMLAR = {
    'Kurumsal': {},
    'Modern': {'sablon': 'Modern', 'ana_renk': '#0E5A73', 'vurgu_renk': '#F28C28'},
    'Sade': {'sablon': 'Sade', 'ana_renk': '#333333', 'vurgu_renk': '#888888',
             'goster': {'on_yazi': False, 'uyarlama_dayanak': False, 'bakim': False, 'tco': False,
                        'paket_kodu': False, 'dia_ek_bilgi': False}},
}


def tamamla(tasarim=None):
    """Eksik anahtarlari varsayilanla doldurur (eski kayitlar / yeni eklenen alanlar icin)."""
    t = copy.deepcopy(VARSAYILAN)
    tasarim = tasarim or {}
    for k, v in tasarim.items():
        if k == 'goster':
            t['goster'].update(v or {})
        elif v is not None:
            t[k] = v
    if t['sablon'] not in SABLONLAR:
        t['sablon'] = 'Kurumsal'
    for k in ('ana_renk', 'vurgu_renk'):
        if not re.fullmatch(r'#[0-9a-fA-F]{6}', str(t[k])):
            t[k] = VARSAYILAN[k]
    try:
        t['yazi_boyutu'] = min(max(float(t['yazi_boyutu']), 7), 12)
    except (TypeError, ValueError):
        t['yazi_boyutu'] = VARSAYILAN['yazi_boyutu']
    return t


def hazir(ad):
    return tamamla(HAZIR_TASARIMLAR.get(ad, {}))


def json_yaz(t):
    return json.dumps(tamamla(t), ensure_ascii=False)


def json_oku(metin):
    try:
        return tamamla(json.loads(metin) if metin else {})
    except (TypeError, ValueError):
        return tamamla()


def yer_tutucu(metin, degerler):
    """'{musteri}' gibi alanlari doldurur; bilinmeyen alan oldugu gibi kalir."""
    return re.sub(r'\{(\w+)\}', lambda m: str(degerler.get(m.group(1), m.group(0))), metin or '')


def acik_ton(renk, oran=0.88):
    """'#1F3A5F' -> beyaza dogru acilmis ton (tablo satirlari / zemin icin)."""
    r, g, b = (int(renk[i:i + 2], 16) for i in (1, 3, 5))
    return '#%02X%02X%02X' % tuple(int(x + (255 - x) * oran) for x in (r, g, b))
