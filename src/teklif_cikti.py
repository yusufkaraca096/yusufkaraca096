"""Musteriye gidecek teklif ciktisi: PDF (reportlab) ve Excel (openpyxl).

Icerik once 'bolumler' listesine donusturulur, PDF ve Excel ayni listeden cizilir (iki format hep ayni icerik).
Gorunum ve hangi alanlarin basilacagi 'tasarim'dan gelir (teklif_tasarim: Kurumsal / Klasik / Kapakli / Seritli /
Modern / Sade). Teklif TL, USD veya EUR (elle girilen kur) ve Turkce veya Ingilizce basilabilir.
Ic bilgiler (maliyet, kar, rakip, komisyon, satis asistani) ciktiya YAZILMAZ.
"""
import math
import os
import re
from datetime import datetime, timedelta

import teklif_tasarim

_FONT_ADAYLARI = [
    ('C:/Windows/Fonts/arial.ttf', 'C:/Windows/Fonts/arialbd.ttf'),
    ('C:/Windows/Fonts/segoeui.ttf', 'C:/Windows/Fonts/segoeuib.ttf'),
    ('C:/Windows/Fonts/calibri.ttf', 'C:/Windows/Fonts/calibrib.ttf'),
    ('/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf', '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'),
    ('/Library/Fonts/Arial Unicode.ttf', '/Library/Fonts/Arial Unicode.ttf'),
]
_ASCII = str.maketrans('çğıöşüÇĞİÖŞÜ★✔•', 'cgiosuCGIOSU*+-')
DOVIZLER = ['TL', 'USD', 'EUR']

# Ingilizce teklif icin sabit metinler (veri: paket adlari vb. oldugu gibi kalir)
_EN = {
    'Teklif No': 'Quote No', 'Tarih': 'Date', 'Geçerlilik': 'Valid until', 'Paket': 'Edition',
    'Kullanıcı': 'Users', 'Müşteri': 'Customer', 'Yetkili': 'Contact person', 'Telefon': 'Phone',
    'E-posta': 'E-mail', 'Adres': 'Address', 'VD / VKN': 'Tax office / No', 'Hazırlayan': 'Prepared by',
    'İletişim': 'Contact', 'TEKLİFİ VEREN': 'FROM', 'TEKLİF EDİLEN': 'TO', 'Lisans': 'Licenses',
    'Paket / Modül': 'Package / Module', 'Kapsam': 'Scope', 'Liste Tutarı': 'List Price', 'İndirim': 'Discount',
    'Net Tutar': 'Net Amount', 'LİSANS TOPLAMI': 'LICENSE TOTAL',
    'Ek +1 masaüstü kullanıcı bedeli': 'Each additional desktop user',
    'Ek +1 mobil kullanıcı bedeli': 'Each additional mobile user',
    'Yıllık bulut hizmet bedeli': 'Annual cloud service fee',
    'Uyarlama Hizmeti (Kurulum, Aktarım, Eğitim ve Danışmanlık)':
        'Implementation Services (Installation, Data Migration, Training and Consulting)',
    'Hizmet': 'Service', 'Kapsam ve hesap dayanağı': 'Scope', 'Süre': 'Duration', 'Tutar': 'Amount',
    'UYARLAMA TOPLAMI': 'IMPLEMENTATION TOTAL', 'UYARLAMA HİZMET BEDELİ': 'IMPLEMENTATION FEE',
    'Uyarlama Hizmeti': 'Implementation Services', 'Lisans uyarlama payı': 'License configuration share',
    'Teklif Özeti': 'Quote Summary', 'ARA TOPLAM': 'SUBTOTAL', 'NET TOPLAM (KDV Hariç)': 'NET TOTAL (excl. VAT)',
    'GENEL TOPLAM (KDV Dahil)': 'GRAND TOTAL (incl. VAT)', 'KDV': 'VAT',
    'Ödeme Seçenekleri': 'Payment Options', 'Ödeme Seçeneği': 'Payment option', 'Taksit': 'Installments',
    'Taksit Tutarı (KDV Dahil)': 'Installment (incl. VAT)', 'Toplam (KDV Dahil)': 'Total (incl. VAT)',
    'Not': 'Note', 'Tek ödeme': 'Single payment', 'Vade farksız': 'No interest',
    'Bakım ve Destek Anlaşması Seçenekleri': 'Maintenance and Support Plans',
    'Aylık (KDV Hr.)': 'Monthly (excl. VAT)', 'Yıllık (KDV Hr.)': 'Annual (excl. VAT)',
    'Yanıt Süresi': 'Response time', 'Yerinde': 'On-site', 'Destek': 'Support',
    'Yıllık Toplam Maliyet': 'Year Total Cost of Ownership', 'Yıl': 'Year', 'Lisans + Uyarlama': 'License + Services',
    'DIA Yıllık Bakım': 'DIA Annual Maintenance', 'Destek Anlaşması': 'Support Plan',
    'Toplam (KDV Hr.)': 'Total (excl. VAT)', 'KDV Dahil': 'Incl. VAT', 'TOPLAM': 'TOTAL',
    'Ticari Şartlar': 'Terms and Conditions', 'Banka Bilgileri': 'Bank Details', 'Notlar': 'Notes',
    'MÜŞTERİ ONAYI': 'CUSTOMER APPROVAL', 'İmza': 'Signature', 'Firma Kaşe / İmza': 'Company stamp / Signature',
    'Yetkili Ad Soyad': 'Authorized name', 'Sayfa': 'Page', 'Teklif': 'Proposal',
    'DIA yıllık bakım bedeli DIA\'nın güncel fiyatlarına göre tahminidir.':
        'DIA annual maintenance is estimated from the current DIA price list.',
    'Kapsam dışı': 'Out of scope', 'Hazırlayan kişi': 'Prepared by', 'Önerilen': 'Recommended',
    'Seçilen': 'Selected', 'Revizyon': 'Revision',
}
_HIZMET_EN = {
    'kurulum': ('Installation and configuration', 'Installation of packages, company / period / branch '
                'definitions, parameters and user authorizations'),
    'veri_aktarimi': ('Data migration', 'Cleansing and migration of customer, item and other records to DIA, '
                      'opening balances'),
    'egitim': ('User training', 'User group trainings'),
    'rapor': ('Custom reports and forms', 'Customer-specific reports and form (invoice, delivery note) designs'),
    'entegrasyon': ('Integrations', 'Setup of e-commerce, POS, bank and similar integrations'),
    'canli_gecis': ('Go-live support', 'Go-live day and first week intensive support'),
    'proje': ('Project management', 'Planning, coordination, testing and acceptance'),
}
_BAKIM_EN = {'Temel': 'Basic', 'Standart': 'Standard', 'Premium': 'Premium'}


def para(x):
    """1234567.8 -> '1.234.567,80 TL' (ekranda ve TL tekliflerde)"""
    return bicim(x)


def bicim(x, birim='TL', dil='TR'):
    s = '{:,.2f}'.format(float(x or 0))
    if dil != 'EN':
        s = s.replace(',', 'X').replace('.', ',').replace('X', '.')
    return s + ' ' + birim


def _g(x, dil='TR'):
    if not isinstance(x, (int, float)):
        return str(x)
    return '%g' % x if dil == 'EN' else ('%g' % x).replace('.', ',')


def _tarih(girdi):
    t = girdi.get('tarih')
    try:
        return datetime.strptime(t, '%Y-%m-%d') if t else datetime.now()
    except ValueError:
        return datetime.now()


def doviz_bilgisi(girdi):
    """(birim, kur): TL -> ('TL', 1). Kur elle girilir (1 USD = kur TL)."""
    birim = str(girdi.get('doviz') or 'TL').upper()
    if birim not in DOVIZLER:
        birim = 'TL'
    try:
        kur = float(girdi.get('kur') or 0)
    except (TypeError, ValueError):
        kur = 0
    if birim == 'TL' or kur <= 0:
        return ('TL', 1.0) if birim == 'TL' or kur <= 0 else (birim, kur)
    return birim, kur


class _Dil:
    """Teklifin dili ve para birimi: t('Lisans') -> 'Licenses', p(1000) -> '24,10 USD'."""
    def __init__(self, girdi):
        self.dil = 'EN' if str(girdi.get('dil') or '').upper().startswith('EN') else 'TR'
        self.birim, self.kur = doviz_bilgisi(girdi)

    def t(self, metin):
        return _EN.get(metin, metin) if self.dil == 'EN' else metin

    def deger(self, tl):
        return round(float(tl or 0) / self.kur, 2)

    def p(self, tl):
        return bicim(self.deger(tl), self.birim, self.dil)

    def g(self, x):
        return _g(x, self.dil)

    def saat(self, x):
        return '%s %s' % (self.g(x), 'h' if self.dil == 'EN' else 'saat')


def satir_aciklamasi(s, d=None):
    en = d is not None and d.dil == 'EN'
    if s['tur'].lower().startswith(('mas', 'mob')):
        return (('%d user' if s['adet'] == 1 else '%d users') if en else '%d kullanıcı') % s['adet']
    return (('%d pcs' if en else '%d adet') % s['adet']) if s['adet'] != 1 else ''


def _kisa(metin, n=150):
    metin = re.sub(r'\s+', ' ', metin or '').strip()
    return metin if len(metin) <= n else metin[:n].rsplit(' ', 1)[0] + ' …'


# Musteriye gosterilmeyecek ic gerekceler (elle degisiklik, kalibrasyon)
_IC_GEREKCE = ('Elle girildi', 'Kalibrasyon')


def _sure(saat, ayar, gun=True, d=None):
    d = d or _Dil({})
    gun_saat = float(ayar.get('is_gunu_saat') or 0)
    metin = d.saat(saat)
    if gun and gun_saat > 0 and saat >= gun_saat / 2:
        metin += '\n(≈ %s %s)' % (d.g(round(saat / gun_saat, 1)), 'working days' if d.dil == 'EN' else 'iş günü')
    return metin


def _hizmet_metni(k, d):
    """Kalemin PDF'teki adi ve kapsami: teklife ozel ad > Ingilizce karsilik > tanimli ad."""
    ad, kapsam = k['ad'], k['kapsam']
    if d.dil == 'EN' and k['kod'] in _HIZMET_EN:
        if not k.get('ozel_ad'):
            ad = _HIZMET_EN[k['kod']][0]
        if not k.get('ozel_kapsam'):
            kapsam = _HIZMET_EN[k['kod']][1]
    return ad, kapsam


def uyarlama_satirlari(s, ayar, goster=None, d=None, gorunum='Detaylı'):
    """Uyarlama tablosu: her kalemin kapsami, (Detayli'da) neye gore hesaplandigi, suresi ve tutari."""
    G = goster or teklif_tasarim.VARSAYILAN['goster']
    d = d or _Dil({})
    detayli = gorunum == 'Detaylı'
    u = s['uyarlama']
    rows = []
    for k in s['hizmet']['kalemler']:
        if not k['saat']:
            continue
        ad, metin = _hizmet_metni(k, d)
        dayanak = [n for n in k.get('nedenler', []) if not n.startswith(_IC_GEREKCE)]
        if detayli and dayanak and G.get('uyarlama_dayanak', True) and d.dil == 'TR':
            metin += '\nDayanak: ' + '; '.join(dayanak)
        if detayli and u['saat_ucreti'] and G.get('uyarlama_hesap', True):
            metin += '\n%s: %s × %s = %s' % ('Calculation' if d.dil == 'EN' else 'Hesap', d.saat(k['saat']),
                                             d.p(u['saat_ucreti']), d.p(k['tutar']))
        rows.append([ad, metin.strip(), _sure(k['saat'], ayar, G.get('uyarlama_gun', True), d), d.p(k['tutar'])])
    if u['lisans_payi']:
        metin = ('Configuration of the purchased modules for your company, parameter and version compatibility.'
                 if d.dil == 'EN' else
                 'Satın alınan modüllerin firmanıza göre yapılandırılması, parametre ve sürüm uyumu.')
        if detayli and G.get('uyarlama_hesap', True):
            metin += '\n%s: %s × %%%s = %s' % ('Calculation: net license' if d.dil == 'EN' else 'Hesap: lisans net tutarı',
                                               d.p(s['lisans']['net']), d.g(float(ayar['lisans_yuzdesi'])),
                                               d.p(u['lisans_payi']))
        rows.append([d.t('Lisans uyarlama payı'), metin, '', d.p(u['lisans_payi'])])
    if u['sinir']:
        alt = u['sinir'] == 'ALT SINIR'
        if d.dil == 'EN':
            rows.append(['Minimum fee adjustment' if alt else 'Maximum fee adjustment', '', '',
                         d.p(u['net'] - u['ham'])])
        else:
            rows.append(['Asgari uyarlama farkı' if alt else 'Azami uyarlama indirimi',
                         'Uyarlama bedeli asgari tutara tamamlanmıştır.' if alt else
                         'Uyarlama bedeli azami tutarla sınırlandırılmıştır.', '', d.p(u['net'] - u['ham'])])
    rows.append([d.t('UYARLAMA TOPLAMI'), '', d.saat(u['saat']), d.p(u['net'])])
    return rows


def uyarlama_ozeti(s, ayar, goster=None, d=None, kalemler_dahil=False):
    G = goster or teklif_tasarim.VARSAYILAN['goster']
    d = d or _Dil({})
    u = s['uyarlama']
    en = d.dil == 'EN'
    parca = []
    if kalemler_dahil:
        adlar = [_hizmet_metni(k, d)[0] for k in s['hizmet']['kalemler'] if k['saat']]
        if adlar:
            parca.append(', '.join(adlar))
    parca.append(('incl. VAT %s' if en else 'KDV dahil %s') % d.p(u['dahil']))
    if u['saat']:
        parca.append(('Total %s' if en else 'Toplam %s') %
                     _sure(u['saat'], ayar, G.get('uyarlama_gun', True), d).replace('\n', ' '))
    if u['saat_ucreti'] and not kalemler_dahil:
        parca.append(('Hourly rate %s + VAT' if en else 'Saat ücreti %s + KDV') % d.p(u['saat_ucreti']))
    if u['lisans_payi'] and not kalemler_dahil:
        parca.append('License configuration share included' if en else 'Lisans uyarlama payı dahil')
    return '  |  '.join(parca)


def firma_satirlari(ayar, d=None):
    """Teklifi veren firmanin iletisim satirlari (ad haric)."""
    en = d is not None and d.dil == 'EN'
    satirlar = []
    if ayar.get('firma_adres'):
        satirlar.append(ayar['firma_adres'])
    iletisim = '  |  '.join(x for x in (
        (('Phone: ' if en else 'Tel: ') + ayar['firma_telefon']) if ayar.get('firma_telefon') else '',
        ayar.get('firma_eposta') or '', ayar.get('firma_web') or '') if x)
    if iletisim:
        satirlar.append(iletisim)
    if ayar.get('firma_iletisim'):
        satirlar.append(ayar['firma_iletisim'])
    if ayar.get('firma_vergi'):
        satirlar.append(('Tax office / No: ' if en else 'VD / VKN: ') + ayar['firma_vergi'])
    return satirlar


def _hazirlayan(ayar):
    return [x for x in (ayar.get('hazirlayan_ad'), ayar.get('hazirlayan_unvan')) if x]


def _odeme_adi(ad, d):
    if d.dil != 'EN':
        return ad
    k = re.sub(r'\s+', ' ', ad.lower())
    n = re.search(r'(\d+)\s*taksit', k)
    if 'peşin' in k or 'pesin' in k:
        return 'Advance payment (bank transfer)'
    if 'kredi' in k:
        return 'Credit card, %s installments' % n.group(1) if n else 'Credit card, single payment'
    if 'havale' in k and n:
        return 'Bank transfer, %s installments' % n.group(1)
    if 'çek' in k or 'senet' in k:
        return 'Post-dated cheque'
    if n:
        return '%s installments' % n.group(1)
    return ad


# ======================================================================
# Icerik (bolumler)
# ======================================================================
# ('kapak', baslik, alt, musteri, [(etiket, deger)]) -> yalniz Kapakli sablon cizer
# ('baslik', metin, alt) | ('bilgi', [(etiket, deger)]) | ('taraflar', veren[(e, d)], edilen[(e, d)], basliklar)
# ('paragraf', metin) | ('bolum', metin) | ('liste', [(etiket, deger)]) | ('maddeler', [metin])
# ('tablo', basliklar, satirlar, {'sag': {kolon...}, 'toplam': bool, 'oran': [genislik...], 'vurgu': {satir...}})
# ('ozet', [(etiket, deger, kalin)], son_satir_vurgulu, birim, dil) | ('metin', metin, kucuk)
# ('vurgu', etiket, buyuk_tutar_metni, alt_metin) | ('imza', veren_satirlari)
def bolumler(girdi, s, ayar, tasarim=None):
    T = teklif_tasarim.tamamla(tasarim)
    G = T['goster']
    d = _Dil(girdi)
    t, en = d.t, d.dil == 'EN'
    B = []
    no = girdi.get('teklif_no') or '-'
    alt = []
    m = re.search(r'-R(\d+)$', no)
    if m:
        alt.append('%s %s' % (t('Revizyon'), m.group(1)))
    if girdi.get('etiket'):
        alt.append(girdi['etiket'] + ' ' + t('Teklif'))
    baslik = (T['baslik_en'] if en else T['baslik']) or t('Teklif').upper()

    tarih = _tarih(girdi)
    gecerlilik = (tarih + timedelta(days=int(ayar.get('gecerlilik_gun') or 0))).strftime('%d.%m.%Y')
    mb = girdi.get('musteri_bilgi') or {}
    degerler = {'musteri': girdi.get('musteri') or ('your company' if en else 'Firmanız'),
                'yetkili': mb.get('yetkili') or ('Sir / Madam' if en else 'Yetkili'),
                'tarih': tarih.strftime('%d.%m.%Y'), 'gecerlilik': gecerlilik, 'teklif_no': no,
                'firma': ayar.get('firma_adi', ''), 'hazirlayan': ayar.get('hazirlayan_ad', '')}
    kullanici = ('%d desktop / %d mobile' if en else '%d masaüstü / %d mobil') % (
        int(girdi.get('masaustu') or 0), int(girdi.get('mobil') or 0))
    kapak = [(t('Teklif No'), no), (t('Tarih'), tarih.strftime('%d.%m.%Y')), (t('Geçerlilik'), gecerlilik)]
    if ayar.get('hazirlayan_ad'):
        kapak.append((t('Hazırlayan'), ', '.join(_hazirlayan(ayar))))
    B.append(('kapak', baslik, ' · '.join(alt), girdi.get('musteri', ''), kapak))
    B.append(('baslik', baslik, ' · '.join(alt)))
    bilgi = [(t('Teklif No'), no), (t('Tarih'), tarih.strftime('%d.%m.%Y')), (t('Geçerlilik'), gecerlilik),
             (t('Paket'), s['paket_tipi']), (t('Kullanıcı'), kullanici)]
    if d.birim != 'TL':
        bilgi.append(('Currency' if en else 'Para birimi', '%s (1 %s = %s TL)' % (
            d.birim, d.birim, bicim(d.kur, '', d.dil).strip())))
    if G['taraflar']:
        veren = [('', ayar.get('firma_adi', ''))] + [('', x) for x in firma_satirlari(ayar, d)]
        if G['hazirlayan'] and ayar.get('hazirlayan_ad'):
            veren.append((t('Hazırlayan'), ', '.join(_hazirlayan(ayar))))
            iletisim = '  |  '.join(x for x in (ayar.get('hazirlayan_telefon'), ayar.get('hazirlayan_eposta')) if x)
            if iletisim:
                veren.append((t('İletişim'), iletisim))
        edilen = [('', girdi.get('musteri', ''))]
        for etiket, alan in (('Yetkili', 'yetkili'), ('Telefon', 'telefon'), ('E-posta', 'eposta'),
                             ('Adres', 'adres')):
            if mb.get(alan):
                edilen.append((t(etiket), mb[alan]))
        if mb.get('vkn') or mb.get('vergi_dairesi'):
            edilen.append((t('VD / VKN'), ' / '.join(x for x in (mb.get('vergi_dairesi'), mb.get('vkn')) if x)))
        B.append(('bilgi', bilgi))
        B.append(('taraflar', veren, edilen, (t('TEKLİFİ VEREN'), t('TEKLİF EDİLEN'))))
    else:
        bilgi.insert(2, (t('Müşteri'), girdi.get('musteri', '')))
        if mb.get('yetkili'):
            bilgi.append((t('Yetkili'), mb['yetkili']))
        B.append(('bilgi', bilgi))
    on_yazi = (T['on_yazi_en'] if en else T['on_yazi']) or ''
    if G['on_yazi'] and on_yazi.strip():
        B.append(('paragraf', teklif_tasarim.yer_tutucu(on_yazi.strip(), degerler)))

    # 1. Lisans
    basliklar, oran = ['#', t('Paket / Modül')], [4, 0]
    if G['kullanici_kolonu']:
        basliklar.append(t('Kapsam'))
        oran.append(12)
    if G['liste_indirim']:
        basliklar += [t('Liste Tutarı'), t('İndirim')]
        oran += [16, 16]
    basliklar.append(t('Net Tutar'))
    oran.append(17)
    oran[1] = 100 - sum(oran)
    sag = set(range(3 if G['kullanici_kolonu'] else 2, len(basliklar)))
    satirlar = []
    for n, x in enumerate(s['satirlar'], 1):
        ad = x['ad'] + ('\n' + x['kod'] if G['paket_kodu'] else '')
        if G['paket_aciklama'] and x.get('aciklama'):
            ad += '\n' + _kisa(x['aciklama'])
        satir = [str(n), ad]
        if G['kullanici_kolonu']:
            satir.append(satir_aciklamasi(x, d))
        if G['liste_indirim']:
            ind = d.p(x['indirim']) if x['indirim'] else '-'
            if x.get('kamp_ad'):
                ind += '\n' + x['kamp_ad']
            satir += [d.p(x['toplam']), ind]
        satirlar.append(satir + [d.p(x['net'])])
    toplam = ['', t('LİSANS TOPLAMI')] + ([''] if G['kullanici_kolonu'] else [])
    if G['liste_indirim']:
        toplam += [d.p(s['lisans']['toplam']), d.p(s['lisans']['indirim']) if s['lisans']['indirim'] else '-']
    satirlar.append(toplam + [d.p(s['lisans']['net'])])
    bol = 1
    B.append(('bolum', '%d. %s' % (bol, t('Lisans'))))
    B.append(('tablo', basliklar, satirlar, {'sag': sag, 'toplam': True, 'oran': oran}))
    dia = s.get('dia') or {}
    if G['dia_ek_bilgi'] and dia:
        ek = [(t(e), d.p(v) + (' + VAT' if en else ' + KDV'))
              for e, v in (('Ek +1 masaüstü kullanıcı bedeli', dia.get('ek_masaustu')),
                           ('Ek +1 mobil kullanıcı bedeli', dia.get('ek_mobil')),
                           ('Yıllık bulut hizmet bedeli', dia.get('yillik_bulut'))) if v]
        if ek:
            B.append(('liste', ek))

    # 2. Uyarlama hizmeti
    u = s['uyarlama']
    kalemler = [k for k in s['hizmet']['kalemler'] if k['saat']]
    gorunum = T['uyarlama_gorunum']
    if kalemler or u['net']:
        bol += 1
        B.append(('bolum', '%d. %s' % (bol, t('Uyarlama Hizmeti (Kurulum, Aktarım, Eğitim ve Danışmanlık)'))))
        aciklama = '' if en else (ayar.get('uyarlama_aciklama') or '').strip()
        if G['uyarlama_aciklama'] and aciklama and gorunum != 'Özet':
            B.append(('metin', aciklama, False))
        if gorunum != 'Özet':
            B.append(('tablo', [t('Hizmet'), t('Kapsam ve hesap dayanağı') if gorunum == 'Detaylı' else t('Kapsam'),
                                t('Süre'), t('Tutar')], uyarlama_satirlari(s, ayar, G, d, gorunum),
                      {'sag': {2, 3}, 'toplam': True, 'oran': [21, 46, 16, 17]}))
        B.append(('vurgu', t('UYARLAMA HİZMET BEDELİ'), '%s + %s' % (d.p(u['net']), 'VAT' if en else 'KDV'),
                  uyarlama_ozeti(s, ayar, G, d, kalemler_dahil=gorunum == 'Özet')))
        haric = '' if en else ((ayar.get('kapsam_haric') or '').strip() if G['kapsam_haric'] else '')
        if u['saat_ucreti'] and G['kapsam_haric']:
            haric += (' ' if haric else '') + ('Additional services hourly rate: %s + VAT.' if en else
                                              'Ek hizmet saat ücreti: %s + KDV.') % d.p(u['saat_ucreti'])
        if haric:
            B.append(('metin', haric, True))

    # Ozet
    oz_gorunum = T['ozet_gorunum']
    if oz_gorunum != 'Gizle':
        bol += 1
        B.append(('bolum', '%d. %s' % (bol, t('Teklif Özeti'))))
        ozet = []
        if oz_gorunum == 'Detaylı':
            ozet = [(t('Lisans'), s['lisans']['net'], False),
                    ('%s (%s)' % (t('Uyarlama Hizmeti'), d.saat(u['saat'])) if u['saat'] else t('Uyarlama Hizmeti'),
                     u['net'], False)]
            ozet += [(k['ad'], k['net'], False) for k in s['sabit']['kalemler']]
            if s['seyahat']['tutar']:
                ozet.append((('Travel and accommodation (%d days on-site)' if en else
                              'Yol / konaklama giderleri (%d gün yerinde)') % s['seyahat']['gun'],
                             s['seyahat']['tutar'], False))
            ozet.append((t('ARA TOPLAM'), s['ara_toplam'], True))
        if s['kampanya']['tutar']:
            ozet.append(('%s (%s, %%%s)' % ('Campaign discount' if en else 'Kampanya İndirimi', s['kampanya']['ad'],
                                            d.g(s['kampanya']['yuzde'])), -s['kampanya']['tutar'], False))
        ozet.append((t('NET TOPLAM (KDV Hariç)'), s['net_toplam'], True))
        if G['ozet_kdv']:
            ozet += [('%s (%%%s)' % (t('KDV'), d.g(s['kdv_orani'])), s['kdv'], False),
                     (t('GENEL TOPLAM (KDV Dahil)'), s['genel_toplam'], True)]
        B.append(('ozet', [(a, d.deger(v), k) for a, v, k in ozet], True, d.birim, d.dil))

    # Odeme secenekleri
    od = [o for o in s['odeme'] if o['musteriye_goster']] if G['odeme'] else []
    if od:
        bol += 1
        B.append(('bolum', '%d. %s' % (bol, t('Ödeme Seçenekleri'))))
        rows = []
        for o in od:
            if o['indirim'] and o['tur'] == 'PESIN':
                notu = ('%%%s advance payment discount' if en else '%%%s peşin indirim') % d.g(o['indirim'])
            elif o['vade_farki']:
                notu = ('%%%s interest' if en else '%%%s vade farkı') % d.g(o['vade_farki'])
            else:
                notu = t('Vade farksız') if o['taksit'] > 1 else ''
            rows.append([_odeme_adi(o['ad'], d), str(o['taksit']) if o['taksit'] > 1 else t('Tek ödeme'),
                         d.p(o['taksit_tutari']), d.p(o['dahil']), notu])
        B.append(('tablo', [t('Ödeme Seçeneği'), t('Taksit'), t('Taksit Tutarı (KDV Dahil)'), t('Toplam (KDV Dahil)'),
                            t('Not')], rows, {'sag': {2, 3}, 'oran': [28, 12, 22, 22, 16]}))

    # Bakim anlasmasi (genelde aylik)
    b = s.get('bakim') or {}
    if G['bakim'] and b.get('sun') and b.get('paketler') and s['satirlar']:
        bol += 1
        B.append(('bolum', '%d. %s' % (bol, t('Bakım ve Destek Anlaşması Seçenekleri'))))
        bg = T['bakim_gorunum']
        aylik = bg != 'Yıllık'
        rows, vurgu = [], set()
        for n, p in enumerate(b['paketler']):
            ad = (_BAKIM_EN.get(p['ad'], p['ad']) if en else p['ad'])
            ad += (' ★ ' + t('Önerilen') if p['onerilen'] else '') + (' ✔ ' + t('Seçilen')
                                                                      if p['ad'] == b.get('secili') else '')
            if p['onerilen'] or p['ad'] == b.get('secili'):
                vurgu.add(n)
            destek = ('%s / %s' % (d.saat(round(p['saat'] / 12, 1)), 'month' if en else 'ay')) if aylik else \
                ('%s / %s' % (d.saat(p['saat']), 'year' if en else 'yıl'))
            satir = [ad, destek, p['yanit'], d.g(p['yerinde_gun']) + (' days/yr' if en else ' gün/yıl'), p['aciklama']]
            if bg == 'Aylık':
                satir.append(d.p(p['aylik']))
            elif bg == 'Yıllık':
                satir.append(d.p(p['yillik']))
            else:
                satir += [d.p(p['aylik']), d.p(p['yillik'])]
            rows.append(satir)
        basliklar = [t('Paket'), t('Destek'), t('Yanıt Süresi'), t('Yerinde'), t('Kapsam')]
        if bg == 'Aylık':
            basliklar.append(t('Aylık (KDV Hr.)'))
            oran, sag = [16, 12, 12, 10, 34, 16], {5}
        elif bg == 'Yıllık':
            basliklar.append(t('Yıllık (KDV Hr.)'))
            oran, sag = [16, 12, 12, 10, 34, 16], {5}
        else:
            basliklar += [t('Aylık (KDV Hr.)'), t('Yıllık (KDV Hr.)')]
            oran, sag = [15, 11, 11, 9, 28, 13, 13], {5, 6}
        B.append(('tablo', basliklar, rows, {'sag': sag, 'oran': oran, 'vurgu': vurgu}))

    # Toplam sahip olma maliyeti
    tco = s.get('tco') or {}
    if G['tco'] and tco.get('yillar') and (s['ybh']['yillik']['net'] or any(y['bakim'] for y in tco['yillar'])):
        bol += 1
        B.append(('bolum', '%d. %d %s' % (bol, len(tco['yillar']), t('Yıllık Toplam Maliyet'))))
        rows = [[('Year %d' if en else '%d. yıl') % y['yil'], d.p(y['tek_seferlik']) if y['tek_seferlik'] else '-',
                 d.p(y['ybh']) if y['ybh'] else '-', d.p(y['bakim']) if y['bakim'] else '-', d.p(y['toplam']),
                 d.p(y['dahil'])] for y in tco['yillar']]
        rows.append([t('TOPLAM'), '', '', '', d.p(tco['toplam']), d.p(tco['dahil'])])
        B.append(('tablo', [t('Yıl'), t('Lisans + Uyarlama'), t('DIA Yıllık Bakım'), t('Destek Anlaşması'),
                            t('Toplam (KDV Hr.)'), t('KDV Dahil')], rows,
                  {'sag': {1, 2, 3, 4, 5}, 'toplam': True, 'oran': [10, 18, 18, 18, 18, 18]}))
        B.append(('metin', t('DIA yıllık bakım bedeli DIA\'nın güncel fiyatlarına göre tahminidir.'), True))

    sartlar = (T['sartlar_en'] if en else T['sartlar']) or ''
    if d.birim != 'TL':
        sartlar += ('\nAmounts are calculated at 1 %s = %s TL; invoicing in TL at the exchange rate of the invoice '
                    'date.' if en else '\nTutarlar 1 %s = %s TL kuru ile hesaplanmıştır; faturalama fatura '
                    'tarihindeki kur üzerinden TL olarak yapılır.') % (d.birim, bicim(d.kur, '', d.dil).strip())
    if G['sartlar'] and sartlar.strip():
        B.append(('bolum', t('Ticari Şartlar')))
        B.append(('maddeler', [x.strip() for x in teklif_tasarim.yer_tutucu(sartlar, degerler).split('\n')
                               if x.strip()]))
    if G['banka'] and (ayar.get('firma_banka') or '').strip():
        B.append(('bolum', t('Banka Bilgileri')))
        B.append(('metin', ayar['firma_banka'].strip(), False))
    notlar = '\n'.join(x for x in [ayar.get('teklif_notlari', ''), girdi.get('notlar', '')] if x)
    if G['notlar'] and notlar:
        B.append(('bolum', t('Notlar')))
        B.append(('metin', notlar, True))
    kapanis = (T['kapanis_en'] if en else T['kapanis']) or ''
    if kapanis.strip():
        B.append(('paragraf', teklif_tasarim.yer_tutucu(kapanis.strip(), degerler)))
    if G['imza']:
        B.append(('imza', [ayar.get('firma_adi', '')] + (_hazirlayan(ayar) if G['hazirlayan'] else [])))
    return B


def karsilastirma_bolumleri(teklifler, ayar, tasarim=None):
    """teklifler: [(girdi, sonuc)] -> alternatif teklif karsilastirmasi."""
    G = teklif_tasarim.tamamla(tasarim)['goster']
    d = _Dil(teklifler[0][0])
    B = [('baslik', 'DİA ERP TEKLİF SEÇENEKLERİ', 'Karşılaştırma')]
    ilk = teklifler[0][0]
    B.append(('bilgi', [('Müşteri', ilk.get('musteri', '')), ('Tarih', datetime.now().strftime('%d.%m.%Y'))]))
    basliklar = [''] + [(g.get('etiket') or 'Seçenek %d' % n) + '\n' + (g.get('teklif_no') or '')
                        for n, (g, _) in enumerate(teklifler, 1)]
    satirlar = [
        ['Paket tipi'] + [s['paket_tipi'] for _, s in teklifler],
        ['Kullanıcı'] + ['%s masaüstü / %s mobil' % (g.get('masaustu', 0), g.get('mobil', 0)) for g, _ in teklifler],
        ['Paketler'] + ['\n'.join(x['ad'] for x in s['satirlar']) for _, s in teklifler],
        ['Lisans'] + [d.p(s['lisans']['net']) for _, s in teklifler],
        ['Uyarlama'] + ['%s\n(%s saat)' % (d.p(s['uyarlama']['net']), _g(s['uyarlama']['saat'])) for _, s in teklifler],
        ['Diğer (sabit, yol)'] + [d.p(s['sabit']['net'] + s['seyahat']['tutar']) for _, s in teklifler],
        ['Kampanya indirimi'] + [d.p(-s['kampanya']['tutar']) if s['kampanya']['tutar'] else '-' for _, s in teklifler],
        ['NET TOPLAM (KDV Hariç)'] + [d.p(s['net_toplam']) for _, s in teklifler],
        ['GENEL TOPLAM (KDV Dahil)'] + [d.p(s['genel_toplam']) for _, s in teklifler],
    ]
    if all(s.get('tco') for _, s in teklifler):
        satirlar.append(['%d yıllık toplam maliyet' % len(teklifler[0][1]['tco']['yillar'])] +
                        [d.p(s['tco']['dahil']) for _, s in teklifler])
    oran = [24] + [76 // len(teklifler)] * len(teklifler)
    B.append(('tablo', basliklar, satirlar, {'sag': set(range(1, len(teklifler) + 1)), 'oran': oran,
                                            'vurgu': {7, 8}}))
    B.append(('metin', 'Her seçeneğin ayrıntılı teklifi ayrıca sunulmuştur. Tutarlar %s tarihlidir.' %
              datetime.now().strftime('%d.%m.%Y'), True))
    if G['imza']:
        B.append(('imza', [ayar.get('firma_adi', '')] + (_hazirlayan(ayar) if G['hazirlayan'] else [])))
    return B


# ======================================================================
# PDF
# ======================================================================
def _fontlar():
    """Turkce harfleri basabilen TTF kaydeder. Bulamazsa Helvetica + Turkce harfsiz metin."""
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    for normal, kalin in _FONT_ADAYLARI:
        if os.path.exists(normal) and os.path.exists(kalin):
            try:
                pdfmetrics.registerFont(TTFont('TeklifFont', normal))
                pdfmetrics.registerFont(TTFont('TeklifFont-Kalin', kalin))
                return 'TeklifFont', 'TeklifFont-Kalin', None
            except Exception:
                continue
    return 'Helvetica', 'Helvetica-Bold', _ASCII


def _pdf_ciz(yol, B, ayar, belge_no, tasarim=None, dil='TR'):
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.pdfgen import canvas as rl_canvas
    from reportlab.platypus import (Image, KeepTogether, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table,
                                    TableStyle)

    T = teklif_tasarim.tamamla(tasarim)
    G, sablon, fs = T['goster'], T['sablon'], float(T['yazi_boyutu'])
    sade, modern, klasik = sablon == 'Sade', sablon == 'Modern', sablon == 'Klasik'
    kapakli, seritli = sablon == 'Kapaklı', sablon == 'Şeritli'

    def tt(m):
        return _EN.get(m, m) if dil == 'EN' else m
    font, kalin, cevir = _fontlar()
    renk = colors.HexColor
    ana, vurgu = renk(T['ana_renk']), renk(T['vurgu_renk'])
    ana_acik = renk(teklif_tasarim.acik_ton(T['ana_renk'], 0.88))
    ana_cok_acik = renk(teklif_tasarim.acik_ton(T['ana_renk'], 0.95))
    vurgu_acik = renk(teklif_tasarim.acik_ton(T['vurgu_renk'], 0.8))
    gri, gri_yazi, beyaz = renk('#cccccc'), renk('#555555'), colors.white
    tablo_bas_zemin, tablo_bas_yazi = (renk('#eeeeee'), colors.black) if sade else (ana, beyaz)
    sol_bosluk = 24 * mm if seritli else 16 * mm
    genislik = A4[0] - sol_bosluk - 16 * mm

    def t(metin):
        metin = str(metin)
        return metin.translate(cevir) if cevir else metin

    def P(metin, boyut=None, kalin_mi=False, hiza=None, rnk=colors.black):
        boyut = boyut or fs
        stil = ParagraphStyle('p', fontName=kalin if kalin_mi else font, fontSize=boyut, leading=boyut * 1.32,
                              textColor=rnk, alignment=hiza if hiza is not None else 0)
        return Paragraph(t(metin).replace('&', '&amp;').replace('<', '&lt;').replace('\n', '<br/>'), stil)

    def logo_resmi(gen, yuk):
        logo = ayar.get('logo_yolu')
        if logo and os.path.exists(logo):
            try:
                resim = Image(logo)
                oran = min(gen / resim.imageWidth, yuk / resim.imageHeight)
                resim.drawWidth, resim.drawHeight = resim.imageWidth * oran, resim.imageHeight * oran
                return resim
            except Exception:
                return None
        return None

    akis = []
    firma_adi = ayar.get('firma_adi', '')
    d_ = _Dil({'dil': dil})
    satirlar = firma_satirlari(ayar, d_)

    # ---------------------------------------------------------------- kapak (Kapakli)
    kapak = next((b for b in B if b[0] == 'kapak'), None)
    if kapakli and kapak:
        _, baslik, alt, musteri, bilgiler = kapak
        resim = logo_resmi(90 * mm, 40 * mm)
        akis += [Spacer(1, 30 * mm)]
        if resim:
            resim.hAlign = 'CENTER'
            akis += [resim, Spacer(1, 10 * mm)]
        akis += [P(firma_adi, fs + 8, True, TA_CENTER, ana), Spacer(1, 25 * mm),
                 Table([['']], colWidths=[genislik * 0.4], rowHeights=[2], style=[
                     ('BACKGROUND', (0, 0), (-1, -1), vurgu)]), Spacer(1, 8 * mm),
                 P(baslik, fs + 9, True, TA_CENTER, ana)]
        if alt:
            akis.append(P(alt, fs + 2, hiza=TA_CENTER, rnk=vurgu))
        akis += [Spacer(1, 14 * mm), P(musteri, fs + 7, True, TA_CENTER), Spacer(1, 18 * mm)]
        tb = Table([[P(e, fs + 1, True, TA_RIGHT, gri_yazi), P(v, fs + 1)] for e, v in bilgiler],
                   colWidths=[genislik * 0.45, genislik * 0.55])
        tb.setStyle(TableStyle([('TOPPADDING', (0, 0), (-1, -1), 3)]))
        akis += [tb, PageBreak()]

    # ---------------------------------------------------------------- ust bilgi (logo + firma)
    if modern:
        yazi = [P(firma_adi, fs + 7, True, rnk=beyaz)] + [P(x, fs - 1, rnk=renk('#e8eef2')) for x in satirlar]
        resim = logo_resmi(48 * mm, 20 * mm)
        ust = Table([[resim, yazi]] if resim else [[yazi]],
                    colWidths=[56 * mm, genislik - 56 * mm] if resim else [genislik])
        stil = [('BACKGROUND', (0, 0), (-1, -1), ana), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('TOPPADDING', (0, 0), (-1, -1), 8), ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
                ('LEFTPADDING', (0, 0), (-1, -1), 8)]
        if resim:
            stil += [('BACKGROUND', (0, 0), (0, 0), beyaz), ('ALIGN', (0, 0), (0, 0), 'CENTER')]
        ust.setStyle(TableStyle(stil))
        akis += [ust, Spacer(1, 2 * mm)]
    elif klasik:
        resim = logo_resmi(60 * mm, 22 * mm)
        if resim:
            resim.hAlign = 'CENTER'
            akis.append(resim)
        akis.append(P(firma_adi, fs + 6, True, TA_CENTER, ana))
        akis += [P(x, fs - 1.5, hiza=TA_CENTER, rnk=gri_yazi) for x in satirlar]
        akis += [Spacer(1, 2 * mm),
                 Table([['']], colWidths=[genislik], rowHeights=[1.6], style=[('BACKGROUND', (0, 0), (-1, -1), ana)]),
                 Spacer(1, 0.8 * mm),
                 Table([['']], colWidths=[genislik], rowHeights=[0.6], style=[('BACKGROUND', (0, 0), (-1, -1), vurgu)]),
                 Spacer(1, 5 * mm)]
    else:
        yazi = [P(firma_adi, fs + 6, True, TA_RIGHT, colors.black if sade else ana)] + \
               [P(x, fs - 1.5, hiza=TA_RIGHT, rnk=gri_yazi) for x in satirlar]
        resim = logo_resmi(55 * mm, 22 * mm)
        ust = Table([[resim or '', yazi]], colWidths=[62 * mm, genislik - 62 * mm])
        cizgi = [('LINEBELOW', (0, 0), (-1, 0), 0.6, gri)] if sade else [('LINEBELOW', (0, 0), (-1, 0), 2, ana)]
        ust.setStyle(TableStyle([('VALIGN', (0, 0), (-1, -1), 'MIDDLE'), ('LEFTPADDING', (0, 0), (-1, -1), 0),
                                 ('RIGHTPADDING', (0, 0), (-1, -1), 0), ('BOTTOMPADDING', (0, 0), (-1, -1), 6)]
                                + cizgi))
        akis.append(ust)
        if not sade:
            akis.append(Table([['']], colWidths=[genislik], rowHeights=[1.2],
                              style=[('BACKGROUND', (0, 0), (-1, -1), vurgu)]))
        akis.append(Spacer(1, 4 * mm))

    def bolum_basligi(metin):
        if modern:
            tb = Table([['', P(metin, fs + 3, True, rnk=ana)]], colWidths=[2.5 * mm, genislik - 2.5 * mm])
            tb.setStyle(TableStyle([('BACKGROUND', (0, 0), (0, 0), vurgu), ('LEFTPADDING', (1, 0), (1, 0), 6),
                                    ('TOPPADDING', (0, 0), (-1, -1), 2), ('BOTTOMPADDING', (0, 0), (-1, -1), 2)]))
        elif klasik:
            tb = Table([[P(metin.upper(), fs + 1.5, True, rnk=ana)]], colWidths=[genislik])
            tb.setStyle(TableStyle([('LINEABOVE', (0, 0), (-1, -1), 0.6, vurgu),
                                    ('LINEBELOW', (0, 0), (-1, -1), 0.6, vurgu),
                                    ('LEFTPADDING', (0, 0), (-1, -1), 0), ('TOPPADDING', (0, 0), (-1, -1), 3),
                                    ('BOTTOMPADDING', (0, 0), (-1, -1), 3)]))
        elif seritli:
            tb = Table([[P(metin, fs + 3, True, rnk=ana)], ['']], colWidths=[genislik], rowHeights=[None, 1.6])
            tb.setStyle(TableStyle([('BACKGROUND', (0, 1), (0, 1), vurgu), ('LEFTPADDING', (0, 0), (-1, -1), 0),
                                    ('BOTTOMPADDING', (0, 0), (-1, 0), 2), ('TOPPADDING', (0, 1), (-1, 1), 0),
                                    ('BOTTOMPADDING', (0, 1), (-1, 1), 0), ('RIGHTPADDING', (0, 1), (-1, 1),
                                                                            genislik - 30 * mm)]))
        else:
            tb = Table([[P(metin, fs + (2 if sade else 3), True, rnk=colors.black if sade else ana)]],
                       colWidths=[genislik])
            tb.setStyle(TableStyle([('LINEBELOW', (0, 0), (-1, -1), 0.8 if sade else 1.2, gri if sade else vurgu),
                                    ('LEFTPADDING', (0, 0), (-1, -1), 0), ('BOTTOMPADDING', (0, 0), (-1, -1), 3)]))
        return [tb, Spacer(1, 2 * mm)]

    def kutu(baslik, satirlar_, gen):
        veri = [[P(baslik, fs, True, rnk=tablo_bas_yazi), '']]
        for e, d in satirlar_:
            if e:
                veri.append([P(e, fs - 0.5, True, rnk=gri_yazi), P(d, fs - 0.5)])
            else:
                veri.append([P(d, fs + (1 if len(veri) == 1 else -0.5), len(veri) == 1), ''])
        tb = Table(veri, colWidths=[24 * mm, gen - 24 * mm])
        stil = [('BACKGROUND', (0, 0), (-1, 0), tablo_bas_zemin), ('SPAN', (0, 0), (1, 0)),
                ('BOX', (0, 0), (-1, -1), 0.5, gri), ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('BACKGROUND', (0, 1), (-1, -1), beyaz if sade else ana_cok_acik),
                ('TOPPADDING', (0, 1), (-1, -1), 1.5), ('BOTTOMPADDING', (0, 1), (-1, -1), 1.5)]
        for r, (e, _) in enumerate(satirlar_, 1):
            if not e:
                stil.append(('SPAN', (0, r), (1, r)))
        tb.setStyle(TableStyle(stil))
        return tb

    bekleyen = None      # bolum basligi bir sonraki icerikle birlikte tutulur (sayfa sonunda yalniz kalmasin)

    def ekle(parcalar, bosluk=4):
        nonlocal bekleyen
        akis.append(KeepTogether((bekleyen or []) + parcalar))
        bekleyen = None
        akis.append(Spacer(1, bosluk * mm))

    for b in B:
        tur = b[0]
        if tur == 'baslik':
            if modern:
                hucre = [P(b[1], fs + 5, True, rnk=beyaz)] + ([P(b[2], fs, rnk=renk('#fff3e0'))] if b[2] else [])
                akis.append(Table([[hucre]], colWidths=[genislik], style=[
                    ('BACKGROUND', (0, 0), (-1, -1), vurgu), ('TOPPADDING', (0, 0), (-1, -1), 6),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 6), ('LEFTPADDING', (0, 0), (-1, -1), 8)]))
            else:
                hiza = TA_CENTER if klasik else None
                akis.append(P(b[1], fs + 7, True, hiza, colors.black if sade else ana))
                if b[2]:
                    akis.append(P(b[2], fs + 1, hiza=hiza, rnk=gri_yazi if sade else vurgu))
            akis.append(Spacer(1, 3 * mm))
        elif tur == 'bilgi':
            ciftler = b[1]
            n = 3 if len(ciftler) > 4 else 2
            hucreler = []
            for i in range(0, len(ciftler), n):
                grup = ciftler[i:i + n] + [('', '')] * (n - len(ciftler[i:i + n]))
                satir = []
                for e, d in grup:
                    satir += [P(e, fs - 0.5, True, rnk=gri_yazi), P(d, fs - 0.5)]
                hucreler.append(satir)
            etiket_gen = 22 * mm
            tb = Table(hucreler, colWidths=[etiket_gen, (genislik - n * etiket_gen) / n] * n)
            stil = [('GRID', (0, 0), (-1, -1), 0.4, gri), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')]
            for c in range(0, 2 * n, 2):
                stil.append(('BACKGROUND', (c, 0), (c, -1), renk('#f4f4f4') if sade else ana_acik))
            tb.setStyle(TableStyle(stil))
            akis += [tb, Spacer(1, 3 * mm)]
        elif tur == 'taraflar':
            yarim = (genislik - 6 * mm) / 2
            bas = b[3] if len(b) > 3 else ('TEKLİFİ VEREN', 'TEKLİF EDİLEN')
            tb = Table([[kutu(bas[0], b[1], yarim), '', kutu(bas[1], b[2], yarim)]],
                       colWidths=[yarim, 6 * mm, yarim])
            tb.setStyle(TableStyle([('VALIGN', (0, 0), (-1, -1), 'TOP'), ('LEFTPADDING', (0, 0), (-1, -1), 0),
                                    ('RIGHTPADDING', (0, 0), (-1, -1), 0)]))
            akis += [tb, Spacer(1, 4 * mm)]
        elif tur == 'paragraf':
            ekle([P(b[1], fs + 0.5)])
        elif tur == 'bolum':
            bekleyen = bolum_basligi(b[1])
        elif tur == 'tablo':
            _, basliklar, satirlar_, sec = b
            sag = sec.get('sag', set())
            oranlar = sec.get('oran') or [1] * len(basliklar)
            kolon = [genislik * o / sum(oranlar) for o in oranlar]
            veri = [[P(h, fs - 1, True, TA_RIGHT if i in sag else None, tablo_bas_yazi)
                     for i, h in enumerate(basliklar)]]
            for r, satir in enumerate(satirlar_):
                son = sec.get('toplam') and r == len(satirlar_) - 1
                veri.append([P(v, fs - 1, son, TA_RIGHT if i in sag else None) for i, v in enumerate(satir)])
            tb = Table(veri, colWidths=kolon, repeatRows=1)
            stil = [('BACKGROUND', (0, 0), (-1, 0), tablo_bas_zemin), ('LINEBELOW', (0, 0), (-1, -1), 0.3, gri),
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                    ('ROWBACKGROUNDS', (0, 1), (-1, -1), [beyaz, renk('#fafafa') if sade else ana_cok_acik])]
            if sade:
                stil.append(('LINEABOVE', (0, 0), (-1, 0), 0.8, colors.black))
            if klasik:
                stil.append(('LINEBELOW', (0, 0), (-1, 0), 1.2, vurgu))
            if sec.get('toplam'):
                stil.append(('BACKGROUND', (0, -1), (-1, -1), renk('#eeeeee') if sade else ana_acik))
            for r in sec.get('vurgu', ()):
                stil.append(('BACKGROUND', (0, r + 1), (-1, r + 1), vurgu_acik))
            tb.setStyle(TableStyle(stil))
            if len(satirlar_) < 25:
                ekle([tb])
            else:
                if bekleyen:
                    akis.append(KeepTogether(bekleyen))
                    bekleyen = None
                akis += [tb, Spacer(1, 4 * mm)]
        elif tur == 'liste':
            veri = [[P(e, fs - 0.5, rnk=gri_yazi), P(d, fs - 0.5, True, TA_RIGHT)] for e, d in b[1]]
            tb = Table(veri, colWidths=[genislik * 0.6, genislik * 0.4])
            tb.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, -1), renk('#f7f7f7') if sade else ana_cok_acik),
                                    ('BOX', (0, 0), (-1, -1), 0.4, gri),
                                    ('TOPPADDING', (0, 0), (-1, -1), 2), ('BOTTOMPADDING', (0, 0), (-1, -1), 2)]))
            if akis and isinstance(akis[-1], Spacer):
                akis.pop()
                akis.append(Spacer(1, 1.5 * mm))
            akis += [tb, Spacer(1, 4 * mm)]
        elif tur == 'vurgu':
            alt_renk = renk('#dddddd') if sade else vurgu
            kutu_ = Table([[P(b[1], fs + 2, True, rnk=beyaz), P(b[2], fs + 5, True, TA_RIGHT, beyaz)],
                           [P(b[3], fs - 1, rnk=alt_renk), '']], colWidths=[genislik - 70 * mm, 70 * mm])
            kutu_.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, -1), ana), ('SPAN', (0, 1), (1, 1)),
                                       ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'), ('TOPPADDING', (0, 0), (-1, 0), 6),
                                       ('BOTTOMPADDING', (0, 1), (-1, 1), 6), ('LEFTPADDING', (0, 0), (-1, -1), 8),
                                       ('RIGHTPADDING', (0, 0), (-1, -1), 8)]))
            if bekleyen:          # Ozet gorunumde bolum basligi dogrudan kutuya baglanir
                ekle([kutu_], 3)
            else:
                if akis and isinstance(akis[-1], Spacer):
                    akis.pop()
                akis += [Spacer(1, 1 * mm), KeepTogether([kutu_]), Spacer(1, 3 * mm)]
        elif tur == 'ozet':
            birim, odil = (b[3], b[4]) if len(b) > 4 else ('TL', 'TR')
            veri = [[P(a, fs, k), P(bicim(v, birim, odil), fs, k, TA_RIGHT)] for a, v, k in b[1]]
            if b[2]:
                a, v, _ = b[1][-1]
                veri[-1] = [P(a, fs + 1, True, rnk=beyaz), P(bicim(v, birim, odil), fs + 1, True, TA_RIGHT, beyaz)]
            tb = Table(veri, colWidths=[None, 52 * mm])
            stil = [('LINEBELOW', (0, 0), (-1, -1), 0.3, gri)]
            if b[2]:
                stil.append(('BACKGROUND', (0, -1), (-1, -1), ana))
            tb.setStyle(TableStyle(stil))
            ekle([tb])
        elif tur == 'metin':
            ekle([P(b[1], fs - 1 if b[2] else fs, rnk=renk('#444444'))])
        elif tur == 'maddeler':
            ekle([P('•  ' + x, fs - 0.5, rnk=renk('#333333')) for x in b[1]])
        elif tur in ('imza', 'onay'):
            veren = b[1] if tur == 'imza' else [firma_adi]
            yarim = (genislik - 10 * mm) / 2
            baslik_rengi = colors.black if sade else ana
            sol = [P(tt('TEKLİFİ VEREN'), fs + 1, True, rnk=baslik_rengi)] + \
                  [P(x, fs, i == 0) for i, x in enumerate(veren) if x] + \
                  [Spacer(1, 10 * mm), P(tt('İmza') + ': _______________________', fs)]
            sag_ = [P(tt('MÜŞTERİ ONAYI'), fs + 1, True, rnk=baslik_rengi),
                    P(tt('Firma Kaşe / İmza') + ': _______________________', fs), Spacer(1, 3 * mm),
                    P(tt('Yetkili Ad Soyad') + ': _______________________', fs), Spacer(1, 3 * mm),
                    P(tt('Tarih') + ': _______________________', fs)]
            tb = Table([[sol, '', sag_]], colWidths=[yarim, 10 * mm, yarim])
            tb.setStyle(TableStyle([('VALIGN', (0, 0), (-1, -1), 'TOP'),
                                    ('LINEABOVE', (0, 0), (0, 0), 0.8, gri), ('LINEABOVE', (2, 0), (2, 0), 0.8, gri),
                                    ('TOPPADDING', (0, 0), (-1, -1), 6)]))
            akis += [Spacer(1, 4 * mm), KeepTogether([tb])]

    alt_metin = '  ·  '.join(x for x in (firma_adi, ayar.get('firma_telefon'), ayar.get('firma_eposta'),
                                         ayar.get('firma_web')) if x)

    def sayfa_cizimi(tuval, sayfa, toplam):
        tuval.saveState()
        if seritli:          # her sayfada solda renkli serit
            tuval.setFillColor(ana)
            tuval.rect(0, 0, 9 * mm, A4[1], stroke=0, fill=1)
            tuval.setFillColor(vurgu)
            tuval.rect(9 * mm, 0, 1.5 * mm, A4[1], stroke=0, fill=1)
        if kapakli and sayfa == 1:
            tuval.setFillColor(ana)
            tuval.rect(0, 0, A4[0], 14 * mm, stroke=0, fill=1)
            tuval.setFillColor(vurgu)
            tuval.rect(0, 14 * mm, A4[0], 1.5 * mm, stroke=0, fill=1)
            tuval.setFont(font, 8)
            tuval.setFillColor(beyaz)
            tuval.drawCentredString(A4[0] / 2, 5.5 * mm, t(alt_metin))
        elif G['alt_bilgi']:
            tuval.setStrokeColor(gri if sade else ana)
            tuval.setLineWidth(0.6)
            tuval.line(sol_bosluk, 12.5 * mm, A4[0] - 16 * mm, 12.5 * mm)
            tuval.setFont(font, 7)
            tuval.setFillColor(renk('#777777'))
            tuval.drawString(sol_bosluk, 8.5 * mm, t(alt_metin))
            tuval.drawRightString(A4[0] - 16 * mm, 8.5 * mm, t('%s  ·  %s %d / %d' % (belge_no, tt('Sayfa'), sayfa,
                                                                                       toplam)))
        tuval.restoreState()

    class SayfaliTuval(rl_canvas.Canvas):
        """'Sayfa x / y' icin alt bilgi tum sayfalar bitince cizilir."""
        def __init__(self, *a, **k):
            super().__init__(*a, **k)
            self._sayfalar = []

        def showPage(self):
            self._sayfalar.append(dict(self.__dict__))
            self._startPage()

        def save(self):
            toplam = len(self._sayfalar)
            for n, durum in enumerate(self._sayfalar, 1):
                self.__dict__.update(durum)
                sayfa_cizimi(self, n, toplam)
                super().showPage()
            super().save()

    belge = SimpleDocTemplate(yol, pagesize=A4, leftMargin=sol_bosluk, rightMargin=16 * mm, topMargin=14 * mm,
                              bottomMargin=18 * mm, title=t(belge_no), author=t(firma_adi))
    belge.build(akis, canvasmaker=SayfaliTuval)
    return yol


# ======================================================================
# Excel
# ======================================================================
def _excel_ciz(yol, B, ayar, sayfa_adi='Teklif', tasarim=None, dil='TR'):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    T = teklif_tasarim.tamamla(tasarim)
    G, sade = T['goster'], T['sablon'] == 'Sade'

    def tt(m):
        return _EN.get(m, m) if dil == 'EN' else m
    ana, vurgu = T['ana_renk'][1:], T['vurgu_renk'][1:]
    wb = Workbook()
    ws = wb.active
    ws.title = sayfa_adi
    ws.sheet_view.showGridLines = False

    def dolgu(r):
        return PatternFill('solid', start_color=r)
    ana_f = dolgu(ana)
    acik_f = dolgu(teklif_tasarim.acik_ton(T['ana_renk'], 0.88)[1:])
    cok_acik_f = dolgu(teklif_tasarim.acik_ton(T['ana_renk'], 0.95)[1:])
    vurgu_f = dolgu(teklif_tasarim.acik_ton(T['vurgu_renk'], 0.8)[1:])
    bas_f, bas_font = (dolgu('EEEEEE'), Font(bold=True)) if sade else (ana_f, Font(color='FFFFFF', bold=True))
    beyaz = Font(color='FFFFFF', bold=True)
    ince = Border(bottom=Side(style='thin', color='CCCCCC'))
    kaydir = Alignment(wrap_text=True, vertical='top')
    sag = Alignment(horizontal='right', vertical='top', wrap_text=True)
    genislikler = (26, 30, 16, 18, 22, 18, 16)
    for kol, gen in zip('ABCDEFG', genislikler):
        ws.column_dimensions[kol].width = gen

    def birlesik(r, metin, font=None, c1=1, c2=7, hiza=None, dolgu_=None):
        ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
        h = ws.cell(r, c1, metin)
        h.alignment = hiza or Alignment(wrap_text=True, vertical='top')
        if font:
            h.font = font
        if dolgu_:
            for c in range(c1, c2 + 1):
                ws.cell(r, c).fill = dolgu_
        satir_say = sum(max(1, math.ceil(len(x) / (sum(genislikler[c1 - 1:c2]) * 1.1)))
                        for x in str(metin).split('\n'))
        if satir_say > 1:
            ws.row_dimensions[r].height = 15 * satir_say
        return h

    # Ust bilgi: logo + firma
    logo = ayar.get('logo_yolu')
    logo_var = False
    if logo and os.path.exists(logo):
        try:
            from openpyxl.drawing.image import Image as XLImage
            resim = XLImage(logo)
            oran = min(220 / resim.width, 70 / resim.height)
            resim.width, resim.height = int(resim.width * oran), int(resim.height * oran)
            ws.add_image(resim, 'A1')
            logo_var = True
        except Exception:
            logo_var = False
    ad_font = Font(size=16, bold=True, color='000000' if sade else ana)
    firma_s = firma_satirlari(ayar, _Dil({'dil': dil}))
    if logo_var:
        birlesik(1, ayar.get('firma_adi', ''), ad_font, 4, 7, Alignment(horizontal='right'))
        for i, x in enumerate(firma_s[:3], 2):
            birlesik(i, x, Font(size=9, color='555555'), 4, 7, Alignment(horizontal='right'))
        r = 5
    else:
        ws.cell(1, 1, ayar.get('firma_adi', '')).font = ad_font
        r = 2
        for x in firma_s:
            ws.cell(r, 1, x).font = Font(size=9, color='555555')
            r += 1
    for c in range(1, 8):
        ws.cell(r, c).border = Border(bottom=Side(style='medium', color='999999' if sade else ana))
    r += 2

    for b in B:
        tur = b[0]
        if tur == 'kapak':
            continue
        if tur == 'baslik':
            zemin = None if sade else (dolgu(vurgu) if T['sablon'] == 'Modern' else ana_f)
            birlesik(r, b[1] + ('  -  ' + b[2] if b[2] else ''),
                     Font(size=14, bold=True, color='000000' if sade else 'FFFFFF'), dolgu_=zemin)
            r += 2
        elif tur == 'bilgi':
            for i in range(0, len(b[1]), 2):
                for j, (e, d) in enumerate(b[1][i:i + 2]):
                    h = ws.cell(r, 1 + j * 3, e)
                    h.font, h.fill = Font(bold=True), acik_f
                    ws.cell(r, 2 + j * 3, d)
                r += 1
            r += 1
        elif tur == 'taraflar':
            bas = b[3] if len(b) > 3 else ('TEKLİFİ VEREN', 'TEKLİF EDİLEN')
            for c1, baslik, satirlar_ in ((1, bas[0], b[1]), (5, bas[1], b[2])):
                rr = r
                ws.merge_cells(start_row=rr, start_column=c1, end_row=rr, end_column=c1 + 2)
                ws.cell(rr, c1, baslik).font = bas_font
                for c in range(c1, c1 + 3):
                    ws.cell(rr, c).fill = bas_f
                for e, d in satirlar_:
                    rr += 1
                    if e:
                        ws.cell(rr, c1, e).font = Font(bold=True, color='555555')
                        ws.cell(rr, c1 + 1, d)
                    else:
                        ws.cell(rr, c1, d).font = Font(bold=rr == r + 1)
                    for c in range(c1, c1 + 3):
                        ws.cell(rr, c).fill = cok_acik_f
            r += max(len(b[1]), len(b[2])) + 2
        elif tur == 'paragraf':
            birlesik(r, b[1], Font(size=10))
            r += 2
        elif tur == 'bolum':
            ws.cell(r, 1, b[1]).font = Font(size=12, bold=True, color='000000' if sade else ana)
            for c in range(1, 8):
                ws.cell(r, c).border = Border(bottom=Side(style='thin', color='999999' if sade else vurgu))
            r += 1
        elif tur == 'tablo':
            _, basliklar, satirlar_, sec = b
            for c, hb in enumerate(basliklar, 1):
                hc = ws.cell(r, c, hb)
                hc.font, hc.fill, hc.alignment = bas_font, bas_f, kaydir
            r += 1
            for n, satir in enumerate(satirlar_):
                son = sec.get('toplam') and n == len(satirlar_) - 1
                for c, v in enumerate(satir, 1):
                    hc = ws.cell(r, c, v)
                    hc.border = ince
                    hc.alignment = sag if (c - 1) in sec.get('sag', ()) else kaydir
                    if son:
                        hc.font, hc.fill = Font(bold=True), acik_f
                    elif n in sec.get('vurgu', ()):
                        hc.fill = vurgu_f
                    elif n % 2:
                        hc.fill = cok_acik_f
                r += 1
            r += 1
        elif tur == 'liste':
            r -= 1
            for e, d in b[1]:
                ws.cell(r, 2, e).font = Font(color='555555')
                h = ws.cell(r, 4, d)
                h.font, h.alignment = Font(bold=True), Alignment(horizontal='right')
                for c in range(2, 5):
                    ws.cell(r, c).fill = cok_acik_f
                r += 1
            r += 1
        elif tur == 'vurgu':
            for c in range(1, 5):
                ws.cell(r, c).fill = ana_f
                ws.cell(r + 1, c).fill = ana_f
            ws.cell(r, 1, b[1]).font = Font(size=12, bold=True, color='FFFFFF')
            hc = ws.cell(r, 4, b[2])
            hc.font, hc.alignment = Font(size=13, bold=True, color='FFFFFF'), Alignment(horizontal='right')
            ws.cell(r + 1, 1, b[3]).font = Font(size=9, color='DDDDDD' if sade else vurgu)
            r += 3
        elif tur == 'ozet':
            birim = b[3] if len(b) > 3 else 'TL'
            for n, (a, v, k) in enumerate(b[1]):
                ws.cell(r, 1, a).font = Font(bold=k)
                hc = ws.cell(r, 4, round(v, 2))
                hc.number_format = '#,##0.00 "%s"' % birim
                hc.font = Font(bold=k)
                if b[2] and n == len(b[1]) - 1:
                    for c in range(1, 5):
                        ws.cell(r, c).fill = ana_f
                    ws.cell(r, 1).font = ws.cell(r, 4).font = beyaz
                r += 1
            r += 1
        elif tur == 'metin':
            birlesik(r, b[1], Font(size=9 if b[2] else 10, color='444444'))
            r += 2
        elif tur == 'maddeler':
            for x in b[1]:
                birlesik(r, '•  ' + x, Font(size=10))
                r += 1
            r += 1
        elif tur in ('imza', 'onay'):
            veren = b[1] if tur == 'imza' else [ayar.get('firma_adi', '')]
            renk_ = '000000' if sade else ana
            ws.cell(r, 1, tt('TEKLİFİ VEREN')).font = Font(bold=True, color=renk_)
            ws.cell(r, 5, tt('MÜŞTERİ ONAYI')).font = Font(bold=True, color=renk_)
            sol = [x for x in veren if x] + ['', tt('İmza') + ': ______________________']
            sag_ = [tt('Firma Kaşe / İmza') + ': ______________________', '',
                    tt('Yetkili Ad Soyad') + ': ______________________', '', tt('Tarih') + ': ______________________']
            for i in range(max(len(sol), len(sag_))):
                if i < len(sol):
                    ws.cell(r + 1 + i, 1, sol[i]).font = Font(bold=i == 0)
                if i < len(sag_):
                    ws.cell(r + 1 + i, 5, sag_[i])
            r += max(len(sol), len(sag_)) + 2
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins.left = ws.page_margins.right = 0.5
    if G['alt_bilgi']:
        ws.oddFooter.left.text = '  ·  '.join(x for x in (ayar.get('firma_adi'), ayar.get('firma_telefon'),
                                                          ayar.get('firma_web')) if x).replace('&', '&&')
        ws.oddFooter.left.size = 8
        ws.oddFooter.right.text = '%s &P / &N' % tt('Sayfa')
        ws.oddFooter.right.size = 8
    wb.save(yol)
    return yol


# ======================================================================
# Dis arayuz
# ======================================================================
def _dil(girdi):
    return _Dil(girdi).dil


def pdf_olustur(yol, girdi, sonuc, ayar, tasarim=None):
    no = '%s %s' % ('Proposal' if _dil(girdi) == 'EN' else 'Teklif', girdi.get('teklif_no') or '')
    return _pdf_ciz(yol, bolumler(girdi, sonuc, ayar, tasarim), ayar, no, tasarim, _dil(girdi))


def excel_olustur(yol, girdi, sonuc, ayar, tasarim=None):
    return _excel_ciz(yol, bolumler(girdi, sonuc, ayar, tasarim), ayar, 'Proposal' if _dil(girdi) == 'EN'
                      else 'Teklif', tasarim, _dil(girdi))


def karsilastirma_pdf(yol, teklifler, ayar, tasarim=None):
    return _pdf_ciz(yol, karsilastirma_bolumleri(teklifler, ayar, tasarim), ayar, 'Teklif Seçenekleri', tasarim)


def karsilastirma_excel(yol, teklifler, ayar, tasarim=None):
    return _excel_ciz(yol, karsilastirma_bolumleri(teklifler, ayar, tasarim), ayar, 'Seçenekler', tasarim)


def ybh_excel(yol, musteri, tip, sonuc, ayar):
    from openpyxl import Workbook
    from openpyxl.styles import Font
    wb = Workbook()
    ws = wb.active
    ws.title = 'Yıllık Bakım'
    ws.append([ayar.get('firma_adi', '')])
    ws.append(['YILLIK BAKIM (YBH) HESABI', '', 'Müşteri:', musteri, 'Paket:', tip])
    ws.append([])
    ws.append(['Paket Kodu', 'Paket Adı', 'Kullanıcı', 'Modül', 'Lisans Bedeli', 'YBH %', 'Baz Bakım'])
    for s in sonuc['satirlar']:
        ws.append([s['kod'], s['ad'], s['kullanici'], s['modul'], s['lisans'], s['ybh_yuzde'], s['baz']])
    ws.append([])
    for a, b in [('Toplam Baz Bakım', sonuc['baz']), ('Modül Katsayısı (%d modül)' % sonuc['modul'],
                                                       sonuc['modul_katsayi']),
                 ('Kullanıcı Katsayısı (%d kullanıcı)' % sonuc['kullanici'], sonuc['kullanici_katsayi']),
                 ('Toplam Katsayı', sonuc['katsayi']), ('YILLIK BAKIM (KDV Hariç)', sonuc['yillik']['net']),
                 ('KDV', sonuc['yillik']['kdv']), ('KDV DAHİL TOPLAM', sonuc['yillik']['dahil'])]:
        ws.append(['', a, '', '', '', '', b])
    for r in ws.iter_rows(min_row=1, max_row=4):
        for h in r:
            h.font = Font(bold=True)
    for kol, gen in zip('ABCDEFG', (22, 40, 10, 8, 16, 8, 16)):
        ws.column_dimensions[kol].width = gen
    wb.save(yol)
    return yol
