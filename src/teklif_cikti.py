"""Musteriye gidecek teklif ciktisi: PDF (reportlab) ve Excel (openpyxl).

Icerik once 'bolumler' listesine donusturulur, PDF ve Excel ayni listeden cizilir (iki format hep ayni icerik).
Gorunum ve hangi alanlarin basilacagi 'tasarim'dan gelir (teklif_tasarim: Kurumsal / Modern / Sade).
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


def para(x):
    """1234567.8 -> '1.234.567,80 TL'"""
    s = '{:,.2f}'.format(float(x or 0)).replace(',', 'X').replace('.', ',').replace('X', '.')
    return s + ' TL'


def _g(x):
    return ('%g' % x).replace('.', ',') if isinstance(x, (int, float)) else str(x)


def _tarih(girdi):
    t = girdi.get('tarih')
    try:
        return datetime.strptime(t, '%Y-%m-%d') if t else datetime.now()
    except ValueError:
        return datetime.now()


def satir_aciklamasi(s):
    if s['tur'].lower().startswith(('mas', 'mob')):
        return '%d kullanıcı' % s['adet']
    return '%d adet' % s['adet'] if s['adet'] != 1 else ''


def _kisa(metin, n=150):
    metin = re.sub(r'\s+', ' ', metin or '').strip()
    return metin if len(metin) <= n else metin[:n].rsplit(' ', 1)[0] + ' …'


# Musteriye gosterilmeyecek ic gerekceler (elle degisiklik, kalibrasyon)
_IC_GEREKCE = ('Elle girildi', 'Kalibrasyon')


def _sure(saat, ayar, gun=True):
    gun_saat = float(ayar.get('is_gunu_saat') or 0)
    metin = '%s saat' % _g(saat)
    if gun and gun_saat > 0 and saat >= gun_saat / 2:
        metin += '\n(≈ %s iş günü)' % _g(round(saat / gun_saat, 1))
    return metin


def uyarlama_satirlari(s, ayar, goster=None):
    """Uyarlama tablosu: her kalemin kapsami, neye gore hesaplandigi, suresi ve tutari."""
    G = goster or teklif_tasarim.VARSAYILAN['goster']
    u = s['uyarlama']
    rows = []
    for k in s['hizmet']['kalemler']:
        if not k['saat']:
            continue
        dayanak = [n for n in k.get('nedenler', []) if not n.startswith(_IC_GEREKCE)]
        metin = k['kapsam']
        if dayanak and G.get('uyarlama_dayanak', True):
            metin += '\nDayanak: ' + '; '.join(dayanak)
        if u['saat_ucreti'] and G.get('uyarlama_hesap', True):
            metin += '\nHesap: %s saat × %s = %s' % (_g(k['saat']), para(u['saat_ucreti']), para(k['tutar']))
        rows.append([k['ad'], metin.strip(), _sure(k['saat'], ayar, G.get('uyarlama_gun', True)), para(k['tutar'])])
    if u['lisans_payi']:
        metin = 'Satın alınan modüllerin firmanıza göre yapılandırılması, parametre ve sürüm uyumu.'
        if G.get('uyarlama_hesap', True):
            metin += '\nHesap: lisans net tutarı %s × %%%s = %s' % (
                para(s['lisans']['net']), _g(float(ayar['lisans_yuzdesi'])), para(u['lisans_payi']))
        rows.append(['Lisans uyarlama payı', metin, '', para(u['lisans_payi'])])
    if u['sinir']:
        alt = u['sinir'] == 'ALT SINIR'
        rows.append(['Asgari uyarlama farkı' if alt else 'Azami uyarlama indirimi',
                     'Uyarlama bedeli asgari tutara tamamlanmıştır.' if alt else
                     'Uyarlama bedeli azami tutarla sınırlandırılmıştır.', '', para(u['net'] - u['ham'])])
    rows.append(['UYARLAMA TOPLAMI', '', '%s saat' % _g(u['saat']), para(u['net'])])
    return rows


def uyarlama_ozeti(s, ayar, goster=None):
    G = goster or teklif_tasarim.VARSAYILAN['goster']
    u = s['uyarlama']
    parca = ['KDV dahil %s' % para(u['dahil'])]
    if u['saat']:
        parca.append('Toplam %s' % _sure(u['saat'], ayar, G.get('uyarlama_gun', True)).replace('\n', ' '))
    if u['saat_ucreti']:
        parca.append('Saat ücreti %s + KDV' % para(u['saat_ucreti']))
    if u['lisans_payi']:
        parca.append('Lisans uyarlama payı dahil')
    return '  |  '.join(parca)


def firma_satirlari(ayar):
    """Teklifi veren firmanin iletisim satirlari (ad haric)."""
    satirlar = []
    if ayar.get('firma_adres'):
        satirlar.append(ayar['firma_adres'])
    iletisim = '  |  '.join(x for x in (
        ('Tel: ' + ayar['firma_telefon']) if ayar.get('firma_telefon') else '',
        ayar.get('firma_eposta') or '', ayar.get('firma_web') or '') if x)
    if iletisim:
        satirlar.append(iletisim)
    if ayar.get('firma_iletisim'):
        satirlar.append(ayar['firma_iletisim'])
    if ayar.get('firma_vergi'):
        satirlar.append('VD / VKN: ' + ayar['firma_vergi'])
    return satirlar


def _hazirlayan(ayar):
    return [x for x in (ayar.get('hazirlayan_ad'), ayar.get('hazirlayan_unvan')) if x]


# ======================================================================
# Icerik (bolumler)
# ======================================================================
# ('baslik', metin, alt) | ('bilgi', [(etiket, deger)]) | ('taraflar', veren[(e, d)], edilen[(e, d)])
# ('paragraf', metin) | ('bolum', metin) | ('liste', [(etiket, deger)]) | ('maddeler', [metin])
# ('tablo', basliklar, satirlar, {'sag': {kolon...}, 'toplam': bool, 'oran': [genislik...], 'vurgu': {satir...}})
# ('ozet', [(etiket, tutar, kalin)], son_satir_vurgulu) | ('metin', metin, kucuk)
# ('vurgu', etiket, buyuk_tutar_metni, alt_metin) | ('imza', veren_satirlari)
def bolumler(girdi, s, ayar, tasarim=None):
    T = teklif_tasarim.tamamla(tasarim)
    G = T['goster']
    B = []
    no = girdi.get('teklif_no') or '-'
    alt = []
    m = re.search(r'-R(\d+)$', no)
    if m:
        alt.append('Revizyon %s' % m.group(1))
    if girdi.get('etiket'):
        alt.append(girdi['etiket'] + ' Teklif')
    B.append(('baslik', T['baslik'] or 'TEKLİF', ' · '.join(alt)))

    tarih = _tarih(girdi)
    gecerlilik = (tarih + timedelta(days=int(ayar.get('gecerlilik_gun') or 0))).strftime('%d.%m.%Y')
    mb = girdi.get('musteri_bilgi') or {}
    degerler = {'musteri': girdi.get('musteri') or 'Firmanız', 'yetkili': mb.get('yetkili') or 'Yetkili',
                'tarih': tarih.strftime('%d.%m.%Y'), 'gecerlilik': gecerlilik, 'teklif_no': no,
                'firma': ayar.get('firma_adi', ''), 'hazirlayan': ayar.get('hazirlayan_ad', '')}
    kullanici = '%d masaüstü / %d mobil' % (int(girdi.get('masaustu') or 0), int(girdi.get('mobil') or 0))
    bilgi = [('Teklif No', no), ('Tarih', tarih.strftime('%d.%m.%Y')), ('Geçerlilik', gecerlilik),
             ('Paket', s['paket_tipi']), ('Kullanıcı', kullanici)]
    if G['taraflar']:
        veren = [('', ayar.get('firma_adi', ''))] + [('', x) for x in firma_satirlari(ayar)]
        if G['hazirlayan'] and ayar.get('hazirlayan_ad'):
            veren.append(('Hazırlayan', ', '.join(_hazirlayan(ayar))))
            iletisim = '  |  '.join(x for x in (ayar.get('hazirlayan_telefon'), ayar.get('hazirlayan_eposta')) if x)
            if iletisim:
                veren.append(('İletişim', iletisim))
        edilen = [('', girdi.get('musteri', ''))]
        for etiket, alan in (('Yetkili', 'yetkili'), ('Telefon', 'telefon'), ('E-posta', 'eposta'),
                             ('Adres', 'adres')):
            if mb.get(alan):
                edilen.append((etiket, mb[alan]))
        if mb.get('vkn') or mb.get('vergi_dairesi'):
            edilen.append(('VD / VKN', ' / '.join(x for x in (mb.get('vergi_dairesi'), mb.get('vkn')) if x)))
        B.append(('bilgi', bilgi))
        B.append(('taraflar', veren, edilen))
    else:
        bilgi.insert(2, ('Müşteri', girdi.get('musteri', '')))
        if mb.get('yetkili'):
            bilgi.append(('Yetkili', mb['yetkili']))
        B.append(('bilgi', bilgi))
    if G['on_yazi'] and (T['on_yazi'] or '').strip():
        B.append(('paragraf', teklif_tasarim.yer_tutucu(T['on_yazi'].strip(), degerler)))

    # 1. Lisans
    basliklar, oran = ['#', 'Paket / Modül'], [4, 0]
    if G['kullanici_kolonu']:
        basliklar.append('Kapsam')
        oran.append(12)
    if G['liste_indirim']:
        basliklar += ['Liste Tutarı', 'İndirim']
        oran += [16, 16]
    basliklar.append('Net Tutar')
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
            satir.append(satir_aciklamasi(x))
        if G['liste_indirim']:
            ind = para(x['indirim']) if x['indirim'] else '-'
            if x.get('kamp_ad'):
                ind += '\n' + x['kamp_ad']
            satir += [para(x['toplam']), ind]
        satirlar.append(satir + [para(x['net'])])
    toplam = ['', 'LİSANS TOPLAMI'] + ([''] if G['kullanici_kolonu'] else [])
    if G['liste_indirim']:
        toplam += [para(s['lisans']['toplam']), para(s['lisans']['indirim']) if s['lisans']['indirim'] else '-']
    satirlar.append(toplam + [para(s['lisans']['net'])])
    bol = 1
    B.append(('bolum', '%d. Lisans' % bol))
    B.append(('tablo', basliklar, satirlar, {'sag': sag, 'toplam': True, 'oran': oran}))
    d = s.get('dia') or {}
    if G['dia_ek_bilgi'] and d:
        ek = [(e, para(v) + ' + KDV') for e, v in (('Ek +1 masaüstü kullanıcı bedeli', d.get('ek_masaustu')),
                                                   ('Ek +1 mobil kullanıcı bedeli', d.get('ek_mobil')),
                                                   ('Yıllık bulut hizmet bedeli', d.get('yillik_bulut'))) if v]
        if ek:
            B.append(('liste', ek))

    # 2. Uyarlama hizmeti
    u = s['uyarlama']
    kalemler = [k for k in s['hizmet']['kalemler'] if k['saat']]
    if kalemler or u['net']:
        bol += 1
        B.append(('bolum', '%d. Uyarlama Hizmeti (Kurulum, Aktarım, Eğitim ve Danışmanlık)' % bol))
        if G['uyarlama_aciklama'] and (ayar.get('uyarlama_aciklama') or '').strip():
            B.append(('metin', ayar['uyarlama_aciklama'].strip(), False))
        B.append(('tablo', ['Hizmet', 'Kapsam ve hesap dayanağı', 'Süre', 'Tutar'], uyarlama_satirlari(s, ayar, G),
                  {'sag': {2, 3}, 'toplam': True, 'oran': [21, 46, 16, 17]}))
        B.append(('vurgu', 'UYARLAMA HİZMET BEDELİ', '%s + KDV' % para(u['net']), uyarlama_ozeti(s, ayar, G)))
        haric = (ayar.get('kapsam_haric') or '').strip() if G['kapsam_haric'] else ''
        if u['saat_ucreti'] and G['kapsam_haric']:
            haric += (' ' if haric else '') + 'Ek hizmet saat ücreti: %s + KDV.' % para(u['saat_ucreti'])
        if haric:
            B.append(('metin', haric, True))

    # Ozet
    bol += 1
    B.append(('bolum', '%d. Teklif Özeti' % bol))
    ozet = [('Lisans', s['lisans']['net'], False),
            ('Uyarlama Hizmeti (%s saat)' % _g(u['saat']) if u['saat'] else 'Uyarlama Hizmeti', u['net'], False)]
    ozet += [(k['ad'], k['net'], False) for k in s['sabit']['kalemler']]
    if s['seyahat']['tutar']:
        ozet.append(('Yol / konaklama giderleri (%d gün yerinde)' % s['seyahat']['gun'], s['seyahat']['tutar'], False))
    ozet.append(('ARA TOPLAM', s['ara_toplam'], True))
    if s['kampanya']['tutar']:
        ozet.append(('Kampanya İndirimi (%s, %%%s)' % (s['kampanya']['ad'], _g(s['kampanya']['yuzde'])),
                     -s['kampanya']['tutar'], False))
    ozet.append(('NET TOPLAM (KDV Hariç)', s['net_toplam'], True))
    if G['ozet_kdv']:
        ozet += [('KDV (%%%s)' % _g(s['kdv_orani']), s['kdv'], False),
                 ('GENEL TOPLAM (KDV Dahil)', s['genel_toplam'], True)]
    B.append(('ozet', ozet, True))

    # Odeme secenekleri
    od = [o for o in s['odeme'] if o['musteriye_goster']] if G['odeme'] else []
    if od:
        bol += 1
        B.append(('bolum', '%d. Ödeme Seçenekleri' % bol))
        rows = []
        for o in od:
            if o['indirim'] and o['tur'] == 'PESIN':
                notu = '%%%s peşin indirim' % _g(o['indirim'])
            elif o['vade_farki']:
                notu = '%%%s vade farkı' % _g(o['vade_farki'])
            else:
                notu = 'Vade farksız' if o['taksit'] > 1 else ''
            rows.append([o['ad'], str(o['taksit']) if o['taksit'] > 1 else 'Tek ödeme',
                         para(o['taksit_tutari']), para(o['dahil']), notu])
        B.append(('tablo', ['Ödeme Seçeneği', 'Taksit', 'Taksit Tutarı (KDV Dahil)', 'Toplam (KDV Dahil)', 'Not'],
                  rows, {'sag': {2, 3}, 'oran': [28, 12, 22, 22, 16]}))

    # Bakim anlasmasi
    b = s.get('bakim') or {}
    if G['bakim'] and b.get('sun') and b.get('paketler') and s['satirlar']:
        bol += 1
        B.append(('bolum', '%d. Bakım ve Destek Anlaşması Seçenekleri' % bol))
        rows, vurgu = [], set()
        for n, p in enumerate(b['paketler']):
            ad = p['ad'] + (' ★ Önerilen' if p['onerilen'] else '') + (' ✔ Seçilen' if p['ad'] == b.get('secili') else '')
            if p['onerilen'] or p['ad'] == b.get('secili'):
                vurgu.add(n)
            rows.append([ad, '%s saat' % _g(p['saat']), p['yanit'], _g(p['yerinde_gun']) + ' gün', p['aciklama'],
                         para(p['yillik']), para(p['aylik'])])
        B.append(('tablo', ['Paket', 'Yıllık Destek', 'Yanıt Süresi', 'Yerinde', 'Kapsam', 'Yıllık (KDV Hr.)',
                            'Aylık (KDV Hr.)'], rows, {'sag': {5, 6}, 'oran': [16, 11, 11, 8, 28, 13, 13],
                                                         'vurgu': vurgu}))

    # Toplam sahip olma maliyeti
    t = s.get('tco') or {}
    if G['tco'] and t.get('yillar') and (s['ybh']['yillik']['net'] or any(y['bakim'] for y in t['yillar'])):
        bol += 1
        B.append(('bolum', '%d. %d Yıllık Toplam Maliyet' % (bol, len(t['yillar']))))
        rows = [['%d. yıl' % y['yil'], para(y['tek_seferlik']) if y['tek_seferlik'] else '-',
                 para(y['ybh']) if y['ybh'] else '-', para(y['bakim']) if y['bakim'] else '-', para(y['toplam']),
                 para(y['dahil'])] for y in t['yillar']]
        rows.append(['TOPLAM', '', '', '', para(t['toplam']), para(t['dahil'])])
        B.append(('tablo', ['Yıl', 'Lisans + Uyarlama', 'DIA Yıllık Bakım', 'Destek Anlaşması', 'Toplam (KDV Hr.)',
                            'KDV Dahil'], rows, {'sag': {1, 2, 3, 4, 5}, 'toplam': True,
                                                  'oran': [10, 18, 18, 18, 18, 18]}))
        B.append(('metin', 'DIA yıllık bakım bedeli DIA\'nın güncel fiyatlarına göre tahminidir.', True))

    if G['sartlar'] and (T['sartlar'] or '').strip():
        B.append(('bolum', 'Ticari Şartlar'))
        B.append(('maddeler', [x.strip() for x in teklif_tasarim.yer_tutucu(T['sartlar'], degerler).split('\n')
                               if x.strip()]))
    if G['banka'] and (ayar.get('firma_banka') or '').strip():
        B.append(('bolum', 'Banka Bilgileri'))
        B.append(('metin', ayar['firma_banka'].strip(), False))
    notlar = '\n'.join(x for x in [ayar.get('teklif_notlari', ''), girdi.get('notlar', '')] if x)
    if G['notlar'] and notlar:
        B.append(('bolum', 'Notlar'))
        B.append(('metin', notlar, True))
    if (T['kapanis'] or '').strip():
        B.append(('paragraf', teklif_tasarim.yer_tutucu(T['kapanis'].strip(), degerler)))
    if G['imza']:
        B.append(('imza', [ayar.get('firma_adi', '')] + (_hazirlayan(ayar) if G['hazirlayan'] else [])))
    return B


def karsilastirma_bolumleri(teklifler, ayar, tasarim=None):
    """teklifler: [(girdi, sonuc)] -> alternatif teklif karsilastirmasi."""
    G = teklif_tasarim.tamamla(tasarim)['goster']
    B = [('baslik', 'DİA ERP TEKLİF SEÇENEKLERİ', 'Karşılaştırma')]
    ilk = teklifler[0][0]
    B.append(('bilgi', [('Müşteri', ilk.get('musteri', '')), ('Tarih', datetime.now().strftime('%d.%m.%Y'))]))
    basliklar = [''] + [(g.get('etiket') or 'Seçenek %d' % n) + '\n' + (g.get('teklif_no') or '')
                        for n, (g, _) in enumerate(teklifler, 1)]
    satirlar = [
        ['Paket tipi'] + [s['paket_tipi'] for _, s in teklifler],
        ['Kullanıcı'] + ['%s masaüstü / %s mobil' % (g.get('masaustu', 0), g.get('mobil', 0)) for g, _ in teklifler],
        ['Paketler'] + ['\n'.join(x['ad'] for x in s['satirlar']) for _, s in teklifler],
        ['Lisans'] + [para(s['lisans']['net']) for _, s in teklifler],
        ['Uyarlama'] + ['%s\n(%s saat)' % (para(s['uyarlama']['net']), _g(s['uyarlama']['saat'])) for _, s in teklifler],
        ['Diğer (sabit, yol)'] + [para(s['sabit']['net'] + s['seyahat']['tutar']) for _, s in teklifler],
        ['Kampanya indirimi'] + [para(-s['kampanya']['tutar']) if s['kampanya']['tutar'] else '-' for _, s in teklifler],
        ['NET TOPLAM (KDV Hariç)'] + [para(s['net_toplam']) for _, s in teklifler],
        ['GENEL TOPLAM (KDV Dahil)'] + [para(s['genel_toplam']) for _, s in teklifler],
    ]
    if all(s.get('tco') for _, s in teklifler):
        satirlar.append(['%d yıllık toplam maliyet' % len(teklifler[0][1]['tco']['yillar'])] +
                        [para(s['tco']['dahil']) for _, s in teklifler])
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


def _pdf_ciz(yol, B, ayar, belge_no, tasarim=None):
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_RIGHT
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.pdfgen import canvas as rl_canvas
    from reportlab.platypus import Image, KeepTogether, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    T = teklif_tasarim.tamamla(tasarim)
    G, sablon, fs = T['goster'], T['sablon'], float(T['yazi_boyutu'])
    sade, modern = sablon == 'Sade', sablon == 'Modern'
    font, kalin, cevir = _fontlar()
    renk = colors.HexColor
    ana, vurgu = renk(T['ana_renk']), renk(T['vurgu_renk'])
    ana_acik = renk(teklif_tasarim.acik_ton(T['ana_renk'], 0.88))
    ana_cok_acik = renk(teklif_tasarim.acik_ton(T['ana_renk'], 0.95))
    vurgu_acik = renk(teklif_tasarim.acik_ton(T['vurgu_renk'], 0.8))
    gri, gri_yazi, beyaz = renk('#cccccc'), renk('#555555'), colors.white
    tablo_bas_zemin, tablo_bas_yazi = (renk('#eeeeee'), colors.black) if sade else (ana, beyaz)
    genislik = A4[0] - 32 * mm

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

    # ---------------------------------------------------------------- ust bilgi (logo + firma)
    akis = []
    firma_adi = ayar.get('firma_adi', '')
    satirlar = firma_satirlari(ayar)
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
                akis.append(P(b[1], fs + 7, True, rnk=colors.black if sade else ana))
                if b[2]:
                    akis.append(P(b[2], fs + 1, rnk=gri_yazi if sade else vurgu))
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
            tb = Table([[kutu('TEKLİFİ VEREN', b[1], yarim), '', kutu('TEKLİF EDİLEN', b[2], yarim)]],
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
            if akis and isinstance(akis[-1], Spacer):
                akis.pop()
            akis += [Spacer(1, 1 * mm), KeepTogether([kutu_]), Spacer(1, 3 * mm)]
        elif tur == 'ozet':
            veri = [[P(a, fs, k), P(para(v), fs, k, TA_RIGHT)] for a, v, k in b[1]]
            if b[2]:
                a, v, _ = b[1][-1]
                veri[-1] = [P(a, fs + 1, True, rnk=beyaz), P(para(v), fs + 1, True, TA_RIGHT, beyaz)]
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
            sol = [P('TEKLİFİ VEREN', fs + 1, True, rnk=baslik_rengi)] + \
                  [P(x, fs, i == 0) for i, x in enumerate(veren) if x] + \
                  [Spacer(1, 10 * mm), P('İmza: _______________________', fs)]
            sag_ = [P('MÜŞTERİ ONAYI', fs + 1, True, rnk=baslik_rengi),
                    P('Firma Kaşe / İmza: _______________________', fs), Spacer(1, 3 * mm),
                    P('Yetkili Ad Soyad: _______________________', fs), Spacer(1, 3 * mm),
                    P('Tarih: _______________________', fs)]
            tb = Table([[sol, '', sag_]], colWidths=[yarim, 10 * mm, yarim])
            tb.setStyle(TableStyle([('VALIGN', (0, 0), (-1, -1), 'TOP'),
                                    ('LINEABOVE', (0, 0), (0, 0), 0.8, gri), ('LINEABOVE', (2, 0), (2, 0), 0.8, gri),
                                    ('TOPPADDING', (0, 0), (-1, -1), 6)]))
            akis += [Spacer(1, 4 * mm), KeepTogether([tb])]

    alt_metin = '  ·  '.join(x for x in (firma_adi, ayar.get('firma_telefon'), ayar.get('firma_eposta'),
                                         ayar.get('firma_web')) if x)

    def alt_bilgi(tuval, sayfa, toplam):
        if not G['alt_bilgi']:
            return
        tuval.saveState()
        tuval.setStrokeColor(gri if sade else ana)
        tuval.setLineWidth(0.6)
        tuval.line(16 * mm, 12.5 * mm, A4[0] - 16 * mm, 12.5 * mm)
        tuval.setFont(font, 7)
        tuval.setFillColor(renk('#777777'))
        tuval.drawString(16 * mm, 8.5 * mm, t(alt_metin))
        tuval.drawRightString(A4[0] - 16 * mm, 8.5 * mm, t('%s  ·  Sayfa %d / %d' % (belge_no, sayfa, toplam)))
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
                alt_bilgi(self, n, toplam)
                super().showPage()
            super().save()

    belge = SimpleDocTemplate(yol, pagesize=A4, leftMargin=16 * mm, rightMargin=16 * mm, topMargin=14 * mm,
                              bottomMargin=18 * mm, title=t(belge_no), author=t(firma_adi))
    belge.build(akis, canvasmaker=SayfaliTuval)
    return yol


# ======================================================================
# Excel
# ======================================================================
def _excel_ciz(yol, B, ayar, sayfa_adi='Teklif', tasarim=None):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    T = teklif_tasarim.tamamla(tasarim)
    G, sade = T['goster'], T['sablon'] == 'Sade'
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
    if logo_var:
        birlesik(1, ayar.get('firma_adi', ''), ad_font, 4, 7, Alignment(horizontal='right'))
        for i, x in enumerate(firma_satirlari(ayar)[:3], 2):
            birlesik(i, x, Font(size=9, color='555555'), 4, 7, Alignment(horizontal='right'))
        r = 5
    else:
        ws.cell(1, 1, ayar.get('firma_adi', '')).font = ad_font
        r = 2
        for x in firma_satirlari(ayar):
            ws.cell(r, 1, x).font = Font(size=9, color='555555')
            r += 1
    for c in range(1, 8):
        ws.cell(r, c).border = Border(bottom=Side(style='medium', color='999999' if sade else ana))
    r += 2

    for b in B:
        tur = b[0]
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
            for c1, baslik, satirlar_ in ((1, 'TEKLİFİ VEREN', b[1]), (5, 'TEKLİF EDİLEN', b[2])):
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
            for n, (a, v, k) in enumerate(b[1]):
                ws.cell(r, 1, a).font = Font(bold=k)
                hc = ws.cell(r, 4, round(v, 2))
                hc.number_format = '#,##0.00 "TL"'
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
            ws.cell(r, 1, 'TEKLİFİ VEREN').font = Font(bold=True, color=renk_)
            ws.cell(r, 5, 'MÜŞTERİ ONAYI').font = Font(bold=True, color=renk_)
            sol = [x for x in veren if x] + ['', 'İmza: ______________________']
            sag_ = ['Firma Kaşe / İmza: ______________________', '', 'Yetkili Ad Soyad: ______________________', '',
                    'Tarih: ______________________']
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
        ws.oddFooter.right.text = 'Sayfa &P / &N'
        ws.oddFooter.right.size = 8
    wb.save(yol)
    return yol


# ======================================================================
# Dis arayuz
# ======================================================================
def pdf_olustur(yol, girdi, sonuc, ayar, tasarim=None):
    return _pdf_ciz(yol, bolumler(girdi, sonuc, ayar, tasarim), ayar, 'Teklif ' + str(girdi.get('teklif_no') or ''),
                    tasarim)


def excel_olustur(yol, girdi, sonuc, ayar, tasarim=None):
    return _excel_ciz(yol, bolumler(girdi, sonuc, ayar, tasarim), ayar, tasarim=tasarim)


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
