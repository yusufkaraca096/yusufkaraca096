"""Musteriye gidecek teklif ciktisi: PDF (reportlab) ve Excel (openpyxl).

Icerik once 'bolumler' listesine donusturulur, PDF ve Excel ayni listeden cizilir (iki format hep ayni icerik).
Ic bilgiler (maliyet, kar, rakip, komisyon, satis asistani) ciktiya YAZILMAZ.
"""
import os
import re
from datetime import datetime, timedelta

_FONT_ADAYLARI = [
    ('C:/Windows/Fonts/arial.ttf', 'C:/Windows/Fonts/arialbd.ttf'),
    ('C:/Windows/Fonts/segoeui.ttf', 'C:/Windows/Fonts/segoeuib.ttf'),
    ('C:/Windows/Fonts/calibri.ttf', 'C:/Windows/Fonts/calibrib.ttf'),
    ('/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf', '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'),
    ('/Library/Fonts/Arial Unicode.ttf', '/Library/Fonts/Arial Unicode.ttf'),
]
_ASCII = str.maketrans('çğıöşüÇĞİÖŞÜ★✔', 'cgiosuCGIOSU*+')


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


# Musteriye gosterilmeyecek ic gerekceler (elle degisiklik, kalibrasyon)
_IC_GEREKCE = ('Elle girildi', 'Kalibrasyon')


def _sure(saat, ayar):
    gun_saat = float(ayar.get('is_gunu_saat') or 0)
    metin = '%s saat' % _g(saat)
    if gun_saat > 0 and saat >= gun_saat / 2:
        metin += '\n(≈ %s iş günü)' % _g(round(saat / gun_saat, 1))
    return metin


def uyarlama_satirlari(s, ayar):
    """Uyarlama tablosu: her kalemin kapsamı, neye göre hesaplandığı, süresi ve tutarı."""
    u = s['uyarlama']
    rows = []
    for k in s['hizmet']['kalemler']:
        if not k['saat']:
            continue
        dayanak = [n for n in k.get('nedenler', []) if not n.startswith(_IC_GEREKCE)]
        metin = k['kapsam']
        if dayanak:
            metin += '\nDayanak: ' + '; '.join(dayanak)
        if u['saat_ucreti']:
            metin += '\nHesap: %s saat × %s = %s' % (_g(k['saat']), para(u['saat_ucreti']), para(k['tutar']))
        rows.append([k['ad'], metin, _sure(k['saat'], ayar), para(k['tutar'])])
    if u['lisans_payi']:
        rows.append(['Lisans uyarlama payı',
                     'Satın alınan modüllerin firmanıza göre yapılandırılması, parametre ve sürüm uyumu.\n'
                     'Hesap: lisans net tutarı %s × %%%s = %s' % (
                         para(s['lisans']['net']), _g(float(ayar['lisans_yuzdesi'])), para(u['lisans_payi'])),
                     '', para(u['lisans_payi'])])
    if u['sinir']:
        alt = u['sinir'] == 'ALT SINIR'
        rows.append(['Asgari uyarlama farkı' if alt else 'Azami uyarlama indirimi',
                     'Uyarlama bedeli asgari tutara tamamlanmıştır.' if alt else
                     'Uyarlama bedeli azami tutarla sınırlandırılmıştır.', '', para(u['net'] - u['ham'])])
    rows.append(['UYARLAMA TOPLAMI', '', '%s saat' % _g(u['saat']), para(u['net'])])
    return rows


def uyarlama_ozeti(s, ayar):
    u = s['uyarlama']
    parca = ['KDV dahil %s' % para(u['dahil'])]
    if u['saat']:
        parca.append('Toplam %s' % _sure(u['saat'], ayar).replace('\n', ' '))
    if u['saat_ucreti']:
        parca.append('Saat ücreti %s + KDV' % para(u['saat_ucreti']))
    if u['lisans_payi']:
        parca.append('Lisans uyarlama payı dahil')
    return '  |  '.join(parca)


# ======================================================================
# Icerik (bolumler)
# ======================================================================
# ('baslik', metin, alt_metin) | ('bilgi', [(etiket, deger), ...]) | ('bolum', metin)
# ('tablo', basliklar, satirlar, {'sag': {kolon...}, 'toplam': bool, 'oran': [genislik...], 'vurgu': {satir...}})
# ('ozet', [(etiket, tutar, kalin)], son_satir_vurgulu) | ('metin', metin, kucuk) | ('onay',)
# ('vurgu', etiket, buyuk_tutar_metni, alt_metin): one cikan kutu
def bolumler(girdi, s, ayar):
    B = []
    no = girdi.get('teklif_no') or '-'
    alt = []
    m = re.search(r'-R(\d+)$', no)
    if m:
        alt.append('Revizyon %s' % m.group(1))
    if girdi.get('etiket'):
        alt.append(girdi['etiket'] + ' Teklif')
    B.append(('baslik', 'DİA ERP LİSANS + HİZMET TEKLİFİ', ' · '.join(alt)))

    tarih = _tarih(girdi)
    mb = girdi.get('musteri_bilgi') or {}
    bilgi = [('Teklif No', no), ('Tarih', tarih.strftime('%d.%m.%Y')), ('Müşteri', girdi.get('musteri', '')),
             ('Geçerlilik', (tarih + timedelta(days=int(ayar.get('gecerlilik_gun') or 0))).strftime('%d.%m.%Y'))]
    if mb.get('yetkili'):
        bilgi.append(('Yetkili', mb['yetkili']))
    iletisim = ' / '.join(x for x in (mb.get('telefon'), mb.get('eposta')) if x)
    if iletisim:
        bilgi.append(('İletişim', iletisim))
    if mb.get('vkn'):
        bilgi.append(('VKN / VD', ' / '.join(x for x in (mb.get('vkn'), mb.get('vergi_dairesi')) if x)))
    bilgi += [('Kullanıcı', '%d masaüstü / %d mobil' % (int(girdi.get('masaustu') or 0), int(girdi.get('mobil') or 0))),
              ('Paket', s['paket_tipi'])]
    B.append(('bilgi', bilgi))

    # 1. Lisans
    satirlar = [[str(n), '%s\n%s' % (x['ad'], x['kod']), satir_aciklamasi(x), para(x['toplam']),
                 para(x['indirim']) if x['indirim'] else '-', para(x['net'])] for n, x in enumerate(s['satirlar'], 1)]
    satirlar.append(['', 'LİSANS TOPLAMI', '', para(s['lisans']['toplam']),
                     para(s['lisans']['indirim']) if s['lisans']['indirim'] else '-', para(s['lisans']['net'])])
    bol = 1
    B.append(('bolum', '%d. Lisans' % bol))
    B.append(('tablo', ['#', 'Paket / Modül', 'Kapsam', 'Liste Tutarı', 'İndirim', 'Net Tutar'], satirlar,
              {'sag': {3, 4, 5}, 'toplam': True, 'oran': [4, 34, 12, 17, 16, 17]}))

    # 2. Uyarlama hizmeti
    u = s['uyarlama']
    kalemler = [k for k in s['hizmet']['kalemler'] if k['saat']]
    if kalemler or u['net']:
        bol += 1
        B.append(('bolum', '%d. Uyarlama Hizmeti (Kurulum, Aktarım, Eğitim ve Danışmanlık)' % bol))
        if (ayar.get('uyarlama_aciklama') or '').strip():
            B.append(('metin', ayar['uyarlama_aciklama'].strip(), False))
        B.append(('tablo', ['Hizmet', 'Kapsam ve hesap dayanağı', 'Süre', 'Tutar'], uyarlama_satirlari(s, ayar),
                  {'sag': {2, 3}, 'toplam': True, 'oran': [22, 48, 13, 17]}))
        B.append(('vurgu', 'UYARLAMA HİZMET BEDELİ', '%s + KDV' % para(u['net']), uyarlama_ozeti(s, ayar)))
        haric = (ayar.get('kapsam_haric') or '').strip()
        if u['saat_ucreti']:
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
    ozet += [('NET TOPLAM (KDV Hariç)', s['net_toplam'], True), ('KDV (%%%s)' % _g(s['kdv_orani']), s['kdv'], False),
             ('GENEL TOPLAM (KDV Dahil)', s['genel_toplam'], True)]
    B.append(('ozet', ozet, True))

    # Odeme secenekleri
    od = [o for o in s['odeme'] if o['musteriye_goster']]
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
    if b.get('sun') and b.get('paketler') and s['satirlar']:
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
    if t.get('yillar') and (s['ybh']['yillik']['net'] or any(y['bakim'] for y in t['yillar'])):
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

    notlar = '\n'.join(x for x in [ayar.get('teklif_notlari', ''), girdi.get('notlar', '')] if x)
    if notlar:
        B.append(('bolum', 'Notlar ve Şartlar'))
        B.append(('metin', notlar, True))
    B.append(('onay',))
    return B


def karsilastirma_bolumleri(teklifler, ayar):
    """teklifler: [(girdi, sonuc)] -> alternatif teklif karsilastirmasi."""
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
    B.append(('onay',))
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


def _pdf_ciz(yol, B, ayar, belge_no):
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_RIGHT
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import Image, KeepTogether, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    font, kalin, cevir = _fontlar()
    lacivert, acik, gri = colors.HexColor('#2c3e50'), colors.HexColor('#ecf0f1'), colors.HexColor('#cccccc')
    sari = colors.HexColor('#fef5d4')
    genislik = A4[0] - 32 * mm

    def t(metin):
        metin = str(metin)
        return metin.translate(cevir) if cevir else metin

    def P(metin, boyut=9, kalin_mi=False, hiza=None, renk=colors.black):
        stil = ParagraphStyle('p', fontName=kalin if kalin_mi else font, fontSize=boyut, leading=boyut * 1.3,
                              textColor=renk, alignment=hiza if hiza is not None else 0)
        return Paragraph(t(metin).replace('&', '&amp;').replace('<', '&lt;').replace('\n', '<br/>'), stil)

    akis = []
    firma = [P(ayar.get('firma_adi', ''), 15, True, renk=lacivert)]
    for satir in (ayar.get('firma_adres'), ayar.get('firma_iletisim')):
        if satir:
            firma.append(P(satir, 8, renk=colors.HexColor('#555555')))
    logo = ayar.get('logo_yolu')
    baslik = Table([[firma]])
    if logo and os.path.exists(logo):
        try:
            resim = Image(logo)
            oran = min(45 * mm / resim.imageWidth, 20 * mm / resim.imageHeight)
            resim.drawWidth, resim.drawHeight = resim.imageWidth * oran, resim.imageHeight * oran
            baslik = Table([[resim, firma]], colWidths=[50 * mm, None])
        except Exception:
            pass
    baslik.setStyle(TableStyle([('VALIGN', (0, 0), (-1, -1), 'MIDDLE'), ('LEFTPADDING', (0, 0), (-1, -1), 0)]))
    akis += [baslik, Spacer(1, 4 * mm)]

    bekleyen = None      # bolum basligi bir sonraki tabloyla birlikte tutulur (sayfa sonunda yalniz kalmasin)
    for b in B:
        tur = b[0]
        if tur == 'baslik':
            hucre = [P(b[1], 13, True, renk=colors.white)]
            if b[2]:
                hucre.append(P(b[2], 9, renk=colors.HexColor('#f1c40f')))
            akis.append(Table([[hucre]], colWidths=[genislik], style=[
                ('BACKGROUND', (0, 0), (-1, -1), lacivert), ('TOPPADDING', (0, 0), (-1, -1), 5),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 5)]))
            akis.append(Spacer(1, 3 * mm))
        elif tur == 'bilgi':
            hucreler = []
            for i in range(0, len(b[1]), 2):
                cift = b[1][i:i + 2] + [('', '')] * (2 - len(b[1][i:i + 2]))
                hucreler.append([P(cift[0][0], 9, True), P(cift[0][1]), P(cift[1][0], 9, True), P(cift[1][1])])
            tb = Table(hucreler, colWidths=[24 * mm, 66 * mm, 24 * mm, None])
            tb.setStyle(TableStyle([('GRID', (0, 0), (-1, -1), 0.4, colors.HexColor('#bbbbbb')),
                                    ('BACKGROUND', (0, 0), (0, -1), acik), ('BACKGROUND', (2, 0), (2, -1), acik),
                                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')]))
            akis += [tb, Spacer(1, 4 * mm)]
        elif tur == 'bolum':
            bekleyen = [P(b[1], 11, True, renk=lacivert), Spacer(1, 2 * mm)]
        elif tur == 'tablo':
            _, basliklar, satirlar, sec = b
            sag = sec.get('sag', set())
            oranlar = sec.get('oran') or [1] * len(basliklar)
            kolon = [genislik * o / sum(oranlar) for o in oranlar]
            veri = [[P(h, 8, True, TA_RIGHT if i in sag else None, colors.white) for i, h in enumerate(basliklar)]]
            for r, satir in enumerate(satirlar):
                son = sec.get('toplam') and r == len(satirlar) - 1
                veri.append([P(v, 8, son, TA_RIGHT if i in sag else None) for i, v in enumerate(satir)])
            tb = Table(veri, colWidths=kolon, repeatRows=1)
            stil = [('BACKGROUND', (0, 0), (-1, 0), lacivert), ('LINEBELOW', (0, 0), (-1, -1), 0.3, gri),
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                    ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f8f9fa')])]
            if sec.get('toplam'):
                stil.append(('BACKGROUND', (0, -1), (-1, -1), acik))
            for r in sec.get('vurgu', ()):
                stil.append(('BACKGROUND', (0, r + 1), (-1, r + 1), sari))
            tb.setStyle(TableStyle(stil))
            akis.append(KeepTogether((bekleyen or []) + [tb]) if len(satirlar) < 25 else tb)
            if bekleyen and len(satirlar) >= 25:
                akis.insert(len(akis) - 1, KeepTogether(bekleyen))
            bekleyen = None
            akis.append(Spacer(1, 4 * mm))
        elif tur == 'ozet':
            sag = TA_RIGHT
            veri = [[P(a, 9, k), P(para(v), 9, k, sag)] for a, v, k in b[1]]
            if b[2]:
                a, v, _ = b[1][-1]
                veri[-1] = [P(a, 10, True, renk=colors.white), P(para(v), 10, True, sag, colors.white)]
            tb = Table(veri, colWidths=[None, 48 * mm])
            stil = [('LINEBELOW', (0, 0), (-1, -1), 0.3, gri)]
            if b[2]:
                stil.append(('BACKGROUND', (0, -1), (-1, -1), lacivert))
            tb.setStyle(TableStyle(stil))
            akis.append(KeepTogether((bekleyen or []) + [tb]))
            bekleyen = None
            akis.append(Spacer(1, 4 * mm))
        elif tur == 'metin':
            parca = [P(b[1], 8 if b[2] else 9, renk=colors.HexColor('#444444'))]
            akis.append(KeepTogether((bekleyen or []) + parca))
            bekleyen = None
            akis.append(Spacer(1, 4 * mm))
        elif tur == 'vurgu':
            kutu = Table([[P(b[1], 11, True, renk=colors.white), P(b[2], 14, True, TA_RIGHT, colors.white)],
                          [P(b[3], 8, renk=colors.HexColor('#f1c40f')), '']],
                         colWidths=[genislik - 70 * mm, 70 * mm])
            kutu.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, -1), lacivert), ('SPAN', (0, 1), (1, 1)),
                                      ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'), ('TOPPADDING', (0, 0), (-1, 0), 6),
                                      ('BOTTOMPADDING', (0, 1), (-1, 1), 6), ('LEFTPADDING', (0, 0), (-1, -1), 8),
                                      ('RIGHTPADDING', (0, 0), (-1, -1), 8)]))
            if akis and isinstance(akis[-1], Spacer):
                akis.pop()
            akis += [Spacer(1, 1 * mm), KeepTogether([kutu]), Spacer(1, 3 * mm)]
        elif tur == 'onay':
            onay = Table([[P('MÜŞTERİ ONAYI', 10, True, renk=lacivert), ''],
                          [P('Firma Kaşe / İmza:'), P('_______________________________')],
                          [P('Yetkili Ad Soyad:'), P('_______________________________')],
                          [P('Tarih:'), P('_______________________________')]], colWidths=[45 * mm, None])
            onay.setStyle(TableStyle([('TOPPADDING', (0, 1), (-1, -1), 8)]))
            akis += [Spacer(1, 4 * mm), KeepTogether([onay])]

    def sayfa_alti(tuval, doc):
        tuval.saveState()
        tuval.setFont(font, 7)
        tuval.setFillColor(colors.HexColor('#888888'))
        tuval.drawString(16 * mm, 9 * mm, t('%s - %s' % (ayar.get('firma_adi', ''), belge_no)))
        tuval.drawRightString(A4[0] - 16 * mm, 9 * mm, t('Sayfa %d' % doc.page))
        tuval.restoreState()

    belge = SimpleDocTemplate(yol, pagesize=A4, leftMargin=16 * mm, rightMargin=16 * mm, topMargin=14 * mm,
                              bottomMargin=16 * mm, title=t(belge_no), author=t(ayar.get('firma_adi', '')))
    belge.build(akis, onFirstPage=sayfa_alti, onLaterPages=sayfa_alti)
    return yol


# ======================================================================
# Excel
# ======================================================================
def _excel_ciz(yol, B, ayar, sayfa_adi='Teklif'):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    wb = Workbook()
    ws = wb.active
    ws.title = sayfa_adi
    lacivert = PatternFill('solid', start_color='2C3E50')
    acik = PatternFill('solid', start_color='ECF0F1')
    sari = PatternFill('solid', start_color='FEF5D4')
    beyaz = Font(color='FFFFFF', bold=True)
    ince = Border(bottom=Side(style='thin', color='CCCCCC'))
    kaydir = Alignment(wrap_text=True, vertical='top')
    sag = Alignment(horizontal='right', vertical='top', wrap_text=True)
    for kol, gen in zip('ABCDEFG', (26, 30, 16, 18, 22, 18, 16)):
        ws.column_dimensions[kol].width = gen

    ws['A1'] = ayar.get('firma_adi', '')
    ws['A1'].font = Font(size=16, bold=True, color='2C3E50')
    r = 2
    for x in (ayar.get('firma_adres'), ayar.get('firma_iletisim')):
        if x:
            ws.cell(r, 1, x).font = Font(size=9, color='555555')
            r += 1
    r += 1
    for b in B:
        tur = b[0]
        if tur == 'baslik':
            ws.cell(r, 1, b[1] + ('  -  ' + b[2] if b[2] else '')).font = beyaz
            for c in range(1, 8):
                ws.cell(r, c).fill = lacivert
            r += 2
        elif tur == 'bilgi':
            for i in range(0, len(b[1]), 2):
                for j, (e, d) in enumerate(b[1][i:i + 2]):
                    ws.cell(r, 1 + j * 3, e).font = Font(bold=True)
                    ws.cell(r, 2 + j * 3, d)
                r += 1
            r += 1
        elif tur == 'bolum':
            ws.cell(r, 1, b[1]).font = Font(size=12, bold=True, color='2C3E50')
            r += 1
        elif tur == 'tablo':
            _, basliklar, satirlar, sec = b
            for c, h in enumerate(basliklar, 1):
                hc = ws.cell(r, c, h)
                hc.font, hc.fill, hc.alignment = beyaz, lacivert, kaydir
            r += 1
            for n, satir in enumerate(satirlar):
                son = sec.get('toplam') and n == len(satirlar) - 1
                for c, v in enumerate(satir, 1):
                    hc = ws.cell(r, c, v)
                    hc.border = ince
                    hc.alignment = sag if (c - 1) in sec.get('sag', ()) else kaydir
                    if son:
                        hc.font, hc.fill = Font(bold=True), acik
                    elif n in sec.get('vurgu', ()):
                        hc.fill = sari
                r += 1
            r += 1
        elif tur == 'ozet':
            for n, (a, v, k) in enumerate(b[1]):
                ws.cell(r, 1, a).font = Font(bold=k)
                hc = ws.cell(r, 4, round(v, 2))
                hc.number_format = '#,##0.00 "TL"'
                hc.font = Font(bold=k)
                if b[2] and n == len(b[1]) - 1:
                    for c in range(1, 5):
                        ws.cell(r, c).fill = lacivert
                    ws.cell(r, 1).font = ws.cell(r, 4).font = beyaz
                r += 1
            r += 1
        elif tur == 'metin':
            for satir in b[1].split('\n'):
                ws.cell(r, 1, satir).font = Font(size=9, color='444444')
                r += 1
            r += 1
        elif tur == 'vurgu':
            for c in range(1, 5):
                ws.cell(r, c).fill = lacivert
                ws.cell(r + 1, c).fill = lacivert
            ws.cell(r, 1, b[1]).font = Font(size=12, bold=True, color='FFFFFF')
            hc = ws.cell(r, 4, b[2])
            hc.font, hc.alignment = Font(size=13, bold=True, color='FFFFFF'), Alignment(horizontal='right')
            ws.cell(r + 1, 1, b[3]).font = Font(size=9, color='F1C40F')
            r += 3
        elif tur == 'onay':
            ws.cell(r, 1, 'MÜŞTERİ ONAYI').font = Font(bold=True, color='2C3E50')
            for x in ('Firma Kaşe / İmza:', 'Yetkili Ad Soyad:', 'Tarih:'):
                r += 2
                ws.cell(r, 1, x)
                ws.cell(r, 2, '______________________________')
            r += 1
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    wb.save(yol)
    return yol


# ======================================================================
# Dis arayuz
# ======================================================================
def pdf_olustur(yol, girdi, sonuc, ayar):
    return _pdf_ciz(yol, bolumler(girdi, sonuc, ayar), ayar, 'Teklif ' + str(girdi.get('teklif_no') or ''))


def excel_olustur(yol, girdi, sonuc, ayar):
    return _excel_ciz(yol, bolumler(girdi, sonuc, ayar), ayar)


def karsilastirma_pdf(yol, teklifler, ayar):
    return _pdf_ciz(yol, karsilastirma_bolumleri(teklifler, ayar), ayar, 'Teklif Seçenekleri')


def karsilastirma_excel(yol, teklifler, ayar):
    return _excel_ciz(yol, karsilastirma_bolumleri(teklifler, ayar), ayar, 'Seçenekler')


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
