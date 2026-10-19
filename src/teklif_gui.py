"""💼 DIA Teklif Motoru penceresi.

Veri temizleme programindan (Teklif butonu) veya tek basina (teklif_main.py -> RACA_Teklif.exe) acilir.
"""
import os
import re
import subprocess
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import teklif_cikti
from teklif_hesap import hesapla, ybh_hesapla
from teklif_veri import TeklifVeri, AYARLAR, varsayilan_klasor, veri_klasoru_degistir

LACIVERT = '#2c3e50'
_KATLA = str.maketrans('çğıöşüÇĞİÖŞÜ', 'cgiosuCGIOSU')


def katla(s):
    return str(s).replace('İ', 'i').replace('I', 'ı').lower().translate(_KATLA)


def sayi_oku(metin, bos=None):
    """'1.250,50' / '1250.5' / '%15' -> float; bos -> bos"""
    s = str(metin).strip().replace('%', '').replace('TL', '').replace(' ', '')
    if not s:
        return bos
    if ',' in s:
        s = s.replace('.', '').replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return bos


def para(x):
    return teklif_cikti.para(x).replace(' TL', '')


def dosya_ac(yol):
    try:
        if sys.platform.startswith('win'):
            os.startfile(yol)  # noqa
        elif sys.platform == 'darwin':
            subprocess.Popen(['open', yol])
        else:
            subprocess.Popen(['xdg-open', yol])
    except Exception:
        pass


def _dugme(ust, metin, komut, renk='#3498db', **kw):
    return tk.Button(ust, text=metin, command=komut, bg=renk, fg='white', font=('Arial', 10, 'bold'),
                     padx=10, pady=4, relief=tk.FLAT, cursor='hand2', **kw)


# ======================================================================
# Ana teklif penceresi
# ======================================================================
class TeklifPenceresi:
    KOLONLAR = [('kod', 'Paket Kodu', 150), ('ad', 'Paket Adı', 230), ('tur', 'Tür', 80), ('adet', 'Adet ✎', 60),
                ('birim', 'Birim Fiyat', 95), ('kull', 'Kull. %', 55), ('toplam', 'Toplam', 105),
                ('ind_yuzde', 'İnd. % ✎', 65), ('ind_tutar', 'İnd. Tutar ✎', 90), ('net', 'Net (KDV Hr.)', 110),
                ('dahil', 'KDV Dahil', 110), ('saat', 'A/S ✎', 55)]
    DUZENLENIR = {'adet': 'adet', 'ind_yuzde': 'ind_yuzde', 'ind_tutar': 'ind_tutar', 'saat': 'saat'}

    def __init__(self, parent=None, veri_yolu=None):
        self.kendi_kok = parent is None
        self.pencere = tk.Tk() if self.kendi_kok else tk.Toplevel(parent)
        self.pencere.title('DIA Teklif Motoru')
        self.pencere.geometry('1360x860')
        self.pencere.configure(bg='#f0f0f0')
        try:
            self.veri = TeklifVeri(veri_yolu or os.path.join(varsayilan_klasor(), 'teklif.db'))
        except Exception as e:
            messagebox.showerror('Veri açılamadı', str(e), parent=self.pencere)
            self.pencere.destroy()
            return
        self.pencere.protocol('WM_DELETE_WINDOW', self.kapat)
        self.kalemler = []
        self.teklif_no = None
        self.tarih = None
        self.sonuc = None
        self._yukleniyor = False
        self.verileri_yenile()
        self.arayuz()
        self.yeni(sor=False)
        if self.veri.bos_mu():
            self.pencere.after(300, self.ilk_kurulum)

    def verileri_yenile(self):
        self.paketler = self.veri.paketler()
        self.ayar = self.veri.ayarlar()
        self.sabit_fiyatlar = self.veri.sabit_kalemler()
        self.kampanya_listesi = self.veri.kampanyalar()
        self.rakip_listesi = self.veri.rakipler()

    # ------------------------------------------------------------ arayuz
    def arayuz(self):
        p = self.pencere
        ust = tk.Frame(p, bg=LACIVERT, pady=6)
        ust.pack(fill=tk.X)
        tk.Label(ust, text='💼 DIA TEKLİF MOTORU', font=('Arial', 14, 'bold'), bg=LACIVERT,
                 fg='white').pack(side=tk.LEFT, padx=12)
        self.no_lbl = tk.Label(ust, text='', font=('Arial', 11, 'bold'), bg=LACIVERT, fg='#f1c40f')
        self.no_lbl.pack(side=tk.LEFT, padx=10)
        for metin, komut, renk in [('⚙ Veriler / Ayarlar', self.veriler_ac, '#7f8c8d'),
                                   ('🔧 Yıllık Bakım (YBH)', self.ybh_ac, '#8e44ad'),
                                   ('📂 Teklif Geçmişi', self.gecmis_ac, '#16a085'),
                                   ('📊 Excel', self.excel_kaydet, '#27ae60'), ('📄 PDF', self.pdf_kaydet, '#c0392b'),
                                   ('💾 Kaydet', self.kaydet, '#2980b9'), ('＋ Yeni Teklif', self.yeni, '#d35400')]:
            _dugme(ust, metin, komut, renk).pack(side=tk.RIGHT, padx=3)

        # Musteri bilgileri
        mf = tk.LabelFrame(p, text=' Müşteri ', font=('Arial', 10, 'bold'), bg='#f0f0f0', padx=8, pady=6)
        mf.pack(fill=tk.X, padx=10, pady=(8, 4))
        self.musteri = tk.StringVar()
        self.masaustu = tk.StringVar(value='1')
        self.mobil = tk.StringVar(value='0')
        self.tip_secim = tk.StringVar(value='OTOMATİK')
        tk.Label(mf, text='Müşteri:', bg='#f0f0f0').grid(row=0, column=0, sticky='w')
        tk.Entry(mf, textvariable=self.musteri, width=48, font=('Arial', 11)).grid(row=0, column=1, padx=5)
        tk.Label(mf, text='Masaüstü kullanıcı:', bg='#f0f0f0').grid(row=0, column=2, padx=(15, 2))
        tk.Spinbox(mf, from_=0, to=9999, textvariable=self.masaustu, width=6, font=('Arial', 11)).grid(row=0, column=3)
        tk.Label(mf, text='Mobil kullanıcı:', bg='#f0f0f0').grid(row=0, column=4, padx=(15, 2))
        tk.Spinbox(mf, from_=0, to=9999, textvariable=self.mobil, width=6, font=('Arial', 11)).grid(row=0, column=5)
        tk.Label(mf, text='Paket tipi:', bg='#f0f0f0').grid(row=0, column=6, padx=(15, 2))
        ttk.Combobox(mf, textvariable=self.tip_secim, values=['OTOMATİK', 'STANDART', 'ENTERPRISE'], width=12,
                     state='readonly').grid(row=0, column=7)
        self.tip_lbl = tk.Label(mf, text='', font=('Arial', 11, 'bold'), bg='#f0f0f0', fg='#c0392b')
        self.tip_lbl.grid(row=0, column=8, padx=8)
        for v in (self.masaustu, self.mobil, self.tip_secim):
            v.trace_add('write', lambda *a: self.hesapla())

        govde = tk.Frame(p, bg='#f0f0f0')
        govde.pack(fill=tk.BOTH, expand=True, padx=10, pady=4)

        # Kalemler
        kf = tk.LabelFrame(govde, text=' 1) Lisans Kalemleri  (✎ kolonlara çift tıklayıp değiştirin; boş = otomatik) ',
                           font=('Arial', 10, 'bold'), bg='#f0f0f0')
        kf.pack(fill=tk.BOTH, expand=True)
        btn = tk.Frame(kf, bg='#f0f0f0')
        btn.pack(fill=tk.X, pady=3)
        _dugme(btn, '＋ Paket Ekle', self.paket_ekle, '#27ae60').pack(side=tk.LEFT, padx=3)
        _dugme(btn, '🗑 Seçileni Sil', self.kalem_sil, '#e74c3c').pack(side=tk.LEFT, padx=3)
        _dugme(btn, '▲', lambda: self.kalem_tasi(-1), '#95a5a6').pack(side=tk.LEFT, padx=1)
        _dugme(btn, '▼', lambda: self.kalem_tasi(1), '#95a5a6').pack(side=tk.LEFT, padx=1)
        self.aciklama_lbl = tk.Label(btn, text='', bg='#f0f0f0', fg='#555', font=('Arial', 9), anchor='w',
                                     justify='left', wraplength=850)
        self.aciklama_lbl.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)

        tf = tk.Frame(kf)
        tf.pack(fill=tk.BOTH, expand=True)
        self.tree = ttk.Treeview(tf, columns=[k for k, _, _ in self.KOLONLAR], show='headings', height=10)
        for k, b, w in self.KOLONLAR:
            self.tree.heading(k, text=b)
            self.tree.column(k, width=w, anchor='w' if k in ('kod', 'ad', 'tur') else 'e')
        sb = ttk.Scrollbar(tf, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.pack(fill=tk.BOTH, expand=True)
        self.tree.tag_configure('hata', background='#fadbd8')
        self.tree.tag_configure('elle', background='#fef9e7')
        self.tree.tag_configure('toplam', background='#d6eaf8', font=('Arial', 9, 'bold'))
        self.tree.bind('<Double-1>', self.hucre_duzenle)
        self.tree.bind('<Delete>', lambda e: self.kalem_sil())
        self.tree.bind('<<TreeviewSelect>>', self._secim_aciklamasi)

        # Alt: secenekler | ozet | odeme+rakip
        alt = tk.Frame(govde, bg='#f0f0f0')
        alt.pack(fill=tk.X, pady=(6, 0))

        sec = tk.LabelFrame(alt, text=' 2) Seçenekler ', font=('Arial', 10, 'bold'), bg='#f0f0f0', padx=6, pady=4)
        sec.pack(side=tk.LEFT, fill=tk.BOTH)
        tk.Label(sec, text='Sabit kalemler:', bg='#f0f0f0', font=('Arial', 9, 'bold')).pack(anchor='w')
        self.sabit_cerceve = tk.Frame(sec, bg='#f0f0f0')
        self.sabit_cerceve.pack(anchor='w')
        self.sabit_vars = {}
        kamp = tk.Frame(sec, bg='#f0f0f0')
        kamp.pack(anchor='w', pady=(6, 0))
        tk.Label(kamp, text='Kampanya:', bg='#f0f0f0', font=('Arial', 9, 'bold')).grid(row=0, column=0, sticky='w')
        self.kampanya = tk.StringVar(value='YOK')
        self.kampanya_cb = ttk.Combobox(kamp, textvariable=self.kampanya, width=22, state='readonly')
        self.kampanya_cb.grid(row=0, column=1, padx=3)
        self.kampanya_cb.bind('<<ComboboxSelected>>', self._kampanya_secildi)
        tk.Label(kamp, text='İndirim %:', bg='#f0f0f0').grid(row=0, column=2)
        self.kampanya_yuzde = tk.StringVar(value='0')
        tk.Entry(kamp, textvariable=self.kampanya_yuzde, width=6).grid(row=0, column=3, padx=3)
        tk.Label(kamp, text='Maliyet (DIA\'ya ödenen):', bg='#f0f0f0').grid(row=1, column=0, columnspan=2,
                                                                           sticky='w', pady=(6, 0))
        self.maliyet = tk.StringVar(value='0')
        tk.Entry(kamp, textvariable=self.maliyet, width=14).grid(row=1, column=2, columnspan=2, sticky='w', pady=(6, 0))
        for v in (self.kampanya_yuzde, self.maliyet):
            v.trace_add('write', lambda *a: self.hesapla())
        tk.Label(sec, text='Teklife özel not:', bg='#f0f0f0', font=('Arial', 9, 'bold')).pack(anchor='w', pady=(6, 0))
        self.notlar = tk.Text(sec, width=48, height=3, font=('Arial', 9))
        self.notlar.pack(anchor='w')

        oz = tk.LabelFrame(alt, text=' 3) Özet (KDV Hariç / KDV / KDV Dahil) ', font=('Arial', 10, 'bold'),
                           bg='white', padx=6, pady=4)
        oz.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=6)
        self.ozet_lbl = {}
        satirlar = [('lisans', 'Lisans'), ('uyarlama', 'Uyarlama'), ('sabit', 'Sabit Kalemler'), ('ara', 'ARA TOPLAM'),
                    ('kampanya', 'Kampanya İndirimi'), ('net', 'NET TOPLAM'), ('kar', 'Kâr (net - maliyet)')]
        for c, b in enumerate(['', 'KDV Hariç', 'KDV', 'KDV Dahil']):
            tk.Label(oz, text=b, bg='white', font=('Arial', 9, 'bold'), fg='#555').grid(row=0, column=c, sticky='e',
                                                                                       padx=6)
        for r, (k, b) in enumerate(satirlar, 1):
            kalin = k in ('ara', 'net')
            tk.Label(oz, text=b, bg='white', font=('Arial', 10, 'bold' if kalin else 'normal')).grid(
                row=r, column=0, sticky='w')
            self.ozet_lbl[k] = []
            for c in range(1, 4):
                lbl = tk.Label(oz, text='', bg='white', font=('Arial', 10, 'bold' if kalin else 'normal'))
                lbl.grid(row=r, column=c, sticky='e', padx=6)
                self.ozet_lbl[k].append(lbl)
        self.uyarlama_detay = tk.Label(oz, text='', bg='white', fg='#555', font=('Arial', 8), justify='left')
        self.uyarlama_detay.grid(row=len(satirlar) + 1, column=0, columnspan=4, sticky='w', pady=(4, 0))
        self.genel_lbl = tk.Label(oz, text='', bg=LACIVERT, fg='white', font=('Arial', 13, 'bold'), pady=4)
        self.genel_lbl.grid(row=len(satirlar) + 2, column=0, columnspan=4, sticky='we', pady=(4, 0))

        sag = tk.LabelFrame(alt, text=' 4) Ödeme Planı ve Rakip ', font=('Arial', 10, 'bold'), bg='white',
                            padx=6, pady=4)
        sag.pack(side=tk.LEFT, fill=tk.BOTH)
        self.odeme_lbl = tk.Label(sag, text='', bg='white', font=('Consolas', 9), justify='left', anchor='nw')
        self.odeme_lbl.pack(anchor='w')
        self.rakip_lbl = tk.Label(sag, text='', bg='white', font=('Consolas', 9), justify='left', anchor='nw')
        self.rakip_lbl.pack(anchor='w', pady=(6, 0))
        self.oneri_lbl = tk.Label(sag, text='', bg='white', font=('Arial', 10, 'bold'))
        self.oneri_lbl.pack(anchor='w')

        self.uyari_lbl = tk.Label(p, text='', bg='#f0f0f0', fg='#c0392b', font=('Arial', 9, 'bold'), anchor='w',
                                  justify='left')
        self.uyari_lbl.pack(fill=tk.X, padx=12, pady=(0, 4))
        self.durum_lbl = tk.Label(p, text='', bg='#f0f0f0', fg='#555', font=('Arial', 8), anchor='w')
        self.durum_lbl.pack(fill=tk.X, padx=12, pady=(0, 4))
        self._secenekleri_doldur()

    def _secenekleri_doldur(self):
        for w in self.sabit_cerceve.winfo_children():
            w.destroy()
        eski = {ad: v.get() for ad, v in self.sabit_vars.items()}
        self.sabit_vars = {}
        for ad, fiyat in self.sabit_fiyatlar.items():
            v = tk.BooleanVar(value=eski.get(ad, False))
            tk.Checkbutton(self.sabit_cerceve, text='%s  (%s)' % (ad, para(fiyat)), variable=v, bg='#f0f0f0',
                           command=self.hesapla).pack(anchor='w')
            self.sabit_vars[ad] = v
        if not self.sabit_fiyatlar:
            tk.Label(self.sabit_cerceve, text='(tanımlı değil: Veriler > Sabit Kalemler)', bg='#f0f0f0',
                     fg='#888').pack(anchor='w')
        self.kampanya_cb['values'] = [k['ad'] for k in self.kampanya_listesi]
        self.durum_lbl.config(text='Veri: %s   |   %d paket   |   Son yedek: %s' % (
            self.veri.yol, len(self.paketler),
            self.veri.yedekler()[0][1] if self.veri.yedekler() else 'yok'))

    def _kampanya_secildi(self, _e=None):
        for k in self.kampanya_listesi:
            if k['ad'] == self.kampanya.get():
                self.kampanya_yuzde.set('%g' % k['yuzde'])
        self.hesapla()

    # ------------------------------------------------------------ hesap
    def girdi(self):
        tip = self.tip_secim.get().replace('İ', 'I')
        return {'teklif_no': self.teklif_no, 'tarih': self.tarih, 'musteri': self.musteri.get().strip(),
                'masaustu': int(sayi_oku(self.masaustu.get(), 0)), 'mobil': int(sayi_oku(self.mobil.get(), 0)),
                'paket_tipi': 'OTOMATIK' if tip.startswith('OTOMAT') else tip,
                'kalemler': [dict(k) for k in self.kalemler],
                'sabitler': [ad for ad, v in self.sabit_vars.items() if v.get()],
                'kampanya': self.kampanya.get(), 'kampanya_yuzde': sayi_oku(self.kampanya_yuzde.get(), 0),
                'maliyet': sayi_oku(self.maliyet.get(), 0), 'notlar': self.notlar.get('1.0', 'end').strip()}

    def hesapla(self):
        if self._yukleniyor or not hasattr(self, 'tree'):
            return
        g = self.girdi()
        s = self.sonuc = hesapla(g, self.paketler, self.ayar, self.sabit_fiyatlar, self.rakip_listesi)
        self.tip_lbl.config(text='→ ' + s['paket_tipi'])
        secili = self.tree.selection()
        self.tree.delete(*self.tree.get_children())
        sira = {x['sira']: x for x in s['satirlar']}
        for i, k in enumerate(self.kalemler):
            x = sira.get(i)
            if not x:
                self.tree.insert('', 'end', iid=str(i), values=(k['kod'], '(paket listesinde yok)'), tags=('hata',))
                continue
            elle = any(k.get(a) not in (None, '') for a in ('adet', 'saat')) or \
                any(k.get(a) for a in ('ind_yuzde', 'ind_tutar'))
            self.tree.insert('', 'end', iid=str(i), tags=('elle',) if elle else (), values=(
                x['kod'], x['ad'], x['tur'], x['adet'], para(x['birim']), '%g' % x['kull_yuzde'], para(x['toplam']),
                '%g' % x['ind_yuzde'] if x['ind_yuzde'] else '', para(x['ind_tutar']) if x['ind_tutar'] else '',
                para(x['net']), para(x['dahil']), '%g' % x['saat']))
        li = s['lisans']
        self.tree.insert('', 'end', iid='toplam', tags=('toplam',), values=(
            'TOPLAM', '%d kalem' % len(s['satirlar']), '', '', '', '', para(li['toplam']), '',
            para(li['indirim']) if li['indirim'] else '', para(li['net']), para(li['dahil']),
            '%g' % s['uyarlama']['saat']))
        for iid in secili:
            if self.tree.exists(iid):
                self.tree.selection_add(iid)

        kdv = s['kdv_orani'] / 100

        def yaz(k, net):
            v = [net, net * kdv, net * (1 + kdv)]
            for lbl, deger in zip(self.ozet_lbl[k], v):
                lbl.config(text=para(deger))
        yaz('lisans', li['net'])
        yaz('uyarlama', s['uyarlama']['net'])
        yaz('sabit', s['sabit']['net'])
        yaz('ara', s['ara_toplam'])
        yaz('kampanya', -s['kampanya']['tutar'])
        yaz('net', s['net_toplam'])
        for lbl, deger in zip(self.ozet_lbl['kar'], [s['kar']['brut'], None, None]):
            lbl.config(text=para(deger) if deger is not None else '')
        self.ozet_lbl['kar'][1].config(text='marj %%%g' % s['kar']['marj'] if s['kar']['maliyet'] else '')
        u = s['uyarlama']
        self.uyarlama_detay.config(text='Uyarlama = lisans × %%%g (%s) + %g saat × %s (%s)%s' % (
            self.ayar['lisans_yuzdesi'], para(u['lisans_payi']), u['saat'], para(self.ayar['saatlik_ucret']),
            para(u['saat_bedeli']), ('   → ' + u['sinir'] + ' uygulandı') if u['sinir'] else ''))
        self.genel_lbl.config(text='GENEL TOPLAM (KDV Dahil):  %s TL' % para(s['genel_toplam']))
        self.odeme_lbl.config(text='\n'.join('%-20s %16s' % (o['plan'], para(o['dahil']) + (
            ' x%d' % o['adet'] if o['adet'] > 1 else '')) for o in s['odeme']))
        r = s['rakip']
        satir = ['%-12s %14s' % ('RACA (biz)', para(r['biz']))]
        satir += ['%-12s %14s %+6.1f%%' % (x['firma'][:12], para(x['toplam']), x['fark_yuzde']) for x in r['satirlar']]
        if r['ortalama']:
            satir.append('%-12s %14s' % ('Ortalama', para(r['ortalama'])))
        baslik = 'Uyarlama karşılaştırması (%s):' % ('aynı %g saat × rakip ücreti' % s['uyarlama']['saat']
                                                     if self.ayar.get('rakip_ayni_saat') else 'rakip tablosundaki saat')
        self.rakip_lbl.config(text=baslik + '\n' + '\n'.join(satir))
        self.oneri_lbl.config(text=r['oneri'] + ('  (%+g%%)' % r['konum'] if r['konum'] is not None else ''),
                              fg='#c0392b' if 'Pahalı' in r['oneri'] else '#27ae60')
        self.uyari_lbl.config(text='\n'.join('⚠ ' + x for x in s['uyarilar']))

    # ------------------------------------------------------------ kalemler
    def paket_ekle(self):
        if not self.paketler:
            messagebox.showinfo('Paket yok', 'Önce Veriler > Excel\'den İçe Aktar ile paket listesini yükleyin.',
                                parent=self.pencere)
            return
        PaketSecici(self.pencere, self.paketler, self.ayar, self._paket_eklendi)

    def _paket_eklendi(self, kod):
        self.kalemler.append({'kod': kod, 'adet': None, 'ind_yuzde': 0, 'ind_tutar': 0, 'saat': None})
        self.hesapla()

    def kalem_sil(self):
        sec = [int(i) for i in self.tree.selection() if i != 'toplam']
        for i in sorted(sec, reverse=True):
            del self.kalemler[i]
        self.hesapla()

    def kalem_tasi(self, yon):
        sec = [int(i) for i in self.tree.selection() if i != 'toplam']
        if len(sec) != 1:
            return
        i, j = sec[0], sec[0] + yon
        if 0 <= j < len(self.kalemler):
            self.kalemler[i], self.kalemler[j] = self.kalemler[j], self.kalemler[i]
            self.hesapla()
            self.tree.selection_set(str(j))

    def _secim_aciklamasi(self, _e=None):
        sec = [i for i in self.tree.selection() if i != 'toplam']
        if sec:
            p = self.paketler.get(self.kalemler[int(sec[0])]['kod'], {})
            self.aciklama_lbl.config(text=(p.get('aciklama') or '').replace('\n', ' ')[:400])

    def hucre_duzenle(self, olay):
        iid = self.tree.identify_row(olay.y)
        kolon = self.tree.identify_column(olay.x)
        if not iid or iid == 'toplam':
            return
        ad = self.KOLONLAR[int(kolon[1:]) - 1][0]
        if ad not in self.DUZENLENIR:
            return
        x, y, w, h = self.tree.bbox(iid, kolon)
        alan = self.DUZENLENIR[ad]
        kalem = self.kalemler[int(iid)]
        giris = tk.Entry(self.tree, justify='right')
        mevcut = kalem.get(alan)
        giris.insert(0, '' if mevcut in (None, 0) else '%g' % mevcut)
        giris.place(x=x, y=y, width=w, height=h)
        giris.focus_set()
        giris.select_range(0, 'end')

        def bitir(_e=None, kaydet=True):
            if kaydet:
                deger = sayi_oku(giris.get(), None)
                if alan in ('ind_yuzde', 'ind_tutar'):
                    deger = deger or 0
                elif alan == 'adet' and deger is not None:
                    deger = int(deger)
                kalem[alan] = deger
            giris.destroy()
            self.hesapla()
        giris.bind('<Return>', bitir)
        giris.bind('<FocusOut>', bitir)
        giris.bind('<Escape>', lambda e: bitir(kaydet=False))

    # ------------------------------------------------------------ teklif islemleri
    def yeni(self, sor=True):
        if sor and self.kalemler and not messagebox.askyesno('Yeni teklif', 'Ekrandaki teklif temizlensin mi?\n'
                                                             '(Kaydetmediyseniz kaybolur)', parent=self.pencere):
            return
        self.teklif_no = None
        self.tarih = None
        self.kalemler = []
        self._yukleniyor = True
        self.musteri.set('')
        self.masaustu.set('1')
        self.mobil.set('0')
        self.tip_secim.set('OTOMATİK')
        self.kampanya.set('YOK')
        self.kampanya_yuzde.set('0')
        self.maliyet.set('0')
        self.notlar.delete('1.0', 'end')
        for v in self.sabit_vars.values():
            v.set(False)
        self._yukleniyor = False
        self.no_lbl.config(text='(yeni teklif - kaydedilmedi)')
        self.hesapla()

    def yukle(self, girdi):
        self._yukleniyor = True
        self.teklif_no = girdi.get('teklif_no')
        self.tarih = girdi.get('tarih')
        self.musteri.set(girdi.get('musteri', ''))
        self.masaustu.set(str(girdi.get('masaustu', 0)))
        self.mobil.set(str(girdi.get('mobil', 0)))
        tip = girdi.get('paket_tipi', 'OTOMATIK')
        self.tip_secim.set('OTOMATİK' if tip == 'OTOMATIK' else tip)
        self.kalemler = [dict(k) for k in girdi.get('kalemler', [])]
        for ad, v in self.sabit_vars.items():
            v.set(ad in girdi.get('sabitler', []))
        self.kampanya.set(girdi.get('kampanya') or 'YOK')
        self.kampanya_yuzde.set('%g' % (girdi.get('kampanya_yuzde') or 0))
        self.maliyet.set('%g' % (girdi.get('maliyet') or 0))
        self.notlar.delete('1.0', 'end')
        self.notlar.insert('1.0', girdi.get('notlar', ''))
        self._yukleniyor = False
        self.no_lbl.config(text='Teklif No: ' + str(self.teklif_no))
        self.hesapla()

    def kaydet(self, sessiz=False):
        if not self.musteri.get().strip():
            messagebox.showwarning('Müşteri', 'Müşteri adını girin.', parent=self.pencere)
            return None
        if not self.kalemler:
            messagebox.showwarning('Kalem yok', 'En az bir paket ekleyin.', parent=self.pencere)
            return None
        self.hesapla()
        g = self.girdi()
        if not g['tarih']:
            from datetime import datetime
            g['tarih'] = self.tarih = datetime.now().strftime('%Y-%m-%d')
        self.teklif_no = self.veri.teklif_kaydet(g, self.sonuc, self.teklif_no)
        self.no_lbl.config(text='Teklif No: ' + self.teklif_no + '  ✔ kaydedildi')
        if not sessiz:
            messagebox.showinfo('Kaydedildi', self.teklif_no + ' teklif geçmişine kaydedildi.', parent=self.pencere)
        return self.teklif_no

    def _cikti_yolu(self, uzanti):
        if not self.kaydet(sessiz=True):
            return None
        klasor = os.path.join(self.veri.klasor, 'Teklifler')
        os.makedirs(klasor, exist_ok=True)
        musteri = re.sub(r'[^\w\- ]', '', self.musteri.get())[:40].strip()
        return filedialog.asksaveasfilename(parent=self.pencere, initialdir=klasor, defaultextension=uzanti,
                                            initialfile='%s_%s%s' % (self.teklif_no, musteri, uzanti),
                                            filetypes=[(uzanti.upper()[1:], '*' + uzanti)])

    def pdf_kaydet(self):
        yol = self._cikti_yolu('.pdf')
        if not yol:
            return
        try:
            teklif_cikti.pdf_olustur(yol, self.girdi(), self.sonuc, self.ayar)
        except ImportError:
            messagebox.showerror('PDF', "PDF için 'reportlab' kurulu değil.", parent=self.pencere)
            return
        except Exception as e:
            messagebox.showerror('PDF oluşturulamadı', str(e), parent=self.pencere)
            return
        dosya_ac(yol)

    def excel_kaydet(self):
        yol = self._cikti_yolu('.xlsx')
        if not yol:
            return
        try:
            teklif_cikti.excel_olustur(yol, self.girdi(), self.sonuc, self.ayar)
        except Exception as e:
            messagebox.showerror('Excel oluşturulamadı', str(e), parent=self.pencere)
            return
        dosya_ac(yol)

    # ------------------------------------------------------------ alt pencereler
    def gecmis_ac(self):
        GecmisPenceresi(self)

    def veriler_ac(self):
        VeriPenceresi(self)

    def ybh_ac(self):
        if not self.paketler:
            messagebox.showinfo('Paket yok', 'Önce paket listesini içe aktarın.', parent=self.pencere)
            return
        YbhPenceresi(self)

    def veri_degisti(self):
        self.verileri_yenile()
        self._secenekleri_doldur()
        self.hesapla()

    def ilk_kurulum(self):
        if messagebox.askyesno(
                'İlk kurulum', 'Paket listesi ve ayarlar henüz yok.\n\nMevcut teklif Excel\'iniz '
                '(RACA_DIA_Teklif_Motoru.xlsx) içe aktarılsın mı?\n\nPaket listesi, adam/saat, ayarlar, kampanyalar ve '
                'rakipler alınır.\n\nVeriler şurada tutulur:\n' + self.veri.klasor, parent=self.pencere):
            VeriPenceresi.excelden_ice_aktar_sor(self)

    def kapat(self):
        try:
            self.veri.kapat()
        except Exception:
            pass
        self.pencere.destroy()


# ======================================================================
# Paket secici
# ======================================================================
class PaketSecici:
    def __init__(self, ana, paketler, ayar, secildi):
        self.paketler = [p for p in paketler.values() if p.get('aktif', 1)]
        self.secildi = secildi
        self.p = tk.Toplevel(ana)
        self.p.title('Paket Ekle')
        self.p.geometry('980x560')
        self.p.transient(ana)
        ust = tk.Frame(self.p, pady=6)
        ust.pack(fill=tk.X, padx=8)
        tk.Label(ust, text='Ara (kod, ad, açıklama):').pack(side=tk.LEFT)
        self.ara = tk.StringVar()
        giris = tk.Entry(ust, textvariable=self.ara, width=40, font=('Arial', 11))
        giris.pack(side=tk.LEFT, padx=6)
        giris.focus_set()
        self.tur = tk.StringVar(value='Tümü')
        turler = ['Tümü'] + sorted({p.get('tur') or '' for p in self.paketler})
        ttk.Combobox(ust, textvariable=self.tur, values=turler, state='readonly', width=16).pack(side=tk.LEFT)
        tk.Label(ust, text='   Çift tık / Enter: ekle (pencere açık kalır)', fg='#555').pack(side=tk.LEFT)
        self.tree = ttk.Treeview(self.p, columns=('kod', 'ad', 'tur', 'std', 'ent', 'saat'), show='headings')
        for k, b, w, a in [('kod', 'Kod', 180, 'w'), ('ad', 'Ad', 320, 'w'), ('tur', 'Tür', 100, 'w'),
                           ('std', 'Std. Fiyat', 100, 'e'), ('ent', 'Ent. Fiyat', 100, 'e'),
                           ('saat', 'A/S Std/Ent', 90, 'e')]:
            self.tree.heading(k, text=b)
            self.tree.column(k, width=w, anchor=a)
        self.tree.pack(fill=tk.BOTH, expand=True, padx=8)
        self.aciklama = tk.Label(self.p, text='', wraplength=940, justify='left', anchor='w', fg='#333')
        self.aciklama.pack(fill=tk.X, padx=8, pady=4)
        self.bilgi = tk.Label(self.p, text='', fg='#27ae60', anchor='w')
        self.bilgi.pack(fill=tk.X, padx=8)
        self.ara.trace_add('write', lambda *a: self.doldur())
        self.tur.trace_add('write', lambda *a: self.doldur())
        self.tree.bind('<Double-1>', self.ekle)
        self.tree.bind('<Return>', self.ekle)
        giris.bind('<Return>', self.ekle)
        giris.bind('<Down>', lambda e: (self.tree.focus_set(), self.tree.selection_set(self.tree.get_children()[:1])))
        self.tree.bind('<<TreeviewSelect>>', self._goster)
        self.doldur()

    def doldur(self):
        aranan = katla(self.ara.get().strip())
        tur = self.tur.get()
        self.tree.delete(*self.tree.get_children())
        for p in self.paketler:
            if tur != 'Tümü' and (p.get('tur') or '') != tur:
                continue
            if aranan and aranan not in katla(p['kod'] + ' ' + (p.get('ad') or '') + ' ' + (p.get('aciklama') or '')):
                continue
            self.tree.insert('', 'end', iid=p['kod'], values=(
                p['kod'], p.get('ad'), p.get('tur'), para(p.get('std_fiyat')), para(p.get('ent_fiyat')),
                '%g / %g' % (p.get('std_saat') or 0, p.get('ent_saat') or 0)))
        ilk = self.tree.get_children()[:1]
        if ilk:
            self.tree.selection_set(ilk)

    def _goster(self, _e=None):
        sec = self.tree.selection()
        if sec:
            p = next((x for x in self.paketler if x['kod'] == sec[0]), {})
            self.aciklama.config(text=(p.get('aciklama') or '').replace('\n', ' '))

    def ekle(self, _e=None):
        for kod in self.tree.selection():
            self.secildi(kod)
            self.bilgi.config(text='Eklendi: ' + kod)


# ======================================================================
# Teklif gecmisi
# ======================================================================
class GecmisPenceresi:
    def __init__(self, ana):
        self.ana = ana
        self.p = tk.Toplevel(ana.pencere)
        self.p.title('Teklif Geçmişi')
        self.p.geometry('1200x600')
        self.p.transient(ana.pencere)
        ust = tk.Frame(self.p, pady=6)
        ust.pack(fill=tk.X, padx=8)
        tk.Label(ust, text='Ara (müşteri / no):').pack(side=tk.LEFT)
        self.ara = tk.StringVar()
        tk.Entry(ust, textvariable=self.ara, width=30).pack(side=tk.LEFT, padx=5)
        self.ara.trace_add('write', lambda *a: self.doldur())
        for metin, komut, renk in [('📂 Aç / Düzenle', self.ac, '#2980b9'), ('📑 Kopyala (yeni no)', self.kopyala,
                                                                            '#8e44ad'),
                                   ('✔ KAZANDI', lambda: self.sonuc('KAZANDI'), '#27ae60'),
                                   ('✖ KAYBETTİ', lambda: self.sonuc('KAYBETTİ'), '#c0392b'),
                                   ('… BEKLİYOR', lambda: self.sonuc('BEKLİYOR'), '#7f8c8d'),
                                   ('🗑 Sil', self.sil, '#e74c3c'), ('📊 Listeyi Excel\'e', self.excele, '#16a085')]:
            _dugme(ust, metin, komut, renk).pack(side=tk.LEFT, padx=2)
        kol = [('no', 'Teklif No', 120), ('tarih', 'Tarih', 85), ('musteri', 'Müşteri', 260), ('paket_tipi', 'Paket', 90),
               ('kullanici', 'Kullanıcı', 140), ('net', 'Net (KDV Hr.)', 110), ('genel_toplam', 'KDV Dahil', 110),
               ('kampanya', 'Kampanya', 100), ('sonuc', 'Sonuç', 85), ('red_nedeni', 'Red Nedeni', 160)]
        self.kol = kol
        self.tree = ttk.Treeview(self.p, columns=[k for k, _, _ in kol], show='headings')
        for k, b, w in kol:
            self.tree.heading(k, text=b)
            self.tree.column(k, width=w, anchor='e' if k in ('net', 'genel_toplam') else 'w')
        self.tree.tag_configure('KAZANDI', background='#d5f5e3')
        self.tree.tag_configure('KAYBETTİ', background='#fadbd8')
        self.tree.pack(fill=tk.BOTH, expand=True, padx=8)
        self.tree.bind('<Double-1>', lambda e: self.ac())
        self.ozet = tk.Label(self.p, text='', anchor='w', font=('Arial', 10, 'bold'))
        self.ozet.pack(fill=tk.X, padx=8, pady=4)
        self.doldur()

    def doldur(self):
        self.tree.delete(*self.tree.get_children())
        liste = self.ana.veri.teklifler(self.ara.get().strip())
        for t in liste:
            self.tree.insert('', 'end', iid=t['no'], tags=(t['sonuc'],), values=[
                para(t[k]) if k in ('net', 'genel_toplam') else (t[k] or '') for k, _, _ in self.kol])
        kaz = sum(1 for t in liste if t['sonuc'] == 'KAZANDI')
        kay = sum(1 for t in liste if t['sonuc'] == 'KAYBETTİ')
        oran = '%%%d' % round(kaz * 100 / (kaz + kay)) if kaz + kay else '-'
        self.ozet.config(text='%d teklif | Kazanılan %d (%s TL) | Kaybedilen %d | Kazanma oranı %s' % (
            len(liste), kaz, para(sum(t['net'] or 0 for t in liste if t['sonuc'] == 'KAZANDI')), kay, oran))

    def _secili(self):
        sec = self.tree.selection()
        if not sec:
            messagebox.showinfo('Seçim', 'Önce bir teklif seçin.', parent=self.p)
        return sec[0] if sec else None

    def ac(self, yeni_no=False):
        no = self._secili()
        if not no:
            return
        g = self.ana.veri.teklif_girdisi(no)
        if not g:
            messagebox.showinfo('Ayrıntı yok', 'Bu kayıt Excel geçmişinden aktarıldı; kalem ayrıntısı yok.', parent=self.p)
            return
        if yeni_no:
            g = dict(g, teklif_no=None, tarih=None)
        self.ana.yukle(g)
        if yeni_no:
            self.ana.no_lbl.config(text='(%s kopyası - kaydedilmedi)' % no)
        self.p.destroy()

    def kopyala(self):
        self.ac(yeni_no=True)

    def sonuc(self, deger):
        no = self._secili()
        if not no:
            return
        neden = ''
        if deger == 'KAYBETTİ':
            from tkinter import simpledialog
            neden = simpledialog.askstring('Red nedeni', 'Neden kaybedildi? (fiyat, rakip, erteleme ...)',
                                           parent=self.p) or ''
        self.ana.veri.teklif_sonuc_yaz(no, deger, neden)
        self.doldur()

    def sil(self):
        no = self._secili()
        if no and messagebox.askyesno('Sil', no + ' silinsin mi? (Yedeklerde kalır)', parent=self.p):
            self.ana.veri.yedek_al('teklif_silme')
            self.ana.veri.teklif_sil(no)
            self.doldur()

    def excele(self):
        import pandas as pd
        yol = filedialog.asksaveasfilename(parent=self.p, defaultextension='.xlsx', initialfile='Teklif_Gecmisi.xlsx',
                                           initialdir=self.ana.veri.klasor)
        if yol:
            pd.DataFrame(self.ana.veri.teklifler()).to_excel(yol, index=False)
            dosya_ac(yol)


# ======================================================================
# Veriler / ayarlar
# ======================================================================
class VeriPenceresi:
    TABLOLAR = [
        ('ayar', '  Ayarlar  ', [('ad', 'Ayar', 170, False), ('deger', 'Değer', 330, True),
                                 ('aciklama', 'Açıklama', 420, False)]),
        ('sabit_kalem', '  Sabit Kalemler  ', [('ad', 'Kalem Adı', 300, True), ('fiyat', 'Fiyat (KDV Hr.)', 150, True),
                                               ('aktif', 'Aktif (1/0)', 90, True)]),
        ('kampanya', '  Kampanyalar  ', [('ad', 'Kampanya', 200, True), ('yuzde', 'İndirim %', 90, True),
                                         ('durum', 'Durum (AKTİF/PASİF)', 140, True), ('aciklama', 'Açıklama', 360, True)]),
        ('rakip', '  Rakipler  ', [('firma', 'Firma', 200, True), ('saat', 'Adam/Saat', 100, True),
                                   ('ucret', 'Saatlik Ücret', 130, True)]),
        ('paket', '  Paketler / Adam-Saat  ', [
            ('kod', 'Kod', 170, True), ('ad', 'Ad', 230, True), ('tur', 'Tür', 90, True),
            ('std_fiyat', 'Std. Fiyat', 90, True), ('ent_fiyat', 'Ent. Fiyat', 90, True),
            ('kull_yuzde', 'Kull. %', 60, True), ('ybh_yuzde', 'YBH %', 60, True), ('std_saat', 'Std A/S', 65, True),
            ('ent_saat', 'Ent A/S', 65, True), ('aktif', 'Aktif', 50, True), ('aciklama', 'Açıklama', 300, True),
            ('urunler', 'Ürünler', 200, True)]),
    ]
    SAYI = {'fiyat', 'yuzde', 'saat', 'ucret', 'std_fiyat', 'ent_fiyat', 'kull_yuzde', 'ybh_yuzde', 'std_saat',
            'ent_saat', 'aktif'}

    def __init__(self, ana):
        self.ana = ana
        self.p = tk.Toplevel(ana.pencere)
        self.p.title('Veriler / Ayarlar')
        self.p.geometry('1250x700')
        self.p.transient(ana.pencere)
        ust = tk.Frame(self.p, bg=LACIVERT, pady=6)
        ust.pack(fill=tk.X)
        for metin, komut, renk in [('📥 Teklif Excel\'inden İçe Aktar', lambda: self.excelden_ice_aktar_sor(ana, self),
                                    '#2980b9'),
                                   ('🔄 DIA Fiyat Listesini Güncelle', self.fiyat_guncelle, '#8e44ad'),
                                   ('💾 Yedek Al', self.yedek_al, '#27ae60'),
                                   ('⏪ Yedekten Geri Yükle', self.yedekten_don, '#c0392b'),
                                   ('📁 Veri Klasörü', self.klasor, '#7f8c8d')]:
            _dugme(ust, metin, komut, renk).pack(side=tk.LEFT, padx=4)
        self.bilgi = tk.Label(self.p, text='', anchor='w', fg='#555')
        self.bilgi.pack(fill=tk.X, padx=8)
        self.nb = ttk.Notebook(self.p)
        self.nb.pack(fill=tk.BOTH, expand=True, padx=8, pady=6)
        self.agaclar = {}
        for tablo, baslik, kol in self.TABLOLAR:
            f = tk.Frame(self.nb)
            self.nb.add(f, text=baslik)
            if tablo == 'paket':
                ara_f = tk.Frame(f)
                ara_f.pack(fill=tk.X)
                tk.Label(ara_f, text='Ara:').pack(side=tk.LEFT)
                self.paket_ara = tk.StringVar()
                tk.Entry(ara_f, textvariable=self.paket_ara, width=30).pack(side=tk.LEFT, padx=4)
                self.paket_ara.trace_add('write', lambda *a: self.doldur('paket'))
                tk.Label(ara_f, text='  Adam/saat burada paket koduyla tutulur; DIA fiyat listesi güncellense de korunur.',
                         fg='#555').pack(side=tk.LEFT)
            tree = ttk.Treeview(f, columns=[k for k, _, _, _ in kol], show='headings')
            for k, b, w, _ in kol:
                tree.heading(k, text=b)
                tree.column(k, width=w, anchor='e' if k in self.SAYI else 'w')
            sb = ttk.Scrollbar(f, orient='vertical', command=tree.yview)
            tree.configure(yscrollcommand=sb.set)
            alt = tk.Frame(f)
            alt.pack(side=tk.BOTTOM, fill=tk.X, pady=4)
            sb.pack(side=tk.RIGHT, fill=tk.Y)
            tree.pack(fill=tk.BOTH, expand=True)
            tree.bind('<Double-1>', lambda e, t=tablo: self.duzenle(t))
            if tablo != 'ayar':
                _dugme(alt, '＋ Ekle', lambda t=tablo: self.duzenle(t, yeni=True), '#27ae60').pack(side=tk.LEFT, padx=3)
                _dugme(alt, '🗑 Sil', lambda t=tablo: self.sil(t), '#e74c3c').pack(side=tk.LEFT, padx=3)
            _dugme(alt, '✎ Düzenle', lambda t=tablo: self.duzenle(t), '#f39c12').pack(side=tk.LEFT, padx=3)
            tk.Label(alt, text='Değişiklikler hemen kaydedilir.', fg='#555').pack(side=tk.LEFT, padx=10)
            self.agaclar[tablo] = (tree, kol)
            self.doldur(tablo)
        self._bilgi()

    def _bilgi(self):
        y = self.ana.veri.yedekler()
        self.bilgi.config(text='Veri dosyası: %s   |   Yedek sayısı: %d   |   Son yedek: %s' % (
            self.ana.veri.yol, len(y), y[0][1] if y else 'yok'))

    def satirlar(self, tablo):
        if tablo == 'ayar':
            mevcut = self.ana.veri.ayarlar()
            return [{'ad': ad, 'deger': mevcut.get(ad, v), 'aciklama': ac} for ad, v, ac in AYARLAR]
        return self.ana.veri.tablo_oku(tablo)

    def doldur(self, tablo):
        tree, kol = self.agaclar[tablo]
        tree.delete(*tree.get_children())
        aranan = katla(self.paket_ara.get().strip()) if tablo == 'paket' else ''
        for n, s in enumerate(self.satirlar(tablo)):
            if aranan and aranan not in katla(' '.join(str(s.get(k) or '') for k in ('kod', 'ad', 'aciklama'))):
                continue
            tree.insert('', 'end', iid=str(n), values=[
                ('%g' % s[k] if isinstance(s.get(k), float) else str(s.get(k) if s.get(k) is not None else '')
                 ).replace('\n', ' ') for k, _, _, _ in kol])

    def duzenle(self, tablo, yeni=False):
        tree, kol = self.agaclar[tablo]
        satirlar = self.satirlar(tablo)
        if yeni:
            mevcut, n = {}, None
        else:
            sec = tree.selection()
            if not sec:
                return
            n = int(sec[0])
            mevcut = satirlar[n]
        d = tk.Toplevel(self.p)
        d.title('Düzenle')
        d.transient(self.p)
        d.grab_set()
        alanlar = {}
        for r, (k, b, _, duz) in enumerate(kol):
            tk.Label(d, text=b + ':').grid(row=r, column=0, sticky='nw', padx=6, pady=3)
            deger = mevcut.get(k)
            deger = '%g' % deger if isinstance(deger, float) else ('' if deger is None else str(deger))
            if k in ('deger', 'aciklama', 'urunler') and (len(deger) > 60 or '\n' in deger or k == 'deger'):
                w = tk.Text(d, width=70, height=4 if k != 'deger' or '\n' in deger else 2, font=('Arial', 10))
                w.insert('1.0', deger)
                al = (lambda w=w: w.get('1.0', 'end').rstrip('\n'))
            else:
                v = tk.StringVar(value=deger)
                w = tk.Entry(d, textvariable=v, width=70)
                al = v.get
            if not duz:
                w.configure(state='disabled')
            w.grid(row=r, column=1, padx=6, pady=3, sticky='w')
            alanlar[k] = al

        def kaydet():
            yeni_s = dict(mevcut)
            for k, al in alanlar.items():
                deger = al()
                if k == 'aktif' and not str(deger).strip():
                    deger = '1'      # bos birakilan 'Aktif' kalemi pasif yapmasin
                yeni_s[k] = (sayi_oku(deger, 0) if k in self.SAYI else deger.strip())
            if tablo == 'ayar':
                self.ana.veri.ayar_yaz(yeni_s['ad'], yeni_s['deger'])
            else:
                anahtar = kol[0][0]
                if not yeni_s.get(anahtar):
                    messagebox.showwarning('Eksik', kol[0][1] + ' boş olamaz.', parent=d)
                    return
                liste = list(satirlar)
                if n is None:
                    liste.append(yeni_s)
                else:
                    liste[n] = yeni_s
                self.ana.veri.tablo_yaz(tablo, liste)
            d.destroy()
            self.doldur(tablo)
            self._bilgi()
            self.ana.veri_degisti()
        _dugme(d, '💾 Kaydet', kaydet, '#27ae60').grid(row=len(kol), column=1, sticky='e', padx=6, pady=8)

    def sil(self, tablo):
        tree, _ = self.agaclar[tablo]
        sec = [int(i) for i in tree.selection()]
        if not sec or not messagebox.askyesno('Sil', '%d satır silinsin mi?' % len(sec), parent=self.p):
            return
        liste = [s for n, s in enumerate(self.satirlar(tablo)) if n not in sec]
        self.ana.veri.tablo_yaz(tablo, liste)
        self.doldur(tablo)
        self.ana.veri_degisti()

    @staticmethod
    def excelden_ice_aktar_sor(ana, pencere=None):
        ust = pencere.p if pencere else ana.pencere
        yol = filedialog.askopenfilename(parent=ust, title='Teklif motoru Excel\'i (RACA_DIA_Teklif_Motoru.xlsx)',
                                         filetypes=[('Excel', '*.xlsx *.xlsm')])
        if not yol:
            return
        try:
            rapor = ana.veri.excelden_ice_aktar(yol)
        except Exception as e:
            messagebox.showerror('İçe aktarılamadı', str(e), parent=ust)
            return
        ana.veri_degisti()
        if pencere:
            for t in pencere.agaclar:
                pencere.doldur(t)
            pencere._bilgi()
        messagebox.showinfo('İçe aktarıldı', 'Alınanlar: ' + ', '.join(rapor) + '\n\nÖnceki veri yedeklendi.',
                            parent=ust)

    def fiyat_guncelle(self):
        yol = filedialog.askopenfilename(parent=self.p, title='DIA paket fiyat listesi (Excel)',
                                         filetypes=[('Excel', '*.xlsx *.xlsm')])
        if not yol:
            return
        try:
            r = self.ana.veri.paket_listesi_guncelle(yol)
        except Exception as e:
            messagebox.showerror('Güncellenemedi', str(e), parent=self.p)
            return
        self.doldur('paket')
        self.ana.veri_degisti()
        ornek = lambda x: (': ' + ', '.join(x[:8]) + (' ...' if len(x) > 8 else '')) if x else ''  # noqa: E731
        messagebox.showinfo('Fiyat listesi güncellendi', '%d paket okundu.\nFiyatı değişen: %d%s\nYeni paket: %d%s\n'
                            'Listeden kalkan (pasif yapıldı): %d%s\n\nYeni paketlerin adam/saatini Paketler sekmesinden '
                            'girin. Önceki hal yedeklendi.' % (
                                r['toplam'], len(r['fiyati_degisen']), ornek(r['fiyati_degisen']), len(r['yeni']),
                                ornek(r['yeni']), len(r['listeden_kalkan']), ornek(r['listeden_kalkan'])), parent=self.p)

    def yedek_al(self):
        yol = self.ana.veri.yedek_al('elle')
        self._bilgi()
        messagebox.showinfo('Yedek', ('Yedek alındı:\n' + yol) if yol else 'Veri boş, yedek alınmadı.', parent=self.p)

    def yedekten_don(self):
        yol = filedialog.askopenfilename(parent=self.p, title='Geri yüklenecek yedek',
                                         initialdir=self.ana.veri.yedek_klasoru, filetypes=[('Yedek', '*.db')])
        if not yol or not messagebox.askyesno('Geri yükle', 'Mevcut veri bu yedekle değiştirilecek.\n(Mevcut hal '
                                              'önce ayrıca yedeklenir.)\n\nDevam edilsin mi?', parent=self.p):
            return
        try:
            self.ana.veri.yedekten_don(yol)
        except Exception as e:
            messagebox.showerror('Geri yüklenemedi', str(e), parent=self.p)
            return
        for t in self.agaclar:
            self.doldur(t)
        self._bilgi()
        self.ana.veri_degisti()
        messagebox.showinfo('Geri yüklendi', os.path.basename(yol) + ' geri yüklendi.', parent=self.p)

    def klasor(self):
        cevap = messagebox.askyesnocancel(
            'Veri klasörü', 'Veriler şurada:\n%s\n\nEvet: klasörü aç\nHayır: başka klasör seç (ör. ortak ağ klasörü; '
            'mevcut veri oraya kopyalanır)' % self.ana.veri.klasor, parent=self.p)
        if cevap is True:
            dosya_ac(self.ana.veri.klasor)
        elif cevap is False:
            yeni = filedialog.askdirectory(parent=self.p, title='Yeni veri klasörü')
            if yeni:
                hedef = veri_klasoru_degistir(yeni)
                messagebox.showinfo('Veri klasörü', 'Yeni veri dosyası:\n%s\n\nProgramı kapatıp açın.' % hedef,
                                    parent=self.p)


# ======================================================================
# Yillik bakim
# ======================================================================
class YbhPenceresi:
    def __init__(self, ana):
        self.ana = ana
        self.kalemler = []
        self.sonuc = None
        self.p = tk.Toplevel(ana.pencere)
        self.p.title('Yıllık Bakım (YBH) Hesabı - Mevcut Müşteri')
        self.p.geometry('1000x600')
        self.p.transient(ana.pencere)
        ust = tk.Frame(self.p, pady=6)
        ust.pack(fill=tk.X, padx=8)
        tk.Label(ust, text='Müşteri:').pack(side=tk.LEFT)
        self.musteri = tk.StringVar(value=ana.musteri.get())
        tk.Entry(ust, textvariable=self.musteri, width=40).pack(side=tk.LEFT, padx=4)
        tk.Label(ust, text='Paket türü:').pack(side=tk.LEFT, padx=(10, 2))
        self.tip = tk.StringVar(value='STANDART')
        cb = ttk.Combobox(ust, textvariable=self.tip, values=['STANDART', 'ENTERPRISE'], state='readonly', width=12)
        cb.pack(side=tk.LEFT)
        cb.bind('<<ComboboxSelected>>', lambda e: self.hesapla())
        _dugme(ust, '＋ Paket', lambda: PaketSecici(self.p, ana.paketler, ana.ayar, self._eklendi), '#27ae60'
               ).pack(side=tk.LEFT, padx=6)
        _dugme(ust, '🗑 Sil', self.sil, '#e74c3c').pack(side=tk.LEFT)
        _dugme(ust, '📊 Excel', self.excel, '#16a085').pack(side=tk.LEFT, padx=6)
        _dugme(ust, 'Tekliften Al', self.tekliften, '#7f8c8d').pack(side=tk.LEFT)
        self.tree = ttk.Treeview(self.p, columns=('kod', 'ad', 'kull', 'modul', 'lisans', 'ybh', 'baz'), show='headings')
        for k, b, w in [('kod', 'Kod', 170), ('ad', 'Ad', 250), ('kull', 'Kullanıcı ✎', 90), ('modul', 'Modül', 60),
                        ('lisans', 'Lisans Bedeli', 120), ('ybh', 'YBH %', 60), ('baz', 'Baz Bakım', 120)]:
            self.tree.heading(k, text=b)
            self.tree.column(k, width=w, anchor='w' if k in ('kod', 'ad') else 'e')
        self.tree.pack(fill=tk.BOTH, expand=True, padx=8)
        self.tree.bind('<Double-1>', self.kullanici_duzenle)
        self.ozet = tk.Label(self.p, text='', font=('Consolas', 10), justify='left', anchor='w')
        self.ozet.pack(fill=tk.X, padx=8, pady=6)
        self.hesapla()

    def _eklendi(self, kod):
        self.kalemler.append({'kod': kod, 'kullanici': 1})
        self.hesapla()

    def tekliften(self):
        s = self.ana.sonuc
        if s:
            self.kalemler = [{'kod': x['kod'], 'kullanici': max(x['adet'], 1)} for x in s['satirlar']]
            self.tip.set(s['paket_tipi'])
            self.hesapla()

    def sil(self):
        for i in sorted((int(x) for x in self.tree.selection()), reverse=True):
            del self.kalemler[i]
        self.hesapla()

    def kullanici_duzenle(self, _e=None):
        from tkinter import simpledialog
        sec = self.tree.selection()
        if not sec:
            return
        k = self.kalemler[int(sec[0])]
        yeni = simpledialog.askinteger('Kullanıcı', k['kod'] + ' kullanıcı sayısı:', initialvalue=k['kullanici'],
                                       minvalue=1, parent=self.p)
        if yeni:
            k['kullanici'] = yeni
            self.hesapla()

    def hesapla(self):
        s = self.sonuc = ybh_hesapla(self.tip.get(), self.kalemler, self.ana.paketler, self.ana.ayar)
        self.tree.delete(*self.tree.get_children())
        for i, x in enumerate(s['satirlar']):
            self.tree.insert('', 'end', iid=str(i), values=(x['kod'], x['ad'], x['kullanici'], x['modul'],
                                                            para(x['lisans']), '%g' % x['ybh_yuzde'], para(x['baz'])))
        self.ozet.config(text=(
            'Toplam baz bakım      : %s\nModül katsayısı       : %g  (%d modül)\nKullanıcı katsayısı   : %g  (%d kullanıcı)\n'
            'Toplam katsayı        : %g\nYILLIK BAKIM (KDV Hr.): %s\nKDV                   : %s\nKDV DAHİL TOPLAM      : %s') % (
            para(s['baz']), s['modul_katsayi'], s['modul'], s['kullanici_katsayi'], s['kullanici'], s['katsayi'],
            para(s['yillik']['net']), para(s['yillik']['kdv']), para(s['yillik']['dahil'])))

    def excel(self):
        yol = filedialog.asksaveasfilename(parent=self.p, defaultextension='.xlsx', initialdir=self.ana.veri.klasor,
                                           initialfile='YBH_%s.xlsx' % re.sub(r'[^\w\- ]', '', self.musteri.get())[:40])
        if yol:
            teklif_cikti.ybh_excel(yol, self.musteri.get(), self.tip.get(), self.sonuc, self.ana.ayar)
            dosya_ac(yol)


def main():
    pencere = TeklifPenceresi()
    if getattr(pencere, 'veri', None):
        pencere.pencere.mainloop()


if __name__ == '__main__':
    main()
