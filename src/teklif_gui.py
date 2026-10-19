"""💼 DIA Teklif Motoru penceresi.

Veri temizleme programindan (Teklif butonu) veya tek basina (teklif_main.py -> RACA_Teklif.exe) acilir.
"""
import os
import re
import subprocess
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

import teklif_cikti
import teklif_tasarim
from teklif_hesap import (hesapla, ybh_hesapla, kod_listesi_coz, sayi_coz, soru_etkisi, soru_etki_metni,
                          standart_cevaplar, kesif_ozeti, satis_brifingi, katla as _katla_h)
from teklif_veri import TeklifVeri, AYARLAR, varsayilan_klasor, veri_klasoru_degistir

LACIVERT = '#2c3e50'
_KATLA = str.maketrans('çğıöşüÇĞİÖŞÜ', 'cgiosuCGIOSU')


def katla(s):
    return str(s).replace('İ', 'i').replace('I', 'ı').lower().translate(_KATLA)


def sayi_oku(metin, bos=None):
    """'1.250,50' / '7.800' / '1250.5' / '%15' -> float; bos -> bos"""
    return sayi_coz(metin, bos)


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


def _ipucu(w, metin):
    """Fare uzerine gelince kucuk aciklama penceresi."""
    metin = (metin or '').strip()
    if not metin:
        return
    durum = {}

    def goster(_e):
        if durum.get('p'):
            return
        p = tk.Toplevel(w)
        p.wm_overrideredirect(True)
        p.wm_geometry('+%d+%d' % (w.winfo_rootx() + 20, w.winfo_rooty() + w.winfo_height() + 2))
        tk.Label(p, text=metin, bg='#ffffe0', relief='solid', bd=1, justify='left', wraplength=420,
                 font=('Arial', 9), padx=6, pady=4).pack()
        durum['p'] = p

    def gizle(_e):
        if durum.get('p'):
            durum.pop('p').destroy()
    w.bind('<Enter>', goster, add='+')
    w.bind('<Leave>', gizle, add='+')


def _dugme(ust, metin, komut, renk='#3498db', **kw):
    return tk.Button(ust, text=metin, command=komut, bg=renk, fg='white', font=('Arial', 10, 'bold'),
                     padx=10, pady=4, relief=tk.FLAT, cursor='hand2', **kw)


# ======================================================================
# Ana teklif penceresi
# ======================================================================
SEVIYE_RENK = {'KRİTİK': '#c0392b', 'UYARI': '#d35400', 'ÖNERİ': '#2471a3', 'İYİ': '#1e8449', 'BİLGİ': '#555555'}


class TeklifPenceresi:
    KOLONLAR = [('kod', 'Paket Kodu', 150), ('ad', 'Paket Adı', 230), ('tur', 'Tür', 80), ('adet', 'Adet ✎', 60),
                ('birim', 'Birim Fiyat', 95), ('kull', 'Kull. %', 55), ('toplam', 'Toplam', 105),
                ('kamp', 'DİA Kamp. İnd.', 100), ('ind_yuzde', 'İnd. % ✎', 65), ('ind_tutar', 'İnd. Tutar ✎', 90), ('net', 'Net (KDV Hr.)', 110),
                ('dahil', 'KDV Dahil', 110), ('saat', 'A/S ✎', 55)]
    DUZENLENIR = {'adet': 'adet', 'ind_yuzde': 'ind_yuzde', 'ind_tutar': 'ind_tutar', 'saat': 'saat'}

    def __init__(self, parent=None, veri_yolu=None, musteri_klasoru=None):
        self.kendi_kok = parent is None
        self.musteri_klasoru = musteri_klasoru
        self.pencere = tk.Tk() if self.kendi_kok else tk.Toplevel(parent)
        self.pencere.title('DIA Teklif Motoru')
        self.pencere.geometry('1420x920')
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
        self.revizyon_onceki = None
        self.sonuc = None
        self.hizmet_elle = {}
        self.odeme_secimi = None
        self.cevap_vars = {}
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
        self.kurallar = self.veri.kurallar()

    # ------------------------------------------------------------ arayuz
    def arayuz(self):
        p = self.pencere
        ust = tk.Frame(p, bg=LACIVERT, pady=6)
        ust.pack(fill=tk.X)
        tk.Label(ust, text='💼 DIA TEKLİF MOTORU', font=('Arial', 14, 'bold'), bg=LACIVERT,
                 fg='white').pack(side=tk.LEFT, padx=12)
        self.no_lbl = tk.Label(ust, text='', font=('Arial', 11, 'bold'), bg=LACIVERT, fg='#f1c40f')
        self.no_lbl.pack(side=tk.LEFT, padx=10)
        for metin, komut, renk in [('⚙ Veriler', self.veriler_ac, '#7f8c8d'),
                                   ('🎨 Tasarım', self.tasarim_ac, '#b9770e'),
                                   ('🔧 YBH', self.ybh_ac, '#8e44ad'),
                                   ('📂 Geçmiş', self.gecmis_ac, '#16a085'),
                                   ('📊 Excel', self.excel_kaydet, '#27ae60'), ('📄 PDF', self.pdf_kaydet, '#c0392b'),
                                   ('📝 Revizyon Kaydet', self.revizyon_kaydet, '#5d6d7e'),
                                   ('💾 Kaydet', self.kaydet, '#2980b9'), ('＋ Yeni', self.yeni, '#d35400')]:
            _dugme(ust, metin, komut, renk).pack(side=tk.RIGHT, padx=2)

        # Musteri bilgileri
        mf = tk.LabelFrame(p, text=' Müşteri ', font=('Arial', 10, 'bold'), bg='#f0f0f0', padx=8, pady=4)
        mf.pack(fill=tk.X, padx=10, pady=(6, 2))
        self.musteri = tk.StringVar()
        self.masaustu = tk.StringVar(value='1')
        self.mobil = tk.StringVar(value='0')
        self.tip_secim = tk.StringVar(value='OTOMATİK')
        self.etiket = tk.StringVar(value='')
        self.tasarim_var = tk.StringVar(value='')
        self.mb = {k: tk.StringVar() for k in ('yetkili', 'telefon', 'eposta', 'vkn', 'vergi_dairesi')}

        # Iki satir ayri cercevede: ayni grid kolonlarini paylasip pencereden tasmasinlar
        satirlar = [tk.Frame(mf, bg='#f0f0f0'), tk.Frame(mf, bg='#f0f0f0')]
        for f in satirlar:
            f.pack(fill=tk.X, pady=1)

        def etiketli(satir, metin, w):
            tk.Label(satirlar[satir], text=metin, bg='#f0f0f0').pack(side=tk.LEFT, padx=(8, 2))
            w.pack(side=tk.LEFT)
        f0, f1 = satirlar
        etiketli(0, 'Müşteri:', tk.Entry(f0, textvariable=self.musteri, width=42, font=('Arial', 11)))
        etiketli(0, 'Masaüstü:', tk.Spinbox(f0, from_=0, to=9999, textvariable=self.masaustu, width=6,
                                            font=('Arial', 11)))
        etiketli(0, 'Mobil:', tk.Spinbox(f0, from_=0, to=9999, textvariable=self.mobil, width=6, font=('Arial', 11)))
        etiketli(0, 'Paket tipi:', ttk.Combobox(f0, textvariable=self.tip_secim, width=12, state='readonly',
                                               values=['OTOMATİK', 'STANDART', 'ENTERPRISE']))
        self.tip_lbl = tk.Label(f0, text='', font=('Arial', 11, 'bold'), bg='#f0f0f0', fg='#c0392b')
        self.tip_lbl.pack(side=tk.LEFT, padx=4)
        etiketli(0, 'Etiket:', ttk.Combobox(f0, textvariable=self.etiket, width=10,
                                            values=['', 'Ekonomik', 'Önerilen', 'Kapsamlı']))
        self.tasarim_cb = ttk.Combobox(f0, textvariable=self.tasarim_var, width=14, state='readonly')
        etiketli(0, 'PDF tasarımı:', self.tasarim_cb)
        etiketli(1, 'Yetkili:', tk.Entry(f1, textvariable=self.mb['yetkili'], width=30))
        etiketli(1, 'Telefon:', tk.Entry(f1, textvariable=self.mb['telefon'], width=16))
        etiketli(1, 'E-posta:', tk.Entry(f1, textvariable=self.mb['eposta'], width=26))
        etiketli(1, 'VKN/TCKN:', tk.Entry(f1, textvariable=self.mb['vkn'], width=13))
        etiketli(1, 'Vergi dairesi:', tk.Entry(f1, textvariable=self.mb['vergi_dairesi'], width=16))
        for v in (self.masaustu, self.mobil, self.tip_secim):
            v.trace_add('write', lambda *a: self.hesapla())

        # Sekmeler
        self.nb = ttk.Notebook(p)
        self.nb.pack(fill=tk.BOTH, expand=True, padx=10, pady=4)
        self.sekme_paket = tk.Frame(self.nb, bg='#f0f0f0')
        self.sekme_kesif = tk.Frame(self.nb, bg='#f0f0f0')
        self.sekme_odeme = tk.Frame(self.nb, bg='#f0f0f0')
        self.sekme_bakim = tk.Frame(self.nb, bg='#f0f0f0')
        self.sekme_asistan = tk.Frame(self.nb, bg='#f0f0f0')
        self.nb.add(self.sekme_paket, text='  1) Paketler  ')
        self.nb.add(self.sekme_kesif, text='  2) Keşif & Uyarlama  ')
        self.nb.add(self.sekme_odeme, text='  3) Ödeme / Taksit  ')
        self.nb.add(self.sekme_bakim, text='  4) Bakım Anlaşması  ')
        self.nb.add(self.sekme_asistan, text='  5) Satış Asistanı & Kontroller  ')
        self._sekme_paketler()
        self._sekme_kesif()
        self._sekme_odeme()
        self._sekme_bakim()
        self._sekme_asistan()
        self._ozet_paneli()
        self.durum_lbl = tk.Label(p, text='', bg='#f0f0f0', fg='#555', font=('Arial', 8), anchor='w')
        self.durum_lbl.pack(fill=tk.X, padx=12, pady=(0, 3))
        self._secenekleri_doldur()

    def _sekme_paketler(self):
        f = self.sekme_paket
        btn = tk.Frame(f, bg='#f0f0f0')
        btn.pack(fill=tk.X, pady=3)
        _dugme(btn, '＋ Paket Ekle', self.paket_ekle, '#27ae60').pack(side=tk.LEFT, padx=3)
        _dugme(btn, '🗑 Seçileni Sil', self.kalem_sil, '#e74c3c').pack(side=tk.LEFT, padx=3)
        _dugme(btn, '▲', lambda: self.kalem_tasi(-1), '#95a5a6').pack(side=tk.LEFT, padx=1)
        _dugme(btn, '▼', lambda: self.kalem_tasi(1), '#95a5a6').pack(side=tk.LEFT, padx=1)
        tk.Label(btn, text='   Kod:', bg='#f0f0f0').pack(side=tk.LEFT)
        self.hizli_kod = tk.StringVar()
        kod_giris = tk.Entry(btn, textvariable=self.hizli_kod, width=26, font=('Arial', 10))
        kod_giris.pack(side=tk.LEFT, padx=3)
        kod_giris.bind('<Return>', lambda e: self.kod_ile_ekle())
        _dugme(btn, 'Ekle', self.kod_ile_ekle, '#27ae60').pack(side=tk.LEFT, padx=1)
        _dugme(btn, '📋 Panodan Yapıştır', self.panodan_ekle, '#8e44ad').pack(side=tk.LEFT, padx=3)
        tk.Label(btn, text='  ✎ kolonlara çift tıklayıp değiştirin; boş = otomatik', bg='#f0f0f0',
                 fg='#555').pack(side=tk.LEFT)
        self.aciklama_lbl = tk.Label(f, text='', bg='#f0f0f0', fg='#555', font=('Arial', 9), anchor='w',
                                     justify='left', wraplength=1300)
        self.aciklama_lbl.pack(fill=tk.X, padx=4)
        tf = tk.Frame(f)
        tf.pack(fill=tk.BOTH, expand=True)
        self.tree = ttk.Treeview(tf, columns=[k for k, _, _ in self.KOLONLAR], show='headings', height=8)
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
        self.tree.bind('<Control-v>', lambda e: self.panodan_ekle())
        self.tree.bind('<Control-V>', lambda e: self.panodan_ekle())
        self.tree.bind('<<TreeviewSelect>>', self._secim_aciklamasi)

        sec = tk.Frame(f, bg='#f0f0f0')
        sec.pack(fill=tk.X, pady=(4, 0))
        sf = tk.LabelFrame(sec, text=' Sabit kalemler ', bg='#f0f0f0', padx=4)
        sf.pack(side=tk.LEFT, fill=tk.Y)
        self.sabit_cerceve = tk.Frame(sf, bg='#f0f0f0')
        self.sabit_cerceve.pack(anchor='w')
        self.sabit_vars = {}
        kamp = tk.LabelFrame(sec, text=' Kampanya / maliyet / ücret ', bg='#f0f0f0', padx=4)
        kamp.pack(side=tk.LEFT, fill=tk.Y, padx=6)
        tk.Label(kamp, text='Kampanya:', bg='#f0f0f0').grid(row=0, column=0, sticky='w')
        self.kampanya = tk.StringVar(value='YOK')
        self.kampanya_cb = ttk.Combobox(kamp, textvariable=self.kampanya, width=20, state='readonly')
        self.kampanya_cb.grid(row=0, column=1, padx=3)
        self.kampanya_cb.bind('<<ComboboxSelected>>', self._kampanya_secildi)
        tk.Label(kamp, text='İndirim %:', bg='#f0f0f0').grid(row=0, column=2)
        self.kampanya_yuzde = tk.StringVar(value='0')
        tk.Entry(kamp, textvariable=self.kampanya_yuzde, width=6).grid(row=0, column=3, padx=3)
        tk.Label(kamp, text='Maliyet (DIA\'ya ödenen):', bg='#f0f0f0').grid(row=1, column=0, sticky='w')
        self.maliyet = tk.StringVar(value='0')
        tk.Entry(kamp, textvariable=self.maliyet, width=14).grid(row=1, column=1, sticky='w', padx=3)
        tk.Label(kamp, text='Bu teklifte saat ücreti:', bg='#f0f0f0').grid(row=2, column=0, sticky='w')
        self.saat_ucreti = tk.StringVar(value='')
        tk.Entry(kamp, textvariable=self.saat_ucreti, width=14).grid(row=2, column=1, sticky='w', padx=3)
        tk.Label(kamp, text='(boş = ayardaki ücret)', bg='#f0f0f0', fg='#777').grid(row=2, column=2, columnspan=2,
                                                                                  sticky='w')
        for v in (self.kampanya_yuzde, self.maliyet, self.saat_ucreti):
            v.trace_add('write', lambda *a: self.hesapla())
        nf = tk.LabelFrame(sec, text=' Teklife özel not (PDF\'e yazılır) ', bg='#f0f0f0', padx=4)
        nf.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.notlar = tk.Text(nf, width=40, height=4, font=('Arial', 9))
        self.notlar.pack(fill=tk.BOTH, expand=True)

    def _sekme_kesif(self):
        f = self.sekme_kesif
        sol = tk.LabelFrame(f, text=' Keşif formu (cevaplar uyarlama saatini belirler) ', bg='#f0f0f0', padx=4)
        sol.pack(side=tk.LEFT, fill=tk.Y)
        ust = tk.Frame(sol, bg='#f0f0f0')
        ust.pack(fill=tk.X)
        tk.Label(ust, text='Profil:', bg='#f0f0f0').pack(side=tk.LEFT)
        self.profil_var = tk.StringVar()
        self.profil_cb = ttk.Combobox(ust, textvariable=self.profil_var, width=18, state='readonly')
        self.profil_cb.pack(side=tk.LEFT, padx=3)
        _dugme(ust, 'Uygula', self.profil_uygula, '#2980b9').pack(side=tk.LEFT, padx=1)
        _dugme(ust, '💾 Profil kaydet', self.profil_kaydet, '#8e44ad').pack(side=tk.LEFT, padx=1)
        _dugme(ust, '🗑', self.profil_sil, '#95a5a6').pack(side=tk.LEFT, padx=1)
        ust2 = tk.Frame(sol, bg='#f0f0f0')
        ust2.pack(fill=tk.X)
        _dugme(ust2, '📁 Müşteri verisinden', self.musteri_verisinden, '#16a085').pack(side=tk.LEFT, pady=2)
        _dugme(ust2, '↺ Standart cevaplar', self.standart_cevaplara_don, '#7f8c8d').pack(side=tk.LEFT, padx=2)
        _dugme(ust2, '✎ Soruları Düzenle', lambda: KesifDuzenleyici(self), '#d35400').pack(side=tk.LEFT)
        self.kesif_durum = tk.Label(sol, text='', bg='#f0f0f0', fg='#1b4f72', font=('Arial', 9, 'bold'), anchor='w')
        self.kesif_durum.pack(fill=tk.X)
        kap = tk.Frame(sol, bg='#f0f0f0')
        kap.pack(fill=tk.BOTH, expand=True)
        tuval = tk.Canvas(kap, bg='#f0f0f0', highlightthickness=0, width=520)
        kay = ttk.Scrollbar(kap, orient='vertical', command=tuval.yview)
        self.kesif_cerceve = tk.Frame(tuval, bg='#f0f0f0')
        self.kesif_cerceve.bind('<Configure>', lambda e: tuval.configure(scrollregion=tuval.bbox('all')))
        tuval.create_window((0, 0), window=self.kesif_cerceve, anchor='nw')
        tuval.configure(yscrollcommand=kay.set)
        kay.pack(side=tk.RIGHT, fill=tk.Y)
        tuval.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sy = tk.LabelFrame(sol, text=' Yerinde çalışma / seyahat ', bg='#f0f0f0', padx=4)
        sy.pack(fill=tk.X, pady=(4, 0))
        self.sehir_disi = tk.BooleanVar(value=False)
        self.yerinde_gun = tk.StringVar(value='0')
        self.yerinde_kisi = tk.StringVar(value='1')
        tk.Checkbutton(sy, text='Şehir dışı', variable=self.sehir_disi, bg='#f0f0f0',
                       command=self.hesapla).grid(row=0, column=0)
        tk.Label(sy, text='Yerinde gün:', bg='#f0f0f0').grid(row=0, column=1, padx=(8, 2))
        tk.Entry(sy, textvariable=self.yerinde_gun, width=5).grid(row=0, column=2)
        tk.Label(sy, text='Kişi:', bg='#f0f0f0').grid(row=0, column=3, padx=(8, 2))
        tk.Entry(sy, textvariable=self.yerinde_kisi, width=4).grid(row=0, column=4)
        for v in (self.yerinde_gun, self.yerinde_kisi):
            v.trace_add('write', lambda *a: self.hesapla())
        self.seyahat_lbl = tk.Label(sy, text='', bg='#f0f0f0', fg='#555', font=('Arial', 8), wraplength=440,
                                    justify='left')
        self.seyahat_lbl.grid(row=1, column=0, columnspan=5, sticky='w')

        sag = tk.LabelFrame(f, text=' Uyarlama hizmet dökümü (Saat ✎ çift tık: elle; boş = öneri) ', bg='#f0f0f0',
                            padx=4)
        sag.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(6, 0))
        self.hizmet_tree = ttk.Treeview(sag, columns=('ad', 'saat', 'tutar', 'neden'), show='headings', height=9)
        for k, b, w, a in [('ad', 'Hizmet', 190, 'w'), ('saat', 'Saat ✎', 70, 'e'), ('tutar', 'Tutar', 110, 'e'),
                           ('neden', 'Neye göre (gerekçe)', 520, 'w')]:
            self.hizmet_tree.heading(k, text=b)
            self.hizmet_tree.column(k, width=w, anchor=a)
        self.hizmet_tree.tag_configure('elle', background='#fef9e7')
        self.hizmet_tree.tag_configure('toplam', background='#d6eaf8', font=('Arial', 9, 'bold'))
        self.hizmet_tree.pack(fill=tk.BOTH, expand=True)
        self.hizmet_tree.bind('<Double-1>', self.hizmet_saat_duzenle)
        self.hizmet_tree.bind('<<TreeviewSelect>>', self._hizmet_gerekce)
        self.hizmet_detay = tk.Label(sag, text='', bg='#f0f0f0', fg='#333', font=('Arial', 9), justify='left',
                                     anchor='w', wraplength=880)
        self.hizmet_detay.pack(fill=tk.X, pady=3)

    def _kesif_formu_kur(self):
        eski = {k: v for k, v in self.girdi_cevaplar().items()} if self.cevap_vars else {}
        for w in self.kesif_cerceve.winfo_children():
            w.destroy()
        self.cevap_vars, self.kesif_etki = {}, {}
        hizmet_adlari = {h['kod']: h['ad'] for h in self.kurallar['hizmetler']}
        r, grup = 0, None
        for s in self.kurallar['sorular']:
            if (s.get('grup') or 'Diğer') != grup:
                grup = s.get('grup') or 'Diğer'
                tk.Label(self.kesif_cerceve, text=grup.upper(), bg='#f0f0f0', fg='#1b4f72',
                         font=('Arial', 9, 'bold')).grid(row=r, column=0, columnspan=3, sticky='w', pady=(8, 1))
                r += 1
            lbl = tk.Label(self.kesif_cerceve, text=s['soru'] + ('  ⓘ' if s.get('aciklama') else ''), bg='#f0f0f0',
                           anchor='w', wraplength=270, justify='left')
            lbl.grid(row=r, column=0, sticky='w', pady=2)
            _ipucu(lbl, (s.get('aciklama') or '') + '\n\nEtkisi: ' + soru_etki_metni(s, hizmet_adlari) +
                   ('\nStandart cevap: ' + str(s['varsayilan']) if s.get('varsayilan') not in (None, '') else ''))
            if s['tur'] == 'EVET_HAYIR':
                v = tk.BooleanVar(value=bool(eski.get(s['kod'], False)))
                w = tk.Checkbutton(self.kesif_cerceve, variable=v, bg='#f0f0f0', command=self.hesapla)
            elif s['tur'] == 'SECIM':
                v = tk.StringVar(value=eski.get(s['kod']) or '')
                w = ttk.Combobox(self.kesif_cerceve, textvariable=v, width=14, state='readonly',
                                 values=[''] + [x.strip() for x in (s.get('secenekler') or '').split(';') if x.strip()])
                v.trace_add('write', lambda *a: self.hesapla())
            else:
                v = tk.StringVar(value='' if eski.get(s['kod']) in (None, '') else '%g' % eski[s['kod']])
                w = tk.Entry(self.kesif_cerceve, textvariable=v, width=10)
                v.trace_add('write', lambda *a: self.hesapla())
            w.grid(row=r, column=1, sticky='w', padx=6, pady=2)
            etki = tk.Label(self.kesif_cerceve, text='', bg='#f0f0f0', fg='#1e8449', font=('Arial', 8, 'bold'))
            etki.grid(row=r, column=2, sticky='w')
            self.cevap_vars[s['kod']] = (s['tur'], v)
            self.kesif_etki[s['kod']] = (s, etki)
            r += 1
        if not self.kurallar['sorular']:
            tk.Label(self.kesif_cerceve, text='Keşif sorusu tanımlı değil ("Soruları Düzenle")', bg='#f0f0f0',
                     fg='#888').grid(row=0, column=0)
        self.profil_cb['values'] = list(self.kurallar.get('profiller') or {})

    def _kesif_etkilerini_goster(self, cevaplar):
        for kod, (soru, lbl) in self.kesif_etki.items():
            tetik, sayi, _ = soru_etkisi(soru, cevaplar.get(kod))
            islem = soru.get('islem')
            if not tetik:
                metin = ''
            elif islem == 'EKLE':
                metin = '+%s saat' % ('%g' % sayi).replace('.', ',')
            elif islem in ('CARP', 'CARP_YUZDE'):
                metin = '× %s' % ('%g' % round(sayi, 3)).replace('.', ',')
            elif islem == 'MIKTAR':
                metin = '%s adet' % ('%g' % sayi).replace('.', ',')
            else:
                metin = ''
            lbl.config(text=metin)
        o = kesif_ozeti(self.kurallar['sorular'], cevaplar)
        self.kesif_durum.config(
            text='Form doluluğu: %d / %d (%%%d)%s' % (o['cevaplanan'], o['toplam'], o['yuzde'],
                                                     '   — boş: ' + ', '.join(o['bos'][:3]) + (' …' if len(o['bos']) > 3
                                                                                          else '') if o['bos'] else ''),
            fg='#1e8449' if o['yuzde'] >= 80 else '#d35400')

    def _cevaplari_yaz(self, cevaplar, sadece_verilenler=False):
        self._yukleniyor = True
        try:
            for kod, (tur, v) in self.cevap_vars.items():
                if sadece_verilenler and kod not in cevaplar:
                    continue
                c = cevaplar.get(kod)
                if tur == 'EVET_HAYIR':
                    v.set(bool(c) if not isinstance(c, str) else _katla_h(c) in ('evet', 'e', '1', 'true'))
                else:
                    v.set('' if c in (None, '') else (c if tur == 'SECIM' else '%g' % sayi_oku(c, 0)))
        finally:
            self._yukleniyor = False
        self.hesapla()

    def standart_cevaplara_don(self):
        if messagebox.askyesno('Standart cevaplar', 'Keşif formu standart cevaplara dönsün mü?\n'
                               '(Elle girdiğiniz cevaplar silinir)', parent=self.pencere):
            self._cevaplari_yaz(standart_cevaplar(self.kurallar['sorular']))

    def profil_uygula(self):
        ad = self.profil_var.get()
        profil = (self.kurallar.get('profiller') or {}).get(ad)
        if profil is None:
            return
        cevaplar = standart_cevaplar(self.kurallar['sorular'])
        cevaplar.update(profil)
        self._cevaplari_yaz(cevaplar)
        self.durum_lbl.config(text='Keşif profili uygulandı: ' + ad)

    def profil_kaydet(self):
        ad = simpledialog.askstring('Keşif profili', 'Mevcut cevaplar hangi adla profil olarak kaydedilsin?\n'
                                    '(ör. Perakende - 3 şube)', initialvalue=self.profil_var.get(), parent=self.pencere)
        if not ad or not ad.strip():
            return
        cevaplar = {k: v for k, v in self.girdi_cevaplar().items() if v not in (None, '', False)}
        for kod in ('cari_sayisi', 'stok_sayisi', 'diger_kayit', 'veri_hata_orani'):   # musteriye ozel sayilar
            cevaplar.pop(kod, None)
        self.veri.kesif_profil_kaydet(ad.strip(), cevaplar)
        self.kurallar = self.veri.kurallar()
        self.profil_cb['values'] = list(self.kurallar['profiller'])
        self.profil_var.set(ad.strip())

    def profil_sil(self):
        ad = self.profil_var.get()
        if ad and messagebox.askyesno('Profil sil', '"%s" profili silinsin mi?' % ad, parent=self.pencere):
            self.veri.kesif_profil_sil(ad)
            self.kurallar = self.veri.kurallar()
            self.profil_cb['values'] = list(self.kurallar['profiller'])
            self.profil_var.set('')

    def _sekme_odeme(self):
        f = self.sekme_odeme
        tk.Label(f, text='Müşteriye hangi ödeme seçenekleri sunulacak? "PDF" kolonuna çift tıklayıp açın/kapatın. '
                         'Komisyon ve bize kalan sadece sizin ekranınızda görünür, PDF\'e yazılmaz. Oranlar: Veriler > '
                         'Ödeme Seçenekleri.', bg='#f0f0f0', fg='#555', wraplength=1350, justify='left'
                 ).pack(anchor='w', pady=3)
        kol = [('pdf', 'PDF', 45, 'center'), ('ad', 'Seçenek', 170, 'w'), ('taksit', 'Taksit', 55, 'e'),
               ('taksit_tutari', 'Taksit Tutarı', 115, 'e'), ('dahil', 'Müşteri Öder (KDV D.)', 140, 'e'),
               ('vf', 'İnd. / Vade F. %', 100, 'e'), ('kom', 'Komisyon %', 80, 'e'),
               ('kom_tutar', 'Komisyon TL', 105, 'e'), ('bize', 'Bize Kalan (KDV Hr.)', 135, 'e'),
               ('kar', 'Kâr', 110, 'e'), ('gereken', 'Komisyonu karşılayan vade f.', 170, 'e')]
        self.odeme_tree = ttk.Treeview(f, columns=[k for k, _, _, _ in kol], show='headings', height=11)
        for k, b, w, a in kol:
            self.odeme_tree.heading(k, text=b)
            self.odeme_tree.column(k, width=w, anchor=a)
        self.odeme_tree.tag_configure('kapali', foreground='#999999')
        self.odeme_tree.tag_configure('acik', background='#fadbd8')
        self.odeme_tree.pack(fill=tk.BOTH, expand=True)
        self.odeme_tree.bind('<Double-1>', self.odeme_goster_degistir)
        self.odeme_not = tk.Label(f, text='', bg='#f0f0f0', fg='#c0392b', anchor='w', justify='left')
        self.odeme_not.pack(fill=tk.X, pady=3)

    def _sekme_bakim(self):
        f = self.sekme_bakim
        ust = tk.Frame(f, bg='#f0f0f0')
        ust.pack(fill=tk.X, pady=3)
        tk.Label(ust, text='Teklife eklenecek bakım anlaşması:', bg='#f0f0f0', font=('Arial', 10, 'bold')
                 ).pack(side=tk.LEFT)
        self.bakim_var = tk.StringVar(value='Ekleme')
        self.bakim_cb = ttk.Combobox(ust, textvariable=self.bakim_var, width=14, state='readonly')
        self.bakim_cb.pack(side=tk.LEFT, padx=5)
        self.bakim_cb.bind('<<ComboboxSelected>>', lambda e: self.hesapla())
        self.bakim_sun = tk.BooleanVar(value=True)
        tk.Checkbutton(ust, text='Seçenekleri PDF\'te göster', variable=self.bakim_sun, bg='#f0f0f0',
                       command=self.hesapla).pack(side=tk.LEFT, padx=10)
        _dugme(ust, 'Önerileni seç', self._bakim_onerileni_sec, '#16a085').pack(side=tk.LEFT)
        self.bakim_tree = ttk.Treeview(f, columns=('ad', 'saat', 'yanit', 'yerinde', 'yillik', 'aylik', 'aciklama'),
                                       show='headings', height=4)
        for k, b, w, a in [('ad', 'Paket', 150, 'w'), ('saat', 'Yıllık Saat', 85, 'e'), ('yanit', 'Yanıt', 90, 'w'),
                           ('yerinde', 'Yerinde Gün', 85, 'e'), ('yillik', 'Yıllık (KDV Hr.)', 120, 'e'),
                           ('aylik', 'Aylık', 100, 'e'), ('aciklama', 'Kapsam', 600, 'w')]:
            self.bakim_tree.heading(k, text=b)
            self.bakim_tree.column(k, width=w, anchor=a)
        self.bakim_tree.tag_configure('onerilen', background='#d5f5e3')
        self.bakim_tree.pack(fill=tk.X)
        self.bakim_detay = tk.Label(f, text='', bg='#f0f0f0', fg='#333', font=('Consolas', 9), justify='left',
                                    anchor='nw')
        self.bakim_detay.pack(fill=tk.X, pady=4)
        self.tco_lbl = tk.Label(f, text='', bg='white', font=('Consolas', 9), justify='left', anchor='nw',
                                relief=tk.GROOVE, padx=6, pady=4)
        self.tco_lbl.pack(fill=tk.X)

    def _sekme_asistan(self):
        f = self.sekme_asistan
        sol = tk.Frame(f, bg='#f0f0f0')
        sol.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 6))
        tk.Label(sol, text='Tahmini kazanma ihtimali', bg='#f0f0f0', font=('Arial', 10, 'bold')).pack(anchor='w')
        self.skor_tuval = tk.Canvas(sol, width=260, height=60, bg='#f0f0f0', highlightthickness=0)
        self.skor_tuval.pack(anchor='w')
        self.skor_gerekce = tk.Label(sol, text='', bg='#f0f0f0', fg='#555', font=('Arial', 8), justify='left',
                                     wraplength=260)
        self.skor_gerekce.pack(anchor='w')
        tk.Label(sol, text='Rakip karşılaştırması', bg='#f0f0f0', font=('Arial', 10, 'bold')).pack(anchor='w',
                                                                                                   pady=(10, 0))
        self.rakip_lbl = tk.Label(sol, text='', bg='white', font=('Consolas', 9), justify='left', anchor='nw',
                                  relief=tk.GROOVE, padx=4)
        self.rakip_lbl.pack(anchor='w', fill=tk.X)
        tk.Label(sol, text='Hazırlık kontrol listesi', bg='#f0f0f0', font=('Arial', 10, 'bold')).pack(anchor='w',
                                                                                                      pady=(10, 0))
        self.hazirlik_lbl = tk.Label(sol, text='', bg='#f0f0f0', font=('Arial', 9), justify='left', anchor='w')
        self.hazirlik_lbl.pack(anchor='w')
        sag = tk.Frame(f, bg='#f0f0f0')
        sag.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        dug = tk.Frame(sag, bg='#f0f0f0')
        dug.pack(fill=tk.X)
        _dugme(dug, '📋 Satış brifingini kopyala', self.brifing_kopyala, '#16a085').pack(side=tk.LEFT, pady=2)
        tk.Label(dug, text='  CRM, e-posta veya mesaja yapıştırın (müşteriye değil, ekibe)', bg='#f0f0f0',
                 fg='#666').pack(side=tk.LEFT)
        kay = ttk.Scrollbar(sag, orient='vertical')
        self.asistan_txt = tk.Text(sag, wrap='word', font=('Arial', 10), height=14, bg='white', padx=8, pady=6,
                                   yscrollcommand=kay.set)
        kay.config(command=self.asistan_txt.yview)
        kay.pack(side=tk.RIGHT, fill=tk.Y)
        self.asistan_txt.pack(fill=tk.BOTH, expand=True)
        for sev, renk in SEVIYE_RENK.items():
            self.asistan_txt.tag_configure(sev, foreground=renk, font=('Arial', 10, 'bold'))
        self.asistan_txt.tag_configure('baslik', font=('Arial', 11, 'bold'), foreground=LACIVERT, spacing1=6)

    def _ozet_paneli(self):
        alt = tk.Frame(self.pencere, bg='#f0f0f0')
        alt.pack(fill=tk.X, padx=10)
        oz = tk.LabelFrame(alt, text=' Özet ', font=('Arial', 10, 'bold'), bg='white', padx=6, pady=2)
        oz.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.ozet_lbl = {}
        satirlar = [('lisans', 'Lisans'), ('uyarlama', 'Uyarlama'), ('sabit', 'Sabit + Seyahat'), ('ara', 'ARA TOPLAM'),
                    ('kampanya', 'Kampanya İndirimi'), ('net', 'NET TOPLAM')]
        for c, b in enumerate(['', 'KDV Hariç', 'KDV', 'KDV Dahil']):
            tk.Label(oz, text=b, bg='white', font=('Arial', 8, 'bold'), fg='#555').grid(row=0, column=c, sticky='e',
                                                                                       padx=6)
        for r, (k, b) in enumerate(satirlar, 1):
            kalin = k in ('ara', 'net')
            tk.Label(oz, text=b, bg='white', font=('Arial', 9, 'bold' if kalin else 'normal')).grid(
                row=r, column=0, sticky='w')
            self.ozet_lbl[k] = []
            for c in range(1, 4):
                lbl = tk.Label(oz, text='', bg='white', font=('Arial', 9, 'bold' if kalin else 'normal'))
                lbl.grid(row=r, column=c, sticky='e', padx=6)
                self.ozet_lbl[k].append(lbl)
        sag = tk.Frame(alt, bg='#f0f0f0')
        sag.pack(side=tk.LEFT, fill=tk.BOTH, padx=(6, 0))
        self.genel_lbl = tk.Label(sag, text='', bg=LACIVERT, fg='white', font=('Arial', 14, 'bold'), pady=6, padx=10)
        self.genel_lbl.pack(fill=tk.X)
        self.uyarlama_lbl = tk.Label(sag, text='', bg='#e67e22', fg='white', font=('Arial', 12, 'bold'), pady=4,
                                     padx=10, cursor='hand2')
        self.uyarlama_lbl.pack(fill=tk.X, pady=(3, 0))
        self.uyarlama_lbl.bind('<Button-1>', lambda e: self.nb.select(self.sekme_kesif))
        self.dia_lbl = tk.Label(sag, text='', bg='#eaf2f8', fg='#1b4f72', font=('Arial', 9, 'bold'), justify='left',
                                anchor='w', padx=6, pady=2)
        self.dia_lbl.pack(fill=tk.X, pady=(3, 0))
        self.uyarlama_detay = tk.Label(sag, text='', bg='#f0f0f0', fg='#555', font=('Arial', 8), justify='left',
                                       wraplength=520)
        self.uyarlama_detay.pack(anchor='w')
        self.kar_lbl = tk.Label(sag, text='', bg='#f0f0f0', font=('Arial', 9, 'bold'))
        self.kar_lbl.pack(anchor='w')
        self.kontrol_lbl = tk.Label(sag, text='', bg='#f0f0f0', font=('Arial', 9, 'bold'), cursor='hand2',
                                    justify='left', wraplength=520)
        self.kontrol_lbl.pack(anchor='w')
        self.kontrol_lbl.bind('<Button-1>', lambda e: self.nb.select(self.sekme_asistan))

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
        self.tasarim_cb['values'] = list(self.veri.tasarimlar())
        self.bakim_cb['values'] = ['Ekleme'] + [b['ad'] for b in self.kurallar['bakim']]
        self._kesif_formu_kur()
        self.durum_lbl.config(text='Veri: %s   |   %d paket   |   Son yedek: %s' % (
            self.veri.yol, len(self.paketler),
            self.veri.yedekler()[0][1] if self.veri.yedekler() else 'yok'))

    def _kampanya_secildi(self, _e=None):
        for k in self.kampanya_listesi:
            if k['ad'] == self.kampanya.get():
                self.kampanya_yuzde.set('%g' % k['yuzde'])
        self.hesapla()

    def _bakim_onerileni_sec(self):
        if self.sonuc and self.sonuc['bakim']['onerilen']:
            self.bakim_var.set(self.sonuc['bakim']['onerilen'])
            self.hesapla()

    # ------------------------------------------------------------ hesap
    def girdi_cevaplar(self):
        cevaplar = {}
        for kod, (tur, v) in self.cevap_vars.items():
            deger = v.get()
            if tur == 'EVET_HAYIR':
                cevaplar[kod] = bool(deger)
            elif tur == 'SECIM':
                cevaplar[kod] = deger or None
            else:
                cevaplar[kod] = sayi_oku(deger, None)
        return cevaplar

    def girdi(self):
        tip = self.tip_secim.get().replace('İ', 'I')
        bakim = self.bakim_var.get()
        return {'teklif_no': self.teklif_no, 'tarih': self.tarih, 'musteri': self.musteri.get().strip(),
                'musteri_bilgi': {k: v.get().strip() for k, v in self.mb.items()}, 'etiket': self.etiket.get(),
                'masaustu': int(sayi_oku(self.masaustu.get(), 0)), 'mobil': int(sayi_oku(self.mobil.get(), 0)),
                'paket_tipi': 'OTOMATIK' if tip.startswith('OTOMAT') else tip,
                'kalemler': [dict(k) for k in self.kalemler],
                'cevaplar': self.girdi_cevaplar(), 'hizmet_elle': dict(self.hizmet_elle),
                'saatlik_ucret': sayi_oku(self.saat_ucreti.get(), None),
                'seyahat': {'sehir_disi': self.sehir_disi.get(), 'gun': sayi_oku(self.yerinde_gun.get(), 0),
                            'kisi': sayi_oku(self.yerinde_kisi.get(), 1)},
                'sabitler': [ad for ad, v in self.sabit_vars.items() if v.get()],
                'kampanya': self.kampanya.get(), 'kampanya_yuzde': sayi_oku(self.kampanya_yuzde.get(), 0),
                'maliyet': sayi_oku(self.maliyet.get(), 0), 'notlar': self.notlar.get('1.0', 'end').strip(),
                'odeme_secimi': list(self.odeme_secimi) if self.odeme_secimi is not None else None,
                'bakim': '' if bakim == 'Ekleme' else bakim, 'bakim_sun': self.bakim_sun.get(),
                'tasarim': self.tasarim_var.get(),
                'revizyon_onceki': self.revizyon_onceki}

    def hesapla(self):
        if self._yukleniyor or not hasattr(self, 'kontrol_lbl'):
            return
        g = self.girdi()
        s = self.sonuc = hesapla(g, self.paketler, self.ayar, self.sabit_fiyatlar, self.rakip_listesi, self.kurallar)
        self.tip_lbl.config(text='→ ' + s['paket_tipi'])
        self._kalemleri_goster(s)
        self._hizmetleri_goster(s)
        self._kesif_etkilerini_goster(g['cevaplar'])
        self._odemeyi_goster(s)
        self._bakimi_goster(s)
        self._asistani_goster(s)
        kdv = s['kdv_orani'] / 100

        def yaz(k, net):
            for lbl, deger in zip(self.ozet_lbl[k], [net, net * kdv, net * (1 + kdv)]):
                lbl.config(text=para(deger))
        yaz('lisans', s['lisans']['net'])
        yaz('uyarlama', s['uyarlama']['net'])
        yaz('sabit', s['sabit']['net'] + s['seyahat']['tutar'])
        yaz('ara', s['ara_toplam'])
        yaz('kampanya', -s['kampanya']['tutar'])
        yaz('net', s['net_toplam'])
        u = s['uyarlama']
        gun_saat = float(self.ayar.get('is_gunu_saat') or 0)
        saat_m = ('%g' % u['saat']).replace('.', ',')
        gun_m = (' ≈ %s iş günü' % ('%g' % round(u['saat'] / gun_saat, 1)).replace('.', ',')) if gun_saat else ''
        self.uyarlama_lbl.config(text='UYARLAMA HİZMET BEDELİ:  %s TL + KDV   (%s saat%s)' % (
            para(u['net']), saat_m, gun_m))
        self.uyarlama_detay.config(text='Uyarlama = lisans × %%%g (%s) + %g saat × %s (%s)%s   — döküm: Keşif & '
                                        'Uyarlama sekmesi' % (
            self.ayar['lisans_yuzdesi'], para(u['lisans_payi']), u['saat'], para(u['saat_ucreti']),
            para(u['saat_bedeli']), ('   → ' + u['sinir'] + ' uygulandı') if u['sinir'] else ''))
        self.genel_lbl.config(text='GENEL TOPLAM (KDV Dahil):  %s TL' % para(s['genel_toplam']))
        d = s['dia']
        self.dia_lbl.config(text='DİA ekranıyla karşılaştırın →  Üyelik Giriş: %s   Ek +1 Masaüstü: %s   Ek +1 Mobil: %s'
                                 '   Yıllık Bulut: %s%s' % (
            para(d['uyelik_giris']), para(d['ek_masaustu']), para(d['ek_mobil']), para(d['yillik_bulut']),
            ('   (DİA kampanyası −%s)' % para(d['kampanya_indirimi'])) if d['kampanya_indirimi'] else ''))
        k = s['kar']
        self.kar_lbl.config(text=('Kâr: %s TL (marj %%%g)' % (para(k['brut']), k['marj'])) if k['maliyet'] else
                            'Kâr: maliyet girilmedi', fg='#1e8449' if k['maliyet'] and k['brut'] > 0 else '#555')
        say = {}
        for x in s['kontrol']['liste']:
            say[x['seviye']] = say.get(x['seviye'], 0) + 1
        ilk = next((x['mesaj'] for x in s['kontrol']['liste'] if x['seviye'] == 'KRİTİK'), '')
        self.kontrol_lbl.config(
            text=('⛔ %d kritik, ⚠ %d uyarı — Satış Asistanı sekmesine bakın.%s' % (
                say.get('KRİTİK', 0), say.get('UYARI', 0), ('\n' + ilk) if ilk else '')
                  if say.get('KRİTİK') or say.get('UYARI') else '✔ Kontroller temiz') + '   |   Kazanma ihtimali: %%%d (%s)' % (
                s['asistan']['skor'], s['asistan']['seviye']),
            fg='#c0392b' if say.get('KRİTİK') else '#d35400' if say.get('UYARI') else '#1e8449')

    def _kalemleri_goster(self, s):
        secili = self.tree.selection()
        self.tree.delete(*self.tree.get_children())
        sira = {x['sira']: x for x in s['satirlar']}
        for i, k in enumerate(self.kalemler):
            x = sira.get(i)
            if not x:
                self.tree.insert('', 'end', iid=str(i), values=(k['kod'], '(paket listesinde yok)'), tags=('hata',))
                continue
            elle = any(k.get(a) not in (None, '') for a in ('adet', 'saat')) or \
                any(k.get(a) for a in ('ind_yuzde', 'ind_tutar'))     # elle 0 saat de elle sayilir
            self.tree.insert('', 'end', iid=str(i), tags=('elle',) if elle else (), values=(
                x['kod'], x['ad'], x['tur'], x['adet'], para(x['birim']), '%g' % x['kull_yuzde'], para(x['toplam']),
                '(kapalı)' if k.get('kamp_kapali') else para(x['kamp_indirim']) if x['kamp_indirim'] else '',
                '%g' % x['ind_yuzde'] if x['ind_yuzde'] else '', para(x['ind_tutar']) if x['ind_tutar'] else '',
                para(x['net']), para(x['dahil']), '%g' % x['saat']))
        li = s['lisans']
        self.tree.insert('', 'end', iid='toplam', tags=('toplam',), values=(
            'TOPLAM', '%d kalem' % len(s['satirlar']), '', '', '', '', para(li['toplam']),
            para(li['kamp_indirim']) if li['kamp_indirim'] else '', '',
            para(li['indirim'] - li['kamp_indirim']) if li['indirim'] - li['kamp_indirim'] else '', para(li['net']), para(li['dahil']),
            '%g' % sum(x['saat'] for x in s['satirlar'])))
        for iid in secili:
            if self.tree.exists(iid):
                self.tree.selection_add(iid)

    def _hizmetleri_goster(self, s):
        secili = self.hizmet_tree.selection()
        self.hizmet_tree.delete(*self.hizmet_tree.get_children())
        for k in s['hizmet']['kalemler']:
            self.hizmet_tree.insert('', 'end', iid=k['kod'], tags=('elle',) if k.get('elle') else (), values=(
                k['ad'], '%g' % k['saat'], para(k['tutar']), ' | '.join(k['nedenler'])))
        u = s['uyarlama']
        if u['lisans_payi']:
            self.hizmet_tree.insert('', 'end', iid='_lisans', values=(
                'Lisansa bağlı pay', '', para(u['lisans_payi']), 'Lisans neti × %%%g' % self.ayar['lisans_yuzdesi']))
        self.hizmet_tree.insert('', 'end', iid='_toplam', tags=('toplam',), values=(
            'UYARLAMA', '%g' % u['saat'], para(u['net']), u['sinir'] and (u['sinir'] + ' uygulandı') or ''))
        for iid in secili:
            if self.hizmet_tree.exists(iid):
                self.hizmet_tree.selection_add(iid)
        sy = s['seyahat']
        self.seyahat_lbl.config(text=('Seyahat: %s = %s' % (sy['detay'], para(sy['tutar']))) if sy['tutar'] else
                                sy.get('detay', ''))

    def _hizmet_gerekce(self, _e=None):
        sec = self.hizmet_tree.selection()
        if sec and self.sonuc:
            k = next((x for x in self.sonuc['hizmet']['kalemler'] if x['kod'] == sec[0]), None)
            if k:
                self.hizmet_detay.config(text='%s — %s\nNeye göre: %s' % (k['ad'], k['kapsam'],
                                                                          '; '.join(k['nedenler']) or '-'))

    def hizmet_saat_duzenle(self, olay):
        iid = self.hizmet_tree.identify_row(olay.y)
        kolon = self.hizmet_tree.identify_column(olay.x)
        if not iid or iid.startswith('_') or kolon != '#2':
            return
        x, y, w, h = self.hizmet_tree.bbox(iid, kolon)
        giris = tk.Entry(self.hizmet_tree, justify='right')
        mevcut = self.hizmet_elle.get(iid)
        giris.insert(0, '' if mevcut is None else '%g' % mevcut)
        giris.place(x=x, y=y, width=w, height=h)
        giris.focus_set()

        def bitir(_e=None, kaydet=True):
            if not giris.winfo_exists() or getattr(giris, '_bitti', False):
                return
            giris._bitti = True
            if kaydet:
                deger = sayi_oku(giris.get(), None)
                if deger is None:
                    self.hizmet_elle.pop(iid, None)
                else:
                    self.hizmet_elle[iid] = deger
            giris.destroy()
            self.hesapla()
        giris.bind('<Return>', bitir)
        giris.bind('<FocusOut>', bitir)
        giris.bind('<Escape>', lambda e: bitir(kaydet=False))

    def _odemeyi_goster(self, s):
        self.odeme_tree.delete(*self.odeme_tree.get_children())
        for o in s['odeme']:
            vf = ('-%%%g' % o['indirim']) if o['indirim'] else ('%%%g' % o['vade_farki'] if o['vade_farki'] else '')
            etiket = () if o['musteriye_goster'] else ('kapali',)
            if o['komisyon_acigi'] > 0 and o['musteriye_goster']:
                etiket = ('acik',)
            self.odeme_tree.insert('', 'end', iid=o['ad'], tags=etiket, values=(
                '✔' if o['musteriye_goster'] else '—', o['ad'], o['taksit'], para(o['taksit_tutari']), para(o['dahil']),
                vf, '%g' % o['komisyon'] if o['komisyon'] else '', para(o['komisyon_tutari']) if o['komisyon'] else '',
                para(o['bize_kalan']), para(o['kar']) if o['kar'] is not None else '',
                ('%%%g' % o['gereken_vade_farki']) if o['komisyon'] else ''))
        acik = [o for o in s['odeme'] if o['musteriye_goster'] and o['komisyon_acigi'] > 0]
        self.odeme_not.config(text='\n'.join('⚠ %s: komisyonun %s\'si size kalıyor (vade farkını en az %%%g yapın).' % (
            o['ad'], para(o['komisyon_acigi']), o['gereken_vade_farki']) for o in acik))

    def odeme_goster_degistir(self, olay):
        iid = self.odeme_tree.identify_row(olay.y)
        if not iid or not self.sonuc:
            return
        gosterilen = [o['ad'] for o in self.sonuc['odeme'] if o['musteriye_goster']]
        if iid in gosterilen:
            gosterilen.remove(iid)
        else:
            gosterilen.append(iid)
        self.odeme_secimi = gosterilen
        self.hesapla()

    def _bakimi_goster(self, s):
        b = s['bakim']
        self.bakim_tree.delete(*self.bakim_tree.get_children())
        for p in b['paketler']:
            self.bakim_tree.insert('', 'end', tags=('onerilen',) if p['onerilen'] else (), values=(
                p['ad'] + ('  ★ önerilen' if p['onerilen'] else ''), '%g' % p['saat'], p['yanit'],
                '%g' % p['yerinde_gun'], para(p['yillik']), para(p['aylik']), p['aciklama']))
        satirlar = ['Yıllık destek saati neye göre belirlendi (ayarlar: Veriler > Ayarlar > destek_*):']
        for ad, deger in b['detay']:
            satirlar.append('  %-48s %s' % (ad, ('%+g saat' % deger) if deger is not None else ''))
        satirlar.append('  %-48s %g saat  (paket çarpanı ile: Temel ×0,6 / Standart ×1 / Premium ×1,5 vb.)' % (
            'TOPLAM', b['yillik_saat']))
        satirlar.append('  Saat ücreti %s;  önerilen paket: %s  (eşikler: %s saat)' % (
            para(b['saat_ucreti']), b['onerilen'] or '-', self.ayar.get('destek_esikleri')))
        self.bakim_detay.config(text='\n'.join(satirlar))
        t = s['tco']
        satirlar = ['%-8s %18s %18s %18s %18s' % ('Yıl', 'Lisans+Uyarlama', 'DIA YBH', 'Bakım Anl.', 'Toplam (KDV Hr.)')]
        for y in t['yillar']:
            satirlar.append('%-8s %18s %18s %18s %18s' % ('%d. yıl' % y['yil'], para(y['tek_seferlik']),
                                                         para(y['ybh']), para(y['bakim']), para(y['toplam'])))
        satirlar.append('%-8s %75s' % ('TOPLAM', para(t['toplam']) + '  (KDV dahil ' + para(t['dahil']) + ')'))
        self.tco_lbl.config(text='Toplam sahip olma maliyeti (PDF\'te gösterilir):\n' + '\n'.join(satirlar))

    def _asistani_goster(self, s):
        a = s['asistan']
        c = self.skor_tuval
        c.delete('all')
        renk = '#1e8449' if a['skor'] >= 65 else '#d68910' if a['skor'] >= 40 else '#c0392b'
        c.create_rectangle(2, 10, 252, 34, outline='#999', fill='#e5e7e9')
        c.create_rectangle(2, 10, 2 + 250 * a['skor'] / 100, 34, outline='', fill=renk)
        c.create_text(130, 22, text='%%%d  %s' % (a['skor'], a['seviye']), font=('Arial', 11, 'bold'), fill='black')
        c.create_text(2, 50, text='Rakip, kampanya, geçmiş teklifler ve kontrollere göre tahmindir.', anchor='w',
                      font=('Arial', 7), fill='#666')
        self.skor_gerekce.config(text='\n'.join(a['gerekce']))
        self.hazirlik_lbl.config(text='\n'.join(('✔ ' if tamam else '✖ ') + m for m, tamam in a['hazirlik']))
        r = s['rakip']
        satir = ['%-12s %14s' % ('RACA (biz)', para(r['biz']))]
        satir += ['%-12s %14s %+6.1f%%' % (x['firma'][:12], para(x['toplam']), x['fark_yuzde']) for x in r['satirlar']]
        if r['ortalama']:
            satir.append('%-12s %14s' % ('Ortalama', para(r['ortalama'])))
        self.rakip_lbl.config(text='Uyarlama (%s):\n' % ('aynı %g saat × rakip ücreti' % s['uyarlama']['saat']
                                                         if self.ayar.get('rakip_ayni_saat') else 'tablodaki saat')
                              + '\n'.join(satir) + '\n' + r['oneri'])
        t = self.asistan_txt
        t.config(state='normal')
        t.delete('1.0', 'end')
        t.insert('end', 'Satış asistanı önerileri\n', 'baslik')
        for o in a['oneriler']:
            t.insert('end', '● %s: ' % o['baslik'], o['seviye'])
            t.insert('end', o['mesaj'] + '\n')
        p = a['pazarlik']
        t.insert('end', '\nPazarlık aralığı (KDV hariç)\n', 'baslik')
        t.insert('end', 'Liste (indirimsiz): %s   →   Teklif: %s   →   Hedef: %s   →   Taban: %s\n' % (
            para(p['liste']), para(p['teklif']), para(p['hedef']), para(p['taban']) if p['taban'] else '—'))
        if p['taban']:
            t.insert('end', 'Taban: %s. Tabanın altına inmeyin; hedefte kapatmaya çalışın.\n' % p['taban_nedeni'])
        t.insert('end', '\nTaviz sırası (fiyattan önce)\n', 'baslik')
        for i, x in enumerate(a['taviz'], 1):
            t.insert('end', '%d. %s\n' % (i, x))
        t.insert('end', '\nİtiraz karşılama\n', 'baslik')
        for itiraz, cevap in a['itirazlar']:
            t.insert('end', itiraz + ': ', 'ÖNERİ')
            t.insert('end', cevap + '\n')
        t.insert('end', '\nTakip planı\n', 'baslik')
        for tarih, is_ in a['takip']:
            t.insert('end', '%s  %s\n' % (tarih, is_))
        t.insert('end', '\nKontroller\n', 'baslik')
        if not s['kontrol']['liste']:
            t.insert('end', '✔ Sorun bulunmadı\n', 'İYİ')
        for k in s['kontrol']['liste']:
            t.insert('end', '%s  ' % k['seviye'], k['seviye'])
            t.insert('end', k['mesaj'] + '\n')
        t.config(state='disabled')

    def brifing_kopyala(self):
        if not self.sonuc:
            return
        self.pencere.clipboard_clear()
        self.pencere.clipboard_append(satis_brifingi(self.girdi(), self.sonuc))
        self.durum_lbl.config(text='Satış brifingi panoya kopyalandı')

    # ------------------------------------------------------------ kalemler
    def paket_ekle(self):
        if not self.paketler:
            messagebox.showinfo('Paket yok', 'Önce Veriler > Excel\'den İçe Aktar ile paket listesini yükleyin.',
                                parent=self.pencere)
            return
        PaketSecici(self.pencere, self.paketler, self.ayar, self._paket_eklendi)

    def _paket_eklendi(self, kod, adet=None, hesapla=True):
        self.kalemler.append({'kod': kod, 'adet': adet, 'ind_yuzde': 0, 'ind_tutar': 0, 'saat': None})
        if hesapla:
            self.hesapla()

    def kodlari_ekle(self, metin):
        """Yapistirilan/yazilan kodlari teklife ekler; bulunamayanlari bildirir. Eklenen sayisini dondurur."""
        if not self.paketler:
            messagebox.showinfo('Paket yok', 'Önce Veriler > Excel\'den İçe Aktar ile paket listesini yükleyin.',
                                parent=self.pencere)
            return 0
        bulunan, bulunamayan, pasif = kod_listesi_coz(metin, self.paketler)
        mevcut = {k['kod'] for k in self.kalemler}
        eklenen, zaten = [], []
        for kod, adet in bulunan:
            if kod in mevcut:
                zaten.append(kod)
                continue
            self._paket_eklendi(kod, adet, hesapla=False)
            mevcut.add(kod)
            eklenen.append(kod)
        if eklenen:
            self.hesapla()
        mesaj = []
        if bulunamayan:
            mesaj.append('Bulunamayan kod(lar):\n  ' + ', '.join(bulunamayan[:30]) +
                         (' …' if len(bulunamayan) > 30 else ''))
        if pasif:
            mesaj.append('Pasif paket(ler), eklenmedi:\n  ' + ', '.join(pasif))
        if zaten:
            mesaj.append('Teklifte zaten var, tekrar eklenmedi:\n  ' + ', '.join(zaten))
        if mesaj:
            messagebox.showwarning('Kod ekleme', '%d paket eklendi.\n\n' % len(eklenen) + '\n\n'.join(mesaj),
                                   parent=self.pencere)
        self.durum_lbl.config(text='Kodla eklendi: %d paket' % len(eklenen))
        return len(eklenen)

    def kod_ile_ekle(self):
        metin = self.hizli_kod.get()
        if metin.strip() and self.kodlari_ekle(metin):
            self.hizli_kod.set('')

    def panodan_ekle(self):
        try:
            metin = self.pencere.clipboard_get()
        except tk.TclError:
            metin = ''
        if not metin.strip():
            messagebox.showinfo('Pano boş', 'Önce Excel\'den veya başka bir yerden paket kodlarını kopyalayın.',
                                parent=self.pencere)
            return
        self.kodlari_ekle(metin)

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
            zorunlu = (' | Gerekli modüller: ' + p['zorunlu']) if p.get('zorunlu') else ''
            x = next((x for x in (self.sonuc or {}).get('satirlar', []) if x['sira'] == int(sec[0])), {})
            kamp = (' | DİA kampanyası: %s (−%s TL; kapatmak için kolona çift tıklayın)' % (
                x['kamp_ad'], para(x['kamp_indirim']))) if x.get('kamp_indirim') else ''
            self.aciklama_lbl.config(text=((p.get('aciklama') or '').replace('\n', ' ')[:400] + zorunlu + kamp))

    def hucre_duzenle(self, olay):
        iid = self.tree.identify_row(olay.y)
        kolon = self.tree.identify_column(olay.x)
        if not iid or iid == 'toplam':
            return
        ad = self.KOLONLAR[int(kolon[1:]) - 1][0]
        if ad == 'kamp':        # DIA kampanyasini bu satirda ac / kapat
            kalem = self.kalemler[int(iid)]
            kalem['kamp_kapali'] = not kalem.get('kamp_kapali')
            self.hesapla()
            return
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
            if not giris.winfo_exists() or getattr(giris, '_bitti', False):
                return
            giris._bitti = True
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

    def musteri_verisinden(self):
        """Veri temizleme programinin musteri klasorundeki dosyalardan kayit sayisi ve hata orani."""
        from teklif_veri import musteri_veri_ozeti
        klasor = filedialog.askdirectory(parent=self.pencere, title='Müşterinin veri temizleme klasörü',
                                         initialdir=self.musteri_klasoru or os.path.expanduser('~'))
        if not klasor:
            return
        ozet = musteri_veri_ozeti(klasor)
        if not ozet['dosyalar']:
            messagebox.showinfo('Veri bulunamadı', 'Klasörde cari_listesi.xlsx / cari_dia_hazir.xlsx gibi dosya yok.',
                                parent=self.pencere)
            return
        eslesme = {'cari_sayisi': ozet['cari'], 'stok_sayisi': ozet['stok'], 'diger_kayit': ozet['diger'],
                   'veri_hata_orani': ozet['hata_orani']}
        for kod, deger in eslesme.items():
            if kod in self.cevap_vars and deger:
                self.cevap_vars[kod][1].set('%g' % deger)
        self.hesapla()
        messagebox.showinfo('Müşteri verisi', 'Okunan dosyalar:\n%s\n\nCari: %d  Stok: %d  Diğer: %d\nHata/uyarı '
                            'oranı: %%%g' % ('\n'.join(ozet['dosyalar']), ozet['cari'], ozet['stok'], ozet['diger'],
                                             ozet['hata_orani']), parent=self.pencere)

    # ------------------------------------------------------------ teklif islemleri
    def yeni(self, sor=True):
        if sor and self.kalemler and not messagebox.askyesno('Yeni teklif', 'Ekrandaki teklif temizlensin mi?\n'
                                                             '(Kaydetmediyseniz kaybolur)', parent=self.pencere):
            return
        self.yukle({}, yeni=True)

    def yukle(self, girdi, yeni=False):
        self._yukleniyor = True
        self.teklif_no = girdi.get('teklif_no')
        self.tarih = girdi.get('tarih')
        self.revizyon_onceki = girdi.get('revizyon_onceki')
        self.musteri.set(girdi.get('musteri', ''))
        mb = girdi.get('musteri_bilgi') or {}
        for k, v in self.mb.items():
            v.set(mb.get(k, ''))
        self.etiket.set(girdi.get('etiket', ''))
        tasarimlar = self.veri.tasarimlar()
        self.tasarim_var.set(girdi.get('tasarim') if girdi.get('tasarim') in tasarimlar else
                             self.veri.varsayilan_tasarim_adi())
        self.masaustu.set(str(girdi.get('masaustu', 1 if yeni else 0)))
        self.mobil.set(str(girdi.get('mobil', 0)))
        tip = girdi.get('paket_tipi', 'OTOMATIK')
        self.tip_secim.set('OTOMATİK' if tip == 'OTOMATIK' else tip)
        self.kalemler = [dict(k) for k in girdi.get('kalemler', [])]
        for ad, v in self.sabit_vars.items():
            v.set(ad in girdi.get('sabitler', []))
        self.kampanya.set(girdi.get('kampanya') or 'YOK')
        self.kampanya_yuzde.set('%g' % (girdi.get('kampanya_yuzde') or 0))
        self.maliyet.set('%g' % (girdi.get('maliyet') or 0))
        self.saat_ucreti.set('' if girdi.get('saatlik_ucret') in (None, '') else '%g' % girdi['saatlik_ucret'])
        self.notlar.delete('1.0', 'end')
        self.notlar.insert('1.0', girdi.get('notlar', ''))
        cevaplar = girdi.get('cevaplar') or {}
        if yeni and not cevaplar:
            cevaplar = standart_cevaplar(self.kurallar['sorular'])
        for kod, (tur, v) in self.cevap_vars.items():
            c = cevaplar.get(kod)
            if tur == 'EVET_HAYIR':
                v.set(bool(c))
            else:
                v.set('' if c in (None, '') else (c if tur == 'SECIM' else '%g' % c))
        self.hizmet_elle = dict(girdi.get('hizmet_elle') or {})
        sy = girdi.get('seyahat') or {}
        self.sehir_disi.set(bool(sy.get('sehir_disi')))
        self.yerinde_gun.set('%g' % (sy.get('gun') or 0))
        self.yerinde_kisi.set('%g' % (sy.get('kisi') or 1))
        self.odeme_secimi = girdi.get('odeme_secimi')
        self.bakim_var.set(girdi.get('bakim') or 'Ekleme')
        self.bakim_sun.set(girdi.get('bakim_sun', True))
        self._yukleniyor = False
        self.no_lbl.config(text='(yeni teklif - kaydedilmedi)' if not self.teklif_no else
                           'Teklif No: ' + str(self.teklif_no))
        self.hesapla()

    def kaydet(self, sessiz=False, no=None):
        if not self.musteri.get().strip():
            messagebox.showwarning('Müşteri', 'Müşteri adını girin.', parent=self.pencere)
            return None
        if not self.kalemler:
            messagebox.showwarning('Kalem yok', 'En az bir paket ekleyin.', parent=self.pencere)
            return None
        self.hesapla()
        g = self.girdi()
        if not g['tarih'] or no:
            from datetime import datetime
            g['tarih'] = self.tarih = datetime.now().strftime('%Y-%m-%d')
        self.teklif_no = self.veri.teklif_kaydet(g, self.sonuc, no or self.teklif_no)
        self.no_lbl.config(text='Teklif No: ' + self.teklif_no + '  ✔ kaydedildi')
        if not sessiz:
            messagebox.showinfo('Kaydedildi', self.teklif_no + ' teklif geçmişine kaydedildi.', parent=self.pencere)
        return self.teklif_no

    def revizyon_kaydet(self):
        """Mevcut teklifi bozmadan -R1, -R2 ... numarasiyla yeni surum olarak kaydeder."""
        if not self.teklif_no:
            return self.kaydet()
        onceki = self.teklif_no
        self.revizyon_onceki = onceki
        yeni_no = self.veri.revizyon_no(onceki)
        if self.kaydet(sessiz=True, no=yeni_no):
            eski = {t['no']: t for t in self.veri.teklifler()}.get(onceki)
            fark = (self.sonuc['genel_toplam'] - eski['genel_toplam']) if eski and eski['genel_toplam'] else None
            messagebox.showinfo('Revizyon', '%s kaydedildi (önceki: %s).%s' % (
                yeni_no, onceki, ('\nGenel toplam farkı: %s TL' % para(fark)) if fark is not None else ''),
                parent=self.pencere)

    def _kritik_onay(self):
        """KRITIK kontrol varsa: yonetici sifresi tanimliysa sorar, degilse onay ister."""
        kritik = [k['mesaj'] for k in self.sonuc['kontrol']['liste'] if k['seviye'] == 'KRİTİK']
        if not kritik:
            return True
        metin = 'Teklifte kritik sorunlar var:\n\n• ' + '\n• '.join(kritik[:8])
        if self.ayar.get('yonetici_sifresi'):
            from tkinter import simpledialog
            from teklif_veri import sifre_ozeti
            sifre = simpledialog.askstring('Yönetici onayı', metin + '\n\nDevam etmek için yönetici şifresi:',
                                           show='*', parent=self.pencere)
            if sifre is None:
                return False
            if sifre_ozeti(sifre) != self.ayar['yonetici_sifresi']:
                messagebox.showerror('Yönetici onayı', 'Şifre hatalı.', parent=self.pencere)
                return False
            return True
        return messagebox.askyesno('Kritik uyarı', metin + '\n\nYine de çıktı alınsın mı?', parent=self.pencere)

    def _cikti_yolu(self, uzanti):
        if not self.kaydet(sessiz=True) or not self._kritik_onay():
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
            teklif_cikti.pdf_olustur(yol, self.girdi(), self.sonuc, self.ayar, self.secili_tasarim())
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
            teklif_cikti.excel_olustur(yol, self.girdi(), self.sonuc, self.ayar, self.secili_tasarim())
        except Exception as e:
            messagebox.showerror('Excel oluşturulamadı', str(e), parent=self.pencere)
            return
        dosya_ac(yol)

    def girdiden_sonuc(self, girdi):
        """Gecmisteki bir teklifi guncel verilerle yeniden hesaplar (karsilastirma icin)."""
        return hesapla(girdi, self.paketler, self.ayar, self.sabit_fiyatlar, self.rakip_listesi, self.kurallar)

    # ------------------------------------------------------------ alt pencereler
    def gecmis_ac(self):
        GecmisPenceresi(self)

    def veriler_ac(self):
        VeriPenceresi(self)

    def tasarim_ac(self):
        TasarimPenceresi(self)

    def secili_tasarim(self):
        return self.veri.tasarim(self.tasarim_var.get())

    def ybh_ac(self):
        if not self.paketler:
            messagebox.showinfo('Paket yok', 'Önce paket listesini içe aktarın.', parent=self.pencere)
            return
        YbhPenceresi(self)

    def veri_degisti(self):
        g = self.girdi()
        self.verileri_yenile()
        self._secenekleri_doldur()
        self.yukle(g)

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
        tam = next((k for k in self.tree.get_children() if aranan and katla(k) == aranan), None)
        if tam:
            self.tree.move(tam, '', 0)
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
        for metin, komut, renk in [('📂 Aç', self.ac, '#2980b9'), ('📑 Kopyala', self.kopyala, '#8e44ad'),
                                   ('✔ KAZANDI', lambda: self.sonuc('KAZANDI'), '#27ae60'),
                                   ('✖ KAYBETTİ', lambda: self.sonuc('KAYBETTİ'), '#c0392b'),
                                   ('… BEKLİYOR', lambda: self.sonuc('BEKLİYOR'), '#7f8c8d'),
                                   ('🗑 Sil', self.sil, '#e74c3c'), ('📊 Excel\'e', self.excele, '#16a085')]:
            _dugme(ust, metin, komut, renk).pack(side=tk.LEFT, padx=2)
        ust2 = tk.Frame(self.p, pady=2)
        ust2.pack(fill=tk.X, padx=8)
        for metin, komut, renk in [('⚖ Seçilenleri Karşılaştır (PDF)', self.karsilastir, '#5d6d7e'),
                                   ('⏱ Gerçekleşen Saat Gir', self.gerceklesen, '#d35400'),
                                   ('📈 Tahmin / Gerçek Raporu', self.tahmin_raporu, '#2471a3')]:
            _dugme(ust2, metin, komut, renk).pack(side=tk.LEFT, padx=2)
        tk.Label(ust2, text='  (Karşılaştırma için Ctrl ile 2-4 teklif seçin)', fg='#555').pack(side=tk.LEFT)
        kol = [('no', 'Teklif No', 120), ('tarih', 'Tarih', 85), ('musteri', 'Müşteri', 260), ('paket_tipi', 'Paket', 90),
               ('kullanici', 'Kullanıcı', 140), ('net', 'Net (KDV Hr.)', 110), ('genel_toplam', 'KDV Dahil', 110),
               ('kampanya', 'Kampanya', 100), ('sonuc', 'Sonuç', 85), ('red_nedeni', 'Red Nedeni', 160)]
        self.kol = kol
        self.tree = ttk.Treeview(self.p, columns=[k for k, _, _ in kol], show='headings', selectmode='extended')
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
        if yeni_no:
            g = dict(g, revizyon_onceki=None)
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

    def karsilastir(self):
        secili = list(self.tree.selection())
        if not 2 <= len(secili) <= 4:
            messagebox.showinfo('Karşılaştırma', 'Ctrl ile 2-4 teklif seçin (örn. Ekonomik / Önerilen / Kapsamlı).',
                                parent=self.p)
            return
        teklifler = []
        for no in sorted(secili):
            g = self.ana.veri.teklif_girdisi(no)
            if not g:
                messagebox.showinfo('Ayrıntı yok', no + ' Excel geçmişinden aktarıldı; kalem ayrıntısı yok.', parent=self.p)
                return
            teklifler.append((g, self.ana.girdiden_sonuc(g)))
        yol = filedialog.asksaveasfilename(parent=self.p, defaultextension='.pdf', initialdir=self.ana.veri.klasor,
                                           initialfile='Secenekler_%s.pdf' % re.sub(r'[^\w\- ]', '',
                                                                                     teklifler[0][0].get('musteri', ''))[:40])
        if yol:
            teklif_cikti.karsilastirma_pdf(yol, teklifler, self.ana.ayar, self.ana.secili_tasarim())
            dosya_ac(yol)

    def gerceklesen(self):
        no = self._secili()
        if not no:
            return
        g = self.ana.veri.teklif_girdisi(no) or {}
        tahmin = (g.get('_ozet') or {}).get('hizmet') or {}
        if not tahmin:
            messagebox.showinfo('Tahmin yok', 'Bu teklif hizmet dökümü olmadan kaydedilmiş.', parent=self.p)
            return
        mevcut = self.ana.veri.gerceklesen_oku(no)
        adlar = {h['kod']: h['ad'] for h in self.ana.kurallar['hizmetler']}
        d = tk.Toplevel(self.p)
        d.title('Gerçekleşen saat - ' + no)
        d.transient(self.p)
        d.grab_set()
        tk.Label(d, text='Proje bitince her hizmet için gerçekte harcanan saati girin. Rapor ve kalibrasyon bu '
                         'verilerle yapılır.', wraplength=520, justify='left').grid(row=0, column=0, columnspan=3,
                                                                                    padx=8, pady=6)
        for c, b in enumerate(['Hizmet', 'Tahmin', 'Gerçekleşen']):
            tk.Label(d, text=b, font=('Arial', 9, 'bold')).grid(row=1, column=c, padx=8)
        vars_ = {}
        for r, (kod, saat) in enumerate(tahmin.items(), 2):
            tk.Label(d, text=adlar.get(kod, kod)).grid(row=r, column=0, sticky='w', padx=8)
            tk.Label(d, text='%g' % saat).grid(row=r, column=1)
            v = tk.StringVar(value='' if kod not in mevcut else '%g' % mevcut[kod][1])
            tk.Entry(d, textvariable=v, width=10).grid(row=r, column=2, padx=8, pady=1)
            vars_[kod] = (saat, v)

        def kaydet():
            self.ana.veri.gerceklesen_yaz(no, {k: (t, sayi_oku(v.get(), None)) for k, (t, v) in vars_.items()})
            d.destroy()
        _dugme(d, '💾 Kaydet', kaydet, '#27ae60').grid(row=len(tahmin) + 2, column=2, pady=8)

    def tahmin_raporu(self):
        rapor = self.ana.veri.tahmin_gercek_raporu()
        d = tk.Toplevel(self.p)
        d.title('Tahmin / Gerçek Raporu')
        d.geometry('980x420')
        d.transient(self.p)
        tk.Label(d, text='Gerçekleşen saat girilmiş projelerden. Sapma artıysa o hizmeti eksik tahmin ediyoruz. '
                         '"Önerilen çarpanı uygula" sonraki tekliflerde o hizmetin saatini düzeltir.',
                 wraplength=940, justify='left').pack(anchor='w', padx=8, pady=6)
        tr = ttk.Treeview(d, columns=('ad', 'proje', 'tahmin', 'gercek', 'sapma', 'carpan', 'oneri'), show='headings')
        for k, b, w in [('ad', 'Hizmet', 240), ('proje', 'Proje', 60), ('tahmin', 'Tahmin (saat)', 110),
                        ('gercek', 'Gerçek (saat)', 110), ('sapma', 'Sapma', 80), ('carpan', 'Mevcut çarpan', 110),
                        ('oneri', 'Önerilen çarpan', 120)]:
            tr.heading(k, text=b)
            tr.column(k, width=w, anchor='w' if k == 'ad' else 'e')
        for r in rapor:
            tr.insert('', 'end', iid=r['kalem'], values=(r['ad'], r['proje'], '%g' % r['tahmin'], '%g' % r['gercek'],
                                                         '%+g%%' % r['sapma'] if r['sapma'] is not None else '',
                                                         '%g' % r['carpan'], r['onerilen_carpan'] or ''))
        tr.pack(fill=tk.BOTH, expand=True, padx=8)
        if not rapor:
            tk.Label(d, text='Henüz gerçekleşen saat girilmemiş.', fg='#888').pack()

        def uygula():
            for kalem in tr.selection():
                r = next(x for x in rapor if x['kalem'] == kalem)
                if r['onerilen_carpan'] and messagebox.askyesno(
                        'Kalibrasyon', '%s çarpanı %g → %g yapılsın mı? (%d projeye dayanıyor)' % (
                            r['ad'], r['carpan'], r['onerilen_carpan'], r['proje']), parent=d):
                    self.ana.veri.carpan_yaz(kalem, r['onerilen_carpan'])
            self.ana.veri_degisti()
            d.destroy()
        _dugme(d, 'Seçili satırlara önerilen çarpanı uygula', uygula, '#d35400').pack(pady=6)

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
        ('kampanya', '  Kampanyalar  ', [('ad', 'Kampanya', 180, True), ('yuzde', 'İndirim %', 80, True),
                                         ('durum', 'Durum (AKTİF/PASİF)', 130, True), ('aciklama', 'Açıklama', 300, True),
                                         ('kosul', 'Koşul (satış asistanı)', 300, True)]),
        ('paket_kampanya', '  DİA Paket Kampanyaları  ', [
            ('ad', 'Kampanya adı', 220, True), ('kodlar', 'Paket kodları (virgülle; * joker: M550_*)', 280, True),
            ('yuzde', 'İnd. %', 70, True), ('tutar', 'İnd. Tutar', 90, True),
            ('baslangic', 'Başlangıç (YYYY-AA-GG)', 140, True), ('bitis', 'Bitiş (YYYY-AA-GG)', 130, True),
            ('aktif', 'Aktif', 50, True), ('kaynak', 'Kaynak', 90, False)]),
        ('rakip', '  Rakipler  ', [('firma', 'Firma', 200, True), ('saat', 'Adam/Saat', 100, True),
                                   ('ucret', 'Saatlik Ücret', 130, True)]),
        ('paket', '  Paketler / Adam-Saat  ', [
            ('kod', 'Kod', 170, True), ('ad', 'Ad', 230, True), ('tur', 'Tür', 90, True),
            ('std_fiyat', 'Std. Fiyat', 90, True), ('ent_fiyat', 'Ent. Fiyat', 90, True),
            ('kull_yuzde', 'Kull. %', 60, True), ('ybh_yuzde', 'YBH %', 60, True), ('std_saat', 'Std A/S', 65, True),
            ('ent_saat', 'Ent A/S', 65, True), ('aktif', 'Aktif', 50, True), ('aciklama', 'Açıklama', 300, True),
            ('urunler', 'Ürünler', 200, True), ('zorunlu', 'Gerekli modüller', 200, True)]),
        ('hizmet', '  Hizmet Kalemleri  ', [
            ('kod', 'Kod', 110, True), ('ad', 'Hizmet', 200, True),
            ('hesap', 'Hesap (PAKET_SAAT/KAYIT/KULLANICI/ADET/SABIT/YUZDE)', 300, True), ('p1', 'P1', 60, True),
            ('p2', 'P2', 60, True), ('carpan', 'Kalibrasyon çarpanı', 120, True), ('kapsam', 'Kapsam (PDF)', 320, True),
            ('sira', 'Sıra', 50, True), ('aktif', 'Aktif', 50, True)]),
        ('kesif', '  Keşif Soruları  ', [
            ('kod', 'Kod', 130, True), ('soru', 'Soru', 280, True), ('tur', 'Tür (EVET_HAYIR/SAYI/SECIM)', 170, True),
            ('secenekler', 'Seçenekler (;)', 150, True), ('hedef', 'Hedef hizmet (* = hepsi)', 140, True),
            ('islem', 'İşlem (EKLE/CARP/CARP_YUZDE/MIKTAR/BILGI)', 220, True), ('deger', 'Değer', 160, True),
            ('esik', 'Eşik', 50, True), ('tetik', 'Tetik (EVET/HAYIR)', 110, True), ('sira', 'Sıra', 50, True),
            ('aktif', 'Aktif', 50, True)]),
        ('odeme_secenegi', '  Ödeme Seçenekleri  ', [
            ('ad', 'Seçenek', 190, True), ('tur', 'Tür (PESIN/TAKSIT/KREDI_KARTI)', 190, True),
            ('taksit', 'Taksit', 60, True), ('indirim', 'İndirim %', 80, True), ('vade_farki', 'Vade farkı %', 90, True),
            ('aylik_vade', 'Aylık vade %', 90, True), ('komisyon', 'Banka komisyonu %', 120, True),
            ('musteriye_goster', 'PDF\'te göster (1/0)', 130, True), ('sira', 'Sıra', 50, True)]),
        ('bakim_paketi', '  Bakım Paketleri  ', [
            ('ad', 'Paket', 120, True), ('saat_carpani', 'Saat çarpanı', 100, True),
            ('yanit_suresi', 'Yanıt süresi', 110, True), ('yerinde_gun', 'Yerinde gün/yıl', 110, True),
            ('indirim', 'İndirim %', 80, True), ('aciklama', 'Kapsam', 450, True), ('sira', 'Sıra', 50, True)]),
    ]
    SAYI = {'fiyat', 'yuzde', 'saat', 'ucret', 'std_fiyat', 'ent_fiyat', 'kull_yuzde', 'ybh_yuzde', 'std_saat',
            'ent_saat', 'aktif', 'p1', 'p2', 'carpan', 'sira', 'esik', 'taksit', 'indirim', 'vade_farki', 'aylik_vade',
            'komisyon', 'musteriye_goster', 'saat_carpani', 'yerinde_gun', 'tutar'}

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
            return [{'ad': ad, 'deger': ('(tanımlı - değiştirmek için düzenleyin)' if mevcut.get(ad) else '')
                     if ad == 'yonetici_sifresi' else mevcut.get(ad, v), 'aciklama': ac} for ad, v, ac in AYARLAR]
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
                if k in ('aktif', 'carpan') and not str(deger).strip():
                    deger = '1'      # bos birakilan 'Aktif' kalemi pasif yapmasin
                yeni_s[k] = (sayi_oku(deger, 0) if k in self.SAYI else deger.strip())
            if tablo == 'ayar':
                deger = yeni_s['deger']
                if yeni_s['ad'] == 'yonetici_sifresi':
                    if deger.startswith('(tanımlı'):
                        d.destroy()
                        return
                    from teklif_veri import sifre_ozeti
                    deger = sifre_ozeti(deger)          # sifre acik yazilmaz, ozeti saklanir
                self.ana.veri.ayar_yaz(yeni_s['ad'], deger)
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
# PDF / Excel tasarimi
# ======================================================================
# Kesif sorulari duzenleyici
# ======================================================================
class KesifDuzenleyici:
    """Kesif sorularini anlasilir bir formla duzenler: soru metni, tur, secenekler, standart cevap, grup,
    aciklama ve etkisi (hangi hizmete ne yapar). Kaydedilene kadar degisiklikler bellekte tutulur."""
    TURLER = [('EVET_HAYIR', 'Evet / Hayır'), ('SAYI', 'Sayı'), ('SECIM', 'Seçim listesi')]
    ISLEMLER = [('EKLE', 'Saat ekle'), ('CARP', 'Saati çarp (×)'), ('CARP_YUZDE', 'Girilen % kadar artır'),
                ('MIKTAR', 'Miktar olarak kullan (kayıt / adet)'), ('BILGI', 'Sadece bilgi (hesaba girmez)')]
    DEGER_IPUCU = {
        ('EVET_HAYIR', 'EKLE'): 'Değer: eklenecek saat (ör. 24)',
        ('EVET_HAYIR', 'CARP'): 'Değer: çarpan (ör. 1,15 = %15 fazla)',
        ('SAYI', 'EKLE'): 'Değer: eşiğin üzerindeki her 1 için saat; Eşik: ücretsiz miktar (ör. 1 şube)',
        ('SAYI', 'CARP'): 'Değer: çarpan',
        ('SECIM', 'EKLE'): 'Değer: Seçenek=saat; ... (ör. Logo=8; Netsis=8; Yok=0)',
        ('SAYI', 'MIKTAR'): 'Değer boş: girilen sayı hizmetin kayıt / adet miktarı olur',
        ('SAYI', 'CARP_YUZDE'): 'Değer boş: girilen yüzde kadar saat artar',
    }

    def __init__(self, ana):
        self.ana = ana
        self.sorular = sorted(ana.veri.tablo_oku('kesif'), key=lambda x: (x.get('sira') or 0, x['kod']))
        self.hizmetler = [h for h in ana.veri.tablo_oku('hizmet')]
        self.hizmet_adlari = {h['kod']: h['ad'] for h in self.hizmetler}
        self.p = tk.Toplevel(ana.pencere)
        self.p.title('Keşif Sorularını Düzenle')
        self.p.geometry('1380x700')
        self.p.transient(ana.pencere)
        ust = tk.Frame(self.p, bg=LACIVERT, pady=6)
        ust.pack(fill=tk.X)
        tk.Label(ust, text='✎ KEŞİF SORULARI', font=('Arial', 13, 'bold'), bg=LACIVERT, fg='white').pack(
            side=tk.LEFT, padx=10)
        tk.Label(ust, text='Soruları yeniden adlandırın, standart cevap verin, etkisini seçin. Sıra = formdaki sıra.',
                 bg=LACIVERT, fg='#f1c40f').pack(side=tk.LEFT)

        sol = tk.Frame(self.p, padx=6, pady=6)
        sol.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.tree = ttk.Treeview(sol, columns=('grup', 'soru', 'tur', 'std', 'etki', 'aktif'), show='headings')
        for k, b, w in [('grup', 'Grup', 110), ('soru', 'Soru', 260), ('tur', 'Tür', 80), ('std', 'Standart', 70),
                        ('etki', 'Etkisi', 260), ('aktif', 'Aktif', 45)]:
            self.tree.heading(k, text=b)
            self.tree.column(k, width=w)
        self.tree.tag_configure('pasif', foreground='#999999')
        self.tree.pack(fill=tk.BOTH, expand=True)
        self.tree.bind('<<TreeviewSelect>>', lambda e: self._sec())
        alt = tk.Frame(sol)
        alt.pack(fill=tk.X, pady=4)
        for metin, komut, renk in [('＋ Yeni soru', self.yeni, '#27ae60'), ('⧉ Kopyala', self.kopyala, '#16a085'),
                                   ('▲', lambda: self.tasi(-1), '#95a5a6'), ('▼', lambda: self.tasi(1), '#95a5a6'),
                                   ('🗑 Sil', self.sil, '#e74c3c')]:
            _dugme(alt, metin, komut, renk).pack(side=tk.LEFT, padx=2)

        sag = tk.LabelFrame(self.p, text=' Seçili soru ', padx=8, pady=6)
        sag.pack(side=tk.LEFT, fill=tk.Y, padx=6, pady=6)
        self.v = {k: tk.StringVar() for k in ('soru', 'kod', 'grup', 'tur', 'secenekler', 'varsayilan', 'hedef',
                                               'islem', 'deger', 'esik', 'tetik', 'aciklama')}
        self.aktif = tk.BooleanVar(value=True)
        satir = 0

        def alan(etiket, w, ipucu=''):
            nonlocal satir
            tk.Label(sag, text=etiket).grid(row=satir, column=0, sticky='nw', pady=3)
            w.grid(row=satir, column=1, sticky='w', pady=3)
            if ipucu:
                _ipucu(w, ipucu)
            satir += 1
        alan('Soru metni:', tk.Entry(sag, textvariable=self.v['soru'], width=48))
        alan('Grup:', ttk.Combobox(sag, textvariable=self.v['grup'], width=24,
                                   values=sorted({s.get('grup') or '' for s in self.sorular} - {''})),
             'Formda başlık olarak görünür; yeni grup adı yazabilirsiniz.')
        alan('Cevap türü:', ttk.Combobox(sag, textvariable=self.v['tur'], state='readonly', width=24,
                                         values=[b for _, b in self.TURLER]))
        alan('Seçenekler:', tk.Entry(sag, textvariable=self.v['secenekler'], width=48),
             'Sadece "Seçim listesi" için: seçenekleri ; ile ayırın (ör. Yok;Logo;Netsis)')
        alan('Standart cevap:', tk.Entry(sag, textvariable=self.v['varsayilan'], width=20),
             'Yeni teklif bu cevapla açılır. Evet/Hayır için EVET veya HAYIR; sayı; ya da seçeneklerden biri.')
        alan('Etkilediği hizmet:', ttk.Combobox(sag, textvariable=self.v['hedef'], state='readonly', width=34,
                                                values=['Tüm hizmetler (*)'] + ['%s (%s)' % (h['ad'], h['kod'])
                                                                                 for h in self.hizmetler]))
        alan('Ne yapar:', ttk.Combobox(sag, textvariable=self.v['islem'], state='readonly', width=34,
                                       values=[b for _, b in self.ISLEMLER]))
        alan('Değer:', tk.Entry(sag, textvariable=self.v['deger'], width=48))
        alan('Eşik:', tk.Entry(sag, textvariable=self.v['esik'], width=8),
             'Sayı sorularında bu miktara kadar saat eklenmez (ör. 1 şube ücretsiz).')
        alan('Hangi cevapta:', ttk.Combobox(sag, textvariable=self.v['tetik'], state='readonly', width=8,
                                            values=['EVET', 'HAYIR']), 'Evet/Hayır sorusu hangi cevapta etkili olur')
        alan('Açıklama:', tk.Entry(sag, textvariable=self.v['aciklama'], width=48),
             'Formda soru üzerine gelince görünen yardım metni')
        alan('Aktif:', tk.Checkbutton(sag, variable=self.aktif))
        alan('Kod:', tk.Entry(sag, textvariable=self.v['kod'], width=24, state='readonly'),
             'Sistem içi anahtar; kampanya koşulları bu kodu kullanır, değiştirilemez.')
        self.ipucu_lbl = tk.Label(sag, text='', fg='#555', wraplength=420, justify='left')
        self.ipucu_lbl.grid(row=satir, column=0, columnspan=2, sticky='w')
        self.etki_lbl = tk.Label(sag, text='', fg='#1e8449', font=('Arial', 10, 'bold'), wraplength=420,
                                 justify='left')
        self.etki_lbl.grid(row=satir + 1, column=0, columnspan=2, sticky='w', pady=6)
        dug = tk.Frame(sag)
        dug.grid(row=satir + 2, column=0, columnspan=2, sticky='w', pady=8)
        _dugme(dug, '✔ Listeye uygula', self.uygula, '#2980b9').pack(side=tk.LEFT, padx=2)
        _dugme(dug, '💾 Kaydet ve Kapat', self.kaydet, '#27ae60').pack(side=tk.LEFT, padx=2)
        _dugme(dug, 'Vazgeç', self.p.destroy, '#95a5a6').pack(side=tk.LEFT, padx=2)
        for k in ('tur', 'islem', 'hedef', 'deger', 'esik', 'tetik', 'secenekler'):
            self.v[k].trace_add('write', lambda *a: self._onizle())
        self.secili = None
        self.doldur()
        if self.sorular:
            self.tree.selection_set('0')

    # ------------------------------------------------------------
    def _kod(self, ad, eslesme):
        for kod, b in eslesme:
            if b == ad:
                return kod
        return ad

    def _formdan(self):
        hedef = self.v['hedef'].get()
        m = re.search(r'\(([^()]*)\)$', hedef)
        return {'kod': self.v['kod'].get(), 'soru': self.v['soru'].get().strip(), 'grup': self.v['grup'].get().strip(),
                'tur': self._kod(self.v['tur'].get(), self.TURLER),
                'secenekler': self.v['secenekler'].get().strip(), 'varsayilan': self.v['varsayilan'].get().strip(),
                'hedef': m.group(1) if m else '', 'islem': self._kod(self.v['islem'].get(), self.ISLEMLER),
                'deger': self.v['deger'].get().strip(), 'esik': sayi_oku(self.v['esik'].get(), 0),
                'tetik': self.v['tetik'].get() or 'EVET', 'aciklama': self.v['aciklama'].get().strip(),
                'aktif': 1 if self.aktif.get() else 0}

    def _onizle(self):
        s = self._formdan()
        self.etki_lbl.config(text='Bu soru: ' + soru_etki_metni(s, self.hizmet_adlari))
        self.ipucu_lbl.config(text=self.DEGER_IPUCU.get((s['tur'], s['islem']), ''))

    def doldur(self):
        self.tree.delete(*self.tree.get_children())
        for i, s in enumerate(self.sorular):
            self.tree.insert('', 'end', iid=str(i), tags=() if s.get('aktif', 1) else ('pasif',), values=(
                s.get('grup') or '', s.get('soru') or s['kod'], dict(self.TURLER).get(s.get('tur'), s.get('tur')),
                s.get('varsayilan') or '', soru_etki_metni(s, self.hizmet_adlari), 'Evet' if s.get('aktif', 1)
                else 'Hayır'))

    def _sec(self):
        sec = self.tree.selection()
        if not sec:
            return
        self.secili = int(sec[0])
        s = self.sorular[self.secili]
        for k in ('soru', 'kod', 'grup', 'secenekler', 'varsayilan', 'deger', 'aciklama'):
            self.v[k].set('' if s.get(k) is None else str(s.get(k)))
        self.v['tur'].set(dict(self.TURLER).get(s.get('tur'), ''))
        self.v['islem'].set(dict(self.ISLEMLER).get(s.get('islem'), ''))
        h = s.get('hedef') or ''
        self.v['hedef'].set('Tüm hizmetler (*)' if h == '*' else
                            '%s (%s)' % (self.hizmet_adlari.get(h, h), h) if h else '')
        self.v['esik'].set('%g' % (s.get('esik') or 0))
        self.v['tetik'].set(s.get('tetik') or 'EVET')
        self.aktif.set(bool(s.get('aktif', 1)))
        self._onizle()

    def uygula(self):
        if self.secili is None:
            return
        yeni = self._formdan()
        if not yeni['soru']:
            messagebox.showwarning('Soru', 'Soru metni boş olamaz.', parent=self.p)
            return
        self.sorular[self.secili] = dict(self.sorular[self.secili], **yeni)
        self.doldur()
        self.tree.selection_set(str(self.secili))

    def _yeni_kod(self, metin):
        taban = re.sub(r'[^a-z0-9]+', '_', _katla_h(metin)).strip('_')[:30] or 'soru'
        kod, n = taban, 2
        mevcut = {s['kod'] for s in self.sorular}
        while kod in mevcut:
            kod, n = '%s_%d' % (taban, n), n + 1
        return kod

    def yeni(self, kaynak=None):
        metin = simpledialog.askstring('Yeni soru', 'Soru metni:', parent=self.p,
                                       initialvalue=(kaynak or {}).get('soru', ''))
        if not metin or not metin.strip():
            return
        s = dict(kaynak or {'tur': 'EVET_HAYIR', 'hedef': '*', 'islem': 'BILGI', 'deger': '', 'esik': 0,
                            'tetik': 'EVET', 'secenekler': '', 'varsayilan': '', 'aciklama': '',
                            'grup': (self.sorular[self.secili].get('grup') if self.secili is not None else '')})
        s.update(soru=metin.strip(), kod=self._yeni_kod(metin), aktif=1)
        konum = (self.secili + 1) if self.secili is not None else len(self.sorular)
        self.sorular.insert(konum, s)
        self.doldur()
        self.tree.selection_set(str(konum))

    def kopyala(self):
        if self.secili is not None:
            self.yeni(dict(self.sorular[self.secili]))

    def tasi(self, yon):
        i = self.secili
        if i is None or not 0 <= i + yon < len(self.sorular):
            return
        self.sorular[i], self.sorular[i + yon] = self.sorular[i + yon], self.sorular[i]
        self.doldur()
        self.tree.selection_set(str(i + yon))

    def sil(self):
        if self.secili is None:
            return
        s = self.sorular[self.secili]
        if messagebox.askyesno('Sil', '"%s" silinsin mi?\n(Geçici olarak kapatmak için "Aktif" işaretini '
                                      'kaldırmanız yeterli)' % s.get('soru'), parent=self.p):
            del self.sorular[self.secili]
            self.secili = None
            self.doldur()

    def kaydet(self):
        if self.secili is not None and self._formdan() != {k: self.sorular[self.secili].get(k)
                                                          for k in self._formdan()}:
            self.uygula()
        for n, s in enumerate(self.sorular, 1):
            s['sira'] = n
        self.ana.veri.tablo_yaz('kesif', self.sorular)
        self.p.destroy()
        self.ana.veri_degisti()


# ======================================================================
class TasarimPenceresi:
    """Teklif tasarimlari: sablon, renk, logo, gosterilecek alanlar, metinler, firma / hazirlayan bilgileri.
    'Kaydet ve Varsayılan Yap' ile tasarim sabitlenir; yeni teklifler bu tasarimla acilir."""
    FIRMA_ALANLARI = [('firma_adi', 'Firma adı'), ('firma_adres', 'Adres'), ('firma_telefon', 'Telefon'),
                      ('firma_eposta', 'E-posta'), ('firma_web', 'Web'), ('firma_vergi', 'Vergi dairesi / no'),
                      ('firma_banka', 'Banka / IBAN'), ('hazirlayan_ad', 'Hazırlayan'),
                      ('hazirlayan_unvan', 'Unvan'), ('hazirlayan_telefon', 'Hazırlayan tel.'),
                      ('hazirlayan_eposta', 'Hazırlayan e-posta')]

    def __init__(self, ana):
        self.ana = ana
        self.p = tk.Toplevel(ana.pencere)
        self.p.title('Teklif Tasarımı (PDF / Excel)')
        self.p.geometry('1180x760')
        self.p.transient(ana.pencere)
        ust = tk.Frame(self.p, bg=LACIVERT, pady=6)
        ust.pack(fill=tk.X)
        tk.Label(ust, text='🎨 TEKLİF TASARIMI', font=('Arial', 13, 'bold'), bg=LACIVERT, fg='white').pack(
            side=tk.LEFT, padx=10)
        tk.Label(ust, text='Şablon seçin, alanları açıp kapatın, metinleri yazın → Önizle → Kaydet. '
                           '"Varsayılan Yap" ile sabitlenir.', bg=LACIVERT, fg='#f1c40f').pack(side=tk.LEFT)

        sol = tk.Frame(self.p, padx=6, pady=6)
        sol.pack(side=tk.LEFT, fill=tk.Y)
        tk.Label(sol, text='Kayıtlı tasarımlar (★ varsayılan)', font=('Arial', 9, 'bold')).pack(anchor='w')
        self.liste = tk.Listbox(sol, width=26, height=18, font=('Arial', 10), exportselection=False)
        self.liste.pack(fill=tk.Y, expand=True)
        self.liste.bind('<<ListboxSelect>>', lambda e: self._sec())
        for metin, komut, renk in [('＋ Yeni (kopya)', self.yeni, '#27ae60'), ('🗑 Sil', self.sil, '#e74c3c'),
                                   ('📌 Varsayılan Yap', self.varsayilan_yap, '#8e44ad')]:
            _dugme(sol, metin, komut, renk).pack(fill=tk.X, pady=2)

        sag = tk.Frame(self.p, padx=6, pady=6)
        sag.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        nb = ttk.Notebook(sag)
        nb.pack(fill=tk.BOTH, expand=True)
        f1, f2, f3, f4 = (tk.Frame(nb, padx=10, pady=8) for _ in range(4))
        nb.add(f1, text='  Görünüm  ')
        nb.add(f2, text='  Gösterilecek Alanlar  ')
        nb.add(f3, text='  Metinler  ')
        nb.add(f4, text='  Firma ve Hazırlayan  ')

        # Gorunum
        self.v = {k: tk.StringVar() for k in ('ad', 'sablon', 'ana_renk', 'vurgu_renk', 'yazi_boyutu', 'baslik')}
        self.logo = tk.StringVar(value=ana.ayar.get('logo_yolu', ''))
        satir = 0

        def alan(etiket, w):
            nonlocal satir
            tk.Label(f1, text=etiket).grid(row=satir, column=0, sticky='w', pady=4)
            w.grid(row=satir, column=1, sticky='w', pady=4)
            satir += 1
        alan('Tasarım adı:', tk.Entry(f1, textvariable=self.v['ad'], width=30))
        alan('Şablon:', ttk.Combobox(f1, textvariable=self.v['sablon'], values=teklif_tasarim.SABLONLAR,
                                     state='readonly', width=14))
        for k, etiket in (('ana_renk', 'Ana renk:'), ('vurgu_renk', 'Vurgu rengi:')):
            fr = tk.Frame(f1)
            tk.Entry(fr, textvariable=self.v[k], width=10).pack(side=tk.LEFT)
            ornek = tk.Label(fr, text='      ', relief='solid', bd=1)
            ornek.pack(side=tk.LEFT, padx=4)
            _dugme(fr, 'Seç…', lambda k=k: self._renk_sec(k), '#7f8c8d').pack(side=tk.LEFT)
            self.v[k].trace_add('write', lambda *a, o=ornek, k=k: self._renk_goster(o, k))
            alan(etiket, fr)
        alan('Yazı boyutu:', tk.Spinbox(f1, from_=7, to=12, increment=0.5, textvariable=self.v['yazi_boyutu'],
                                        width=6))
        alan('Teklif başlığı:', tk.Entry(f1, textvariable=self.v['baslik'], width=50))
        fr = tk.Frame(f1)
        tk.Entry(fr, textvariable=self.logo, width=50).pack(side=tk.LEFT)
        _dugme(fr, 'Seç…', self._logo_sec, '#7f8c8d').pack(side=tk.LEFT, padx=3)
        _dugme(fr, 'Kaldır', lambda: self.logo.set(''), '#95a5a6').pack(side=tk.LEFT)
        alan('Logo (tüm tasarımlar):', fr)
        tk.Label(f1, text='Kurumsal: logo solda, firma bilgileri sağda, ince çizgiler.\n'
                          'Modern: renkli bant başlık, vurgulu bölüm başlıkları.\n'
                          'Sade: siyah-beyaz, az renk, yazıcı dostu.', fg='#555', justify='left').grid(
            row=satir, column=0, columnspan=2, sticky='w', pady=10)

        # Alanlar
        self.goster = {}
        for i, (k, etiket, _) in enumerate(teklif_tasarim.ALANLAR):
            v = tk.BooleanVar()
            tk.Checkbutton(f2, text=etiket, variable=v).grid(row=i % 12, column=i // 12, sticky='w', padx=8, pady=1)
            self.goster[k] = v

        # Metinler
        self.metin = {}
        for i, (k, etiket, yuk) in enumerate([('on_yazi', 'Ön yazı', 5), ('sartlar', 'Ticari şartlar (her satır '
                                                                                    'bir madde)', 6),
                                              ('kapanis', 'Kapanış', 2)]):
            tk.Label(f3, text=etiket + ':', font=('Arial', 9, 'bold')).pack(anchor='w', pady=(6 if i else 0, 0))
            w = tk.Text(f3, height=yuk, width=110, font=('Arial', 10), wrap='word')
            w.pack(fill=tk.X)
            self.metin[k] = w
        tk.Label(f3, text='Kullanılabilen alanlar: {musteri} {yetkili} {tarih} {gecerlilik} {teklif_no} {firma} '
                          '{hazirlayan}', fg='#555').pack(anchor='w', pady=6)

        # Firma
        self.firma = {}
        for i, (k, etiket) in enumerate(self.FIRMA_ALANLARI):
            tk.Label(f4, text=etiket + ':').grid(row=i, column=0, sticky='w', pady=3)
            v = tk.StringVar(value=str(ana.ayar.get(k, '') or '').replace('\n', ' | '))
            tk.Entry(f4, textvariable=v, width=80).grid(row=i, column=1, sticky='w', pady=3)
            self.firma[k] = v
        tk.Label(f4, text='Bu bilgiler tüm tasarımlarda ortaktır (Ayarlar\'a kaydedilir). Banka satırlarını '
                          '" | " ile ayırabilirsiniz.', fg='#555').grid(row=len(self.FIRMA_ALANLARI), column=0,
                                                                      columnspan=2, sticky='w', pady=8)

        alt = tk.Frame(sag)
        alt.pack(fill=tk.X, pady=6)
        for metin, komut, renk in [('👁 Önizle (PDF)', lambda: self.onizle('.pdf'), '#c0392b'),
                                   ('👁 Önizle (Excel)', lambda: self.onizle('.xlsx'), '#27ae60'),
                                   ('💾 Kaydet', self.kaydet, '#2980b9'),
                                   ('📌 Kaydet ve Varsayılan Yap', lambda: self.kaydet(varsayilan=True), '#8e44ad')]:
            _dugme(alt, metin, komut, renk).pack(side=tk.LEFT, padx=3)
        self.durum = tk.Label(alt, text='', fg='#1e8449')
        self.durum.pack(side=tk.LEFT, padx=8)
        self.secili = None
        self._liste_doldur(ana.tasarim_var.get() or ana.veri.varsayilan_tasarim_adi())

    # ------------------------------------------------------------
    def _renk_goster(self, ornek, k):
        r = self.v[k].get()
        if re.fullmatch(r'#[0-9a-fA-F]{6}', r):
            ornek.config(bg=r)

    def _renk_sec(self, k):
        from tkinter import colorchooser
        r = colorchooser.askcolor(color=self.v[k].get() or None, parent=self.p)[1]
        if r:
            self.v[k].set(r.upper())

    def _logo_sec(self):
        yol = filedialog.askopenfilename(parent=self.p, title='Logo seçin',
                                         filetypes=[('Resim', '*.png *.jpg *.jpeg'), ('Tümü', '*.*')])
        if yol:
            self.logo.set(yol)

    def _liste_doldur(self, sec=None):
        self.tumu = self.ana.veri.tasarimlar()
        vars_ = self.ana.veri.varsayilan_tasarim_adi()
        self.adlar = list(self.tumu)
        self.liste.delete(0, 'end')
        for ad in self.adlar:
            self.liste.insert('end', ('★ ' if ad == vars_ else '   ') + ad)
        i = self.adlar.index(sec) if sec in self.adlar else 0
        self.liste.selection_clear(0, 'end')
        self.liste.selection_set(i)
        self._sec()

    def _sec(self):
        sec = self.liste.curselection()
        if not sec:
            return
        self.secili = self.adlar[sec[0]]
        self._forma_yaz(self.secili, self.tumu[self.secili])

    def _forma_yaz(self, ad, t):
        self.v['ad'].set(ad)
        for k in ('sablon', 'ana_renk', 'vurgu_renk', 'baslik'):
            self.v[k].set(t[k])
        self.v['yazi_boyutu'].set('%g' % t['yazi_boyutu'])
        for k, v in self.goster.items():
            v.set(bool(t['goster'].get(k)))
        for k, w in self.metin.items():
            w.delete('1.0', 'end')
            w.insert('1.0', t.get(k) or '')

    def formdan(self):
        return teklif_tasarim.tamamla({
            'sablon': self.v['sablon'].get(), 'ana_renk': self.v['ana_renk'].get().strip(),
            'vurgu_renk': self.v['vurgu_renk'].get().strip(), 'baslik': self.v['baslik'].get().strip(),
            'yazi_boyutu': sayi_oku(self.v['yazi_boyutu'].get(), 9),
            'goster': {k: v.get() for k, v in self.goster.items()},
            **{k: w.get('1.0', 'end').strip() for k, w in self.metin.items()}})

    def _firma_kaydet(self):
        for k, v in self.firma.items():
            deger = v.get().strip()
            if k == 'firma_banka':
                deger = deger.replace(' | ', '\n')
            if str(self.ana.ayar.get(k, '') or '') != deger:
                self.ana.veri.ayar_yaz(k, deger)
        if (self.ana.ayar.get('logo_yolu') or '') != self.logo.get().strip():
            self.ana.veri.ayar_yaz('logo_yolu', self.logo.get().strip())
        self.ana.ayar = self.ana.veri.ayarlar()

    def kaydet(self, varsayilan=False):
        ad = self.v['ad'].get().strip()
        if not ad:
            messagebox.showwarning('Tasarım', 'Tasarım adı boş olamaz.', parent=self.p)
            return
        self._firma_kaydet()
        if self.secili and ad != self.secili and self.secili in self.tumu and messagebox.askyesno(
                'Tasarım', '"%s" adı "%s" olarak değiştirilsin mi?\n(Hayır: yeni tasarım olarak kaydedilir)' % (
                    self.secili, ad), parent=self.p):
            eski_vars = self.ana.veri.varsayilan_tasarim_adi() == self.secili
            self.ana.veri.tasarim_sil(self.secili)
            varsayilan = varsayilan or eski_vars
        self.ana.veri.tasarim_kaydet(ad, self.formdan(), varsayilan=varsayilan)
        self._liste_doldur(ad)
        self.ana.tasarim_cb['values'] = list(self.tumu)
        if varsayilan or self.ana.tasarim_var.get() not in self.tumu:
            self.ana.tasarim_var.set(ad)
        self.durum.config(text='Kaydedildi: %s%s' % (ad, ' (varsayılan)' if varsayilan else ''))

    def yeni(self):
        ad = simpledialog.askstring('Yeni tasarım', 'Yeni tasarımın adı:', parent=self.p)
        if not ad or not ad.strip():
            return
        self.ana.veri.tasarim_kaydet(ad.strip(), self.formdan())
        self._liste_doldur(ad.strip())
        self.ana.tasarim_cb['values'] = list(self.tumu)

    def sil(self):
        if not self.secili or len(self.tumu) <= 1:
            return
        if messagebox.askyesno('Sil', '"%s" tasarımı silinsin mi?' % self.secili, parent=self.p):
            self.ana.veri.tasarim_sil(self.secili)
            self._liste_doldur()
            self.ana.tasarim_cb['values'] = list(self.tumu)
            if self.ana.tasarim_var.get() not in self.tumu:
                self.ana.tasarim_var.set(self.ana.veri.varsayilan_tasarim_adi())

    def varsayilan_yap(self):
        if self.secili:
            self.ana.veri.tasarim_kaydet(self.secili, self.tumu[self.secili], varsayilan=True)
            self._liste_doldur(self.secili)
            self.durum.config(text='Varsayılan: ' + self.secili)

    def onizle(self, uzanti):
        self._firma_kaydet()
        a = self.ana
        g = a.girdi()
        if not g['kalemler']:       # ekranda teklif yoksa ornek: ilk 3 paket
            g = dict(g, musteri=g['musteri'] or 'Örnek Müşteri A.Ş.', teklif_no=g['teklif_no'] or 'ÖNİZLEME',
                     kalemler=[{'kod': k} for k, p in list(a.paketler.items()) if p.get('aktif', 1)][:3])
        s = hesapla(g, a.paketler, a.ayar, a.sabit_fiyatlar, a.rakip_listesi, a.kurallar)
        klasor = os.path.join(a.veri.klasor, 'Teklifler')
        os.makedirs(klasor, exist_ok=True)
        yol = os.path.join(klasor, '_onizleme_%s%s' % (re.sub(r'\W', '_', self.v['ad'].get())[:30], uzanti))
        try:
            (teklif_cikti.pdf_olustur if uzanti == '.pdf' else teklif_cikti.excel_olustur)(
                yol, g, s, a.ayar, self.formdan())
        except PermissionError:
            messagebox.showerror('Önizleme', 'Önizleme dosyası açık; kapatıp tekrar deneyin:\n' + yol, parent=self.p)
            return
        except Exception as e:
            messagebox.showerror('Önizleme', str(e), parent=self.p)
            return
        dosya_ac(yol)


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
