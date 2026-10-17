"""
DIA Kurulum Veri Temizleme - GUI v2+v3
Musteri klasoru secimi + Ayarlar penceresi
"""
import os
import sys
import shutil
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from datetime import datetime
import pandas as pd


class _CanliLog:
    """print() ciktisini satir satir GUI log'una aktarir (EXE modunda canli takip)."""
    def __init__(self, log):
        self._log = log
        self._kalan = ''

    def write(self, metin):
        self._kalan += metin
        while '\n' in self._kalan:
            satir, self._kalan = self._kalan.split('\n', 1)
            if satir.strip():
                self._log(satir)
        return len(metin)

    def flush(self):
        if self._kalan.strip():
            self._log(self._kalan)
        self._kalan = ''


# ============================================================
# ANA PENCERE
# ============================================================
class DiaGUI:
    def __init__(self, root):
        self.root = root
        self.root.title('DIA Kurulum Veri Temizleme')
        self.root.geometry('1000x800')
        self.root.configure(bg='#f0f0f0')

        if getattr(sys, 'frozen', False):
            self.program_klasoru = os.path.dirname(sys.executable)
        else:
            self.program_klasoru = os.path.dirname(os.path.abspath(__file__))
        self.musteri_klasoru = self.program_klasoru

        self.arayuz_olustur()
        self.baslangic_log()

    def arayuz_olustur(self):
        # Baslik
        baslik = tk.Frame(self.root, bg='#2c3e50', pady=12)
        baslik.pack(fill=tk.X)
        tk.Label(
            baslik, text='DIA KURULUM VERI TEMIZLEME',
            font=('Arial', 16, 'bold'), bg='#2c3e50', fg='white'
        ).pack()

        # Musteri klasoru
        ust = tk.Frame(self.root, bg='#f0f0f0', pady=8)
        ust.pack(fill=tk.X, padx=15)

        tk.Label(ust, text='Musteri Klasoru:', font=('Arial', 10, 'bold'),
                 bg='#f0f0f0').pack(side=tk.LEFT)

        self.klasor_var = tk.StringVar(value=self.musteri_klasoru)
        tk.Label(ust, textvariable=self.klasor_var, font=('Arial', 9),
                 bg='white', fg='#333', padx=8, pady=4,
                 anchor='w', relief=tk.SUNKEN, bd=1, width=70
                 ).pack(side=tk.LEFT, padx=8, fill=tk.X, expand=True)

        tk.Button(ust, text='Klasor Sec', command=self.klasor_sec,
                  bg='#3498db', fg='white', font=('Arial', 10, 'bold'),
                  padx=12).pack(side=tk.LEFT)

        # Modul secimi
        modul_frame = tk.LabelFrame(
            self.root, text=' Modul Secimi ', font=('Arial', 10, 'bold'),
            bg='#f0f0f0', padx=10, pady=8
        )
        modul_frame.pack(fill=tk.X, padx=15, pady=8)

        self.modul_var = tk.StringVar(value='cari')
        moduller = [
            ('Cari', 'cari'), ('Stok', 'stok'), ('Banka', 'banka'),
            ('Personel', 'personel'), ('Demirbas', 'demirbas'),
            ('Hesap Plani', 'hesapplan'), ('Muhasebe Kod', 'muhkod'),
            ('Urun Agaci', 'urunagaci'), ('HEPSI', 'hepsi'),
        ]

        for i, (isim, kod) in enumerate(moduller):
            tk.Radiobutton(
                modul_frame, text=isim, variable=self.modul_var, value=kod,
                font=('Arial', 10), bg='#f0f0f0', padx=8
            ).grid(row=i // 5, column=i % 5, sticky='w', padx=5, pady=3)

        # Islem modu
        islem_frame = tk.LabelFrame(
            self.root, text=' Islem Modu ', font=('Arial', 10, 'bold'),
            bg='#f0f0f0', padx=10, pady=8
        )
        islem_frame.pack(fill=tk.X, padx=15, pady=5)

        self.mod_var = tk.StringVar(value='incele')
        modlar = [
            ('Inceleme (silmez, INCELEME uretir)', 'incele'),
            ("Uygula (INCELEME'deki E'leri siler)", 'uygula'),
            ('Otomatik Sil', 'otomatik'),
            ('Kuru Calistir (dosya yazmaz)', 'dryrun'),
        ]
        for isim, kod in modlar:
            tk.Radiobutton(
                islem_frame, text=isim, variable=self.mod_var, value=kod,
                font=('Arial', 10), bg='#f0f0f0'
            ).pack(anchor='w', padx=10)

        # Butonlar
        btn_frame = tk.Frame(self.root, bg='#f0f0f0', pady=10)
        btn_frame.pack(fill=tk.X, padx=15)

        tk.Button(btn_frame, text='▶ CALISTIR', command=self.calistir,
                  bg='#27ae60', fg='white', font=('Arial', 12, 'bold'),
                  padx=25, pady=10).pack(side=tk.LEFT, padx=3)

        tk.Button(btn_frame, text='⚙ Ayarlar', command=self.ayarlar_ac,
                  bg='#9b59b6', fg='white', font=('Arial', 11, 'bold'),
                  padx=15, pady=10).pack(side=tk.LEFT, padx=3)

        tk.Button(btn_frame, text='📚 Eşleme Öğren', command=self.esleme_ogren_ac,
                  bg='#16a085', fg='white', font=('Arial', 11, 'bold'),
                  padx=15, pady=10).pack(side=tk.LEFT, padx=3)

        tk.Button(btn_frame, text='ℹ Hakkında', command=self.hakkinda_ac,
                  bg='#34495e', fg='white', font=('Arial', 11, 'bold'),
                  padx=15, pady=10).pack(side=tk.LEFT, padx=3)

        tk.Button(btn_frame, text='🗑 Gecici Sil', command=self.gecici_sil,
                  bg='#e67e22', fg='white', font=('Arial', 11, 'bold'),
                  padx=15, pady=10).pack(side=tk.LEFT, padx=3)

        tk.Button(btn_frame, text='📁 Klasoru Ac', command=self.klasor_ac,
                  bg='#3498db', fg='white', font=('Arial', 11, 'bold'),
                  padx=15, pady=10).pack(side=tk.LEFT, padx=3)

        tk.Button(btn_frame, text='Temizle', command=self.cikti_temizle,
                  bg='#95a5a6', fg='white', font=('Arial', 10),
                  padx=15, pady=10).pack(side=tk.RIGHT, padx=3)

        # Entegrasyon: Logo -> temizle -> DIA
        ent_frame = tk.Frame(self.root, bg='#f0f0f0')
        ent_frame.pack(fill=tk.X, padx=15)
        tk.Label(ent_frame, text='Entegrasyon:', font=('Arial', 10, 'bold'),
                 bg='#f0f0f0').pack(side=tk.LEFT, padx=(3, 8))
        tk.Button(ent_frame, text="🔌 Logo'dan Çek", command=self.logo_ac,
                  bg='#c0392b', fg='white', font=('Arial', 11, 'bold'),
                  padx=15, pady=6).pack(side=tk.LEFT, padx=3)
        tk.Button(ent_frame, text="☁ DIA'ya Gönder", command=self.dia_ac,
                  bg='#2980b9', fg='white', font=('Arial', 11, 'bold'),
                  padx=15, pady=6).pack(side=tk.LEFT, padx=3)

        # Cikti
        cikti_frame = tk.LabelFrame(
            self.root, text=' Cikti ', font=('Arial', 10, 'bold'),
            bg='#f0f0f0', padx=5, pady=5
        )
        cikti_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=8)

        self.cikti = scrolledtext.ScrolledText(
            cikti_frame, font=('Consolas', 9),
            bg='#1e1e1e', fg='#d4d4d4', insertbackground='white'
        )
        self.cikti.pack(fill=tk.BOTH, expand=True)

        # Durum cubugu
        self.durum_var = tk.StringVar(value='Hazir')
        tk.Label(self.root, textvariable=self.durum_var,
                 bg='#34495e', fg='white', anchor='w',
                 padx=10, font=('Arial', 9)).pack(fill=tk.X, side=tk.BOTTOM)

    def baslangic_log(self):
        self.log('DIA Temizleme v2+v3 hazir.')
        self.log("Musteri klasorunu secin ve modul secip CALISTIR'a basin.")
        self.log("Ayarlar butonuyla kurallar.xlsx'i GUI'den duzenleyebilirsiniz.")
        self.log('')

    def log(self, mesaj):
        saat = datetime.now().strftime('%H:%M:%S')
        self.cikti.insert(tk.END, '[' + saat + '] ' + mesaj + '\n')
        self.cikti.see(tk.END)
        self.root.update()

    def klasor_sec(self):
        secilen = filedialog.askdirectory(
            title='Musteri klasorunu secin',
            initialdir=self.program_klasoru
        )
        if secilen:
            self.musteri_klasoru = secilen
            self.klasor_var.set(secilen)
            self.log('Klasor: ' + secilen)

    def klasor_ac(self):
        os.startfile(self.musteri_klasoru)

    def cikti_temizle(self):
        self.cikti.delete('1.0', tk.END)

    def calistir(self):
        modul = self.modul_var.get()
        mod = self.mod_var.get()

        # kurallar.xlsx musteri klasorunde yoksa program klasorunden kopyala
        k_kaynak = os.path.join(self.program_klasoru, 'kurallar.xlsx')
        k_hedef = os.path.join(self.musteri_klasoru, 'kurallar.xlsx')
        if os.path.exists(k_kaynak) and not os.path.exists(k_hedef):
            try:
                shutil.copy2(k_kaynak, k_hedef)
            except:
                pass

        # EXE icinde calisiyorsa dia_isle'yi dogrudan cagir
        if getattr(sys, 'frozen', False):
            try:
                from contextlib import redirect_stdout
                import dia_isle as _di
                _di.KLASOR = self.musteri_klasoru
                _di_argv = [modul]
                if mod == 'uygula':
                    _di_argv.append('--uygula')
                elif mod == 'otomatik':
                    _di_argv.append('--sil')
                elif mod == 'dryrun':
                    _di_argv.append('--dry-run')
                self.log('============================================================')
                self.log('MODUL: ' + modul.upper() + '  |  MOD: ' + mod)
                self.log('KLASOR: ' + self.musteri_klasoru)
                self.log('============================================================')
                self.durum_var.set('Calisiyor...')
                _buf = _CanliLog(self.log)
                _eski_argv = sys.argv.copy()
                sys.argv = ['dia_isle.py'] + _di_argv
                try:
                    with redirect_stdout(_buf):
                        _di.main()
                finally:
                    sys.argv = _eski_argv
                    _buf.flush()
                self.durum_var.set('Tamamlandi')
            except Exception as e:
                self.log('HATA: ' + str(e))
                import traceback
                self.log(traceback.format_exc())
                self.durum_var.set('Hata')
            return

        cmd = [sys.executable,
               os.path.join(self.program_klasoru, 'dia_isle.py'),
               modul]

        if mod == 'uygula':
            cmd.append('--uygula')
        elif mod == 'otomatik':
            cmd.append('--sil')
        elif mod == 'dryrun':
            cmd.append('--dry-run')

        self.log('============================================================')
        self.log('MODUL: ' + modul.upper() + '  |  MOD: ' + mod)
        self.log('KLASOR: ' + self.musteri_klasoru)
        self.log('============================================================')
        self.durum_var.set('Calisiyor...')

        try:
            sonuc = subprocess.run(
                cmd, cwd=self.musteri_klasoru,
                capture_output=True, text=True,
                encoding='utf-8', errors='replace'
            )
            if sonuc.stdout:
                for satir in sonuc.stdout.split('\n'):
                    if satir.strip():
                        self.log(satir)
            if sonuc.stderr:
                self.log('HATA:')
                for satir in sonuc.stderr.split('\n'):
                    if satir.strip():
                        self.log(satir)
            self.durum_var.set('Tamamlandi')
        except Exception as e:
            self.log('HATA: ' + str(e))
            self.durum_var.set('Hata')

    def gecici_sil(self):
        silinecek = []
        try:
            for dosya in os.listdir(self.musteri_klasoru):
                if (dosya.endswith('_INCELEME.xlsx') or
                        dosya.endswith('_RAPOR.xlsx') or
                        dosya.endswith('_SILINENLER.xlsx')):
                    silinecek.append(dosya)
        except Exception as e:
            self.log('Klasor okunamadi: ' + str(e))
            return

        if not silinecek:
            messagebox.showinfo('Bilgi', 'Silinecek gecici dosya yok.')
            return

        cevap = messagebox.askyesno(
            'Onay',
            str(len(silinecek)) + ' dosya silinecek:\n\n' +
            '\n'.join(silinecek[:10]) +
            ('\n...' if len(silinecek) > 10 else '')
        )
        if cevap:
            for dosya in silinecek:
                try:
                    os.remove(os.path.join(self.musteri_klasoru, dosya))
                    self.log('Silindi: ' + dosya)
                except Exception as e:
                    self.log('Silinemedi: ' + dosya + ' - ' + str(e))
            self.durum_var.set('Temizlendi')

    def hakkinda_ac(self):
        HakkindaPenceresi(self.root, self.program_klasoru)

    def esleme_ogren_ac(self):
        EslemeOgrenPenceresi(self.root, self.musteri_klasoru, self.program_klasoru)

    def ayarlar_ac(self):
        AyarlarPenceresi(self.root, self.musteri_klasoru, self.program_klasoru)

    def logo_ac(self):
        from entegrasyon_gui import LogoPenceresi
        LogoPenceresi(self.root, lambda: self.musteri_klasoru, self.log)

    def dia_ac(self):
        from entegrasyon_gui import DiaPenceresi
        DiaPenceresi(self.root, lambda: self.musteri_klasoru, self.log)



# ============================================================
# HAKKINDA PENCERESI
# ============================================================
class HakkindaPenceresi:
    def __init__(self, parent, program_klasoru):
        self.pencere = tk.Toplevel(parent)
        self.pencere.title('Hakkında - DIA Kurulum Veri Temizleme')
        self.pencere.geometry('720x800')
        self.pencere.configure(bg='#f0f0f0')
        self.pencere.transient(parent)
        self.pencere.grab_set()

        # Baslik
        baslik = tk.Frame(self.pencere, bg='#2c3e50', pady=20)
        baslik.pack(fill=tk.X)
        tk.Label(baslik, text='DIA KURULUM VERİ TEMİZLEME',
                 font=('Arial', 18, 'bold'),
                 bg='#2c3e50', fg='white').pack()
        tk.Label(baslik, text='Sürüm 2.1  •  Excel → DIA Dönüştürücü',
                 font=('Arial', 10),
                 bg='#2c3e50', fg='#bdc3c7').pack(pady=(5, 0))

        # Icerik
        icerik = tk.Frame(self.pencere, bg='white', padx=20, pady=15)
        icerik.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)

        def bas(t):
            tk.Label(icerik, text=t, font=('Arial', 11, 'bold'),
                     bg='white', fg='#2c3e50', anchor='w').pack(fill=tk.X, pady=(10, 4))

        def madde(t):
            tk.Label(icerik, text='  •  ' + t, font=('Arial', 10),
                     bg='white', fg='#333', anchor='w',
                     justify='left', wraplength=650).pack(fill=tk.X, pady=1)

        bas('Ne İşe Yarar?')
        madde('Müşteri Excel dosyalarını DIA muhasebe programına hazır hale getirir.')
        madde('Mükerrer kayıtları bulur, telefon/vergi/e-posta/IBAN formatlarını düzeltir.')
        madde('Türkçe karakterleri kodlarda temizler, varsayılan alanları doldurur.')
        madde('Renkli inceleme raporu üretir; silinecek satırları siz onaylarsınız.')

        bas('Desteklenen Modüller (8)')
        madde('Cari  •  Stok  •  Banka  •  Personel')
        madde('Demirbaş  •  Hesap Planı  •  Muhasebe Kodu  •  Ürün Ağacı')

        bas('Temel Özellikler')
        madde('KDV geçerli oran kontrolü (0, 1, 8, 10, 18, 20)')
        madde('IBAN MOD-97 doğrulaması')
        madde('VKN (10 hane) / TCKN (11 hane) hane kontrolü')
        madde('Tarih sırası kontrolü (Doğum ≤ İşe Başlama ≤ İşten Ayrılma)')
        madde('Hiyerarşi kontrolü (üst hesap var mı?)')
        madde('Seviye atlama kontrolü (Grup1 boş, Grup2 dolu olmamalı)')
        madde('Çoklu müşteri işleme (--musteri-hepsi)')
        madde('Renkli Excel çıktısı (sarı: mükerrer, kırmızı: hata, turuncu: düzeltme)')
        madde('Otomatik sütun eşleme + öğrenme (📚 Eşleme Öğren)')
        madde('Mükerrer grup raporu (hangi veri neden mükerrer?)')
        madde('İl / ilçe düzeltme: aynı hücreye yazılmış veya yer değiştirmiş il-ilçe (81 il, 973 ilçe)')
        madde("VKN baştaki 0 tamamlama (Excel'in sildiği sıfırlar) + VKN / TCKN algoritma kontrolü")
        madde('Telefon biçimi ayarlanabilir (5321234567 / 05321234567 / +905321234567)')
        madde('Tüm yeni özellikler ⚙ Ayarlar > Ayarlar sekmesinden açılıp kapatılabilir')

        bas('Kullanım')
        madde('1) Müşteri klasörünü seçin')
        madde('2) Modülü işaretleyin (Cari / Stok / ... / HEPSİ)')
        madde('3) İşlem modunu seçin: İncele (önerilir), Uygula, Otomatik Sil, Dry-Run')
        madde('4) ▶ ÇALIŞTIR → INCELEME.xlsx açılır → SIL_ONAY sütununa E yazın')
        madde('5) --uygula ile onayladığınız satırlar silinir')

        bas('Geliştirici')
        tk.Label(icerik, text='Raca Teknoloji A.Ş.',
                 font=('Arial', 12, 'bold'),
                 bg='white', fg='#2c3e50').pack(anchor='w', pady=(5, 0))
        tk.Label(icerik, text='Yusuf KARACA',
                 font=('Arial', 11, 'italic'),
                 bg='white', fg='#7f8c8d').pack(anchor='w')

        # Alt
        alt = tk.Frame(self.pencere, bg='#f0f0f0', pady=10)
        alt.pack(fill=tk.X, padx=15)
        tk.Button(alt, text='Kapat', command=self.pencere.destroy,
                  bg='#95a5a6', fg='white', font=('Arial', 11, 'bold'),
                  padx=30, pady=8).pack(side=tk.RIGHT)

        tk.Label(self.pencere,
                 text='© 2026 Raca Teknoloji A.Ş.  •  Tüm hakları saklıdır.',
                 font=('Arial', 8), bg='#f0f0f0', fg='#7f8c8d').pack(side=tk.BOTTOM, pady=5)


# ESLEME OGREN PENCERESI
class EslemeOgrenPenceresi:
    def __init__(self, parent, musteri_klasoru, program_klasoru):
        self.musteri_klasoru = musteri_klasoru
        self.program_klasoru = program_klasoru

        self.kurallar_yolu = os.path.join(musteri_klasoru, 'kurallar.xlsx')
        if not os.path.exists(self.kurallar_yolu):
            kaynak = os.path.join(program_klasoru, 'kurallar.xlsx')
            if os.path.exists(kaynak):
                try:
                    shutil.copy2(kaynak, self.kurallar_yolu)
                except Exception:
                    pass

        self.pencere = tk.Toplevel(parent)
        self.pencere.title('Sutun Esleme Ogren')
        self.pencere.geometry('1000x650')
        self.pencere.configure(bg='#f0f0f0')
        self.pencere.transient(parent)
        self.pencere.grab_set()

        self.dia_alanlari = self._dia_alanlari_yukle()
        self.arayuz_olustur()
        self.tara_ve_doldur()

    def _dia_alanlari_yukle(self):
        try:
            from _ortak import DIA_ALANLARI
            return sorted(DIA_ALANLARI)
        except Exception as e:
            print('DIA_ALANLARI yuklenemedi: ' + str(e))
            return ['carikartkodu', 'unvan', 'telefon1', 'telefon2', 'ceptel', 'eposta', 'vergidairesi', 'verginumarasi', 'tckimlikno2', 'sehir', 'ilce', 'adres1', 'ulke']

    def arayuz_olustur(self):
        ust = tk.Frame(self.pencere, bg='#2c3e50', pady=10)
        ust.pack(fill=tk.X)
        tk.Label(ust, text='TANINMAYAN SUTUN BASLIKLARI',
                 font=('Arial', 12, 'bold'),
                 bg='#2c3e50', fg='white').pack()

        bilgi = tk.Label(
            self.pencere,
            text="Asagidaki sutunlar taninmadi. DIA alanini secip KAYDET'e basin. (Cift tikla -> sec)",
            font=('Arial', 9, 'italic'), bg='#f0f0f0', fg='#555'
        )
        bilgi.pack(anchor='w', padx=15, pady=5)

        cerceve = tk.Frame(self.pencere, bg='white')
        cerceve.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.tree = ttk.Treeview(cerceve,
                                 columns=('modul', 'excel', 'dia'),
                                 show='headings', height=18)
        self.tree.heading('modul', text='Modul')
        self.tree.heading('excel', text='Excel Basligi')
        self.tree.heading('dia', text='DIA Alani (cift tikla)')
        self.tree.column('modul', width=100, anchor='center')
        self.tree.column('excel', width=400)
        self.tree.column('dia', width=300)

        vsb = ttk.Scrollbar(cerceve, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree.bind('<Double-1>', self.hucre_duzenle)

        alt = tk.Frame(self.pencere, bg='#f0f0f0', pady=10)
        alt.pack(fill=tk.X, padx=10)

        tk.Button(alt, text='KAYDET', command=self.kaydet,
                  bg='#27ae60', fg='white', font=('Arial', 12, 'bold'),
                  padx=25, pady=8).pack(side=tk.RIGHT, padx=5)
        tk.Button(alt, text='Iptal', command=self.pencere.destroy,
                  bg='#95a5a6', fg='white', font=('Arial', 11),
                  padx=20, pady=8).pack(side=tk.RIGHT, padx=5)

    def tara_ve_doldur(self):
        try:
            df_es = pd.read_excel(self.kurallar_yolu, sheet_name='ESLESME')
            bilinen = set()
            for _, r in df_es.iterrows():
                m = str(r['modul']).strip()
                b = str(r['excel_baslik']).strip()
                if m and b:
                    bilinen.add((m, b))
        except Exception as e:
            print('ESLESME okunamadi: ' + str(e))
            bilinen = set()

        try:
            df_ay = pd.read_excel(self.kurallar_yolu, sheet_name='AYARLAR')
            girdiler = {}
            for _, r in df_ay.iterrows():
                if str(r['ayar']).strip() == 'girdi_dosya':
                    girdiler[str(r['modul']).strip()] = str(r['deger']).strip()
        except Exception:
            girdiler = {}

        try:
            from _ortak import DIA_ALANLARI
        except Exception:
            DIA_ALANLARI = set()

        bulunan = 0
        for modul, dosya in girdiler.items():
            tam = os.path.join(self.musteri_klasoru, dosya)
            if not os.path.exists(tam):
                continue
            try:
                with pd.ExcelFile(tam) as xl:
                    sayfa_adlari = xl.sheet_names
                for sayfa in sayfa_adlari:
                    df = pd.read_excel(tam, sheet_name=sayfa, nrows=1)
                    for col in df.columns:
                        col_str = str(col).strip()
                        if not col_str or col_str.startswith('Unnamed'):
                            continue
                        if col_str in DIA_ALANLARI:
                            continue
                        if (modul, col_str) in bilinen:
                            continue
                        self.tree.insert('', 'end',
                                         values=(modul, col_str, ''))
                        bulunan += 1
            except Exception as e:
                print('Hata (' + dosya + '): ' + str(e))

        if bulunan == 0:
            self.tree.insert('', 'end',
                             values=('', '(taninmayan baslik yok)', ''))

    def hucre_duzenle(self, event):
        item = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if not item or col != '#3':
            return
        mevcut = self.tree.item(item, 'values')
        if len(mevcut) < 2:
            return

        bbox = self.tree.bbox(item, column='#3')
        if not bbox:
            return
        x, y, w, h = bbox
        combo = ttk.Combobox(self.pencere, values=self.dia_alanlari,
                             state='readonly', font=('Arial', 9))
        combo.place(x=x, y=y, width=w, height=h)
        if len(mevcut) > 2 and mevcut[2]:
            combo.set(mevcut[2])
        combo.focus()

        def sec(_e=None):
            secim = combo.get()
            self.tree.item(item, values=(mevcut[0], mevcut[1], secim))
            combo.destroy()

        combo.bind('<<ComboboxSelected>>', sec)
        combo.bind('<FocusOut>', lambda _e: combo.destroy())

    def kaydet(self):
        try:
            df = pd.read_excel(self.kurallar_yolu, sheet_name='ESLESME')
        except Exception as e:
            messagebox.showerror('Hata', 'ESLESME okunamadi: ' + str(e))
            return

        yeni_satirlar = []
        for item in self.tree.get_children():
            vals = self.tree.item(item, 'values')
            if len(vals) < 3:
                continue
            modul, excel_baslik, dia_alan = vals[0], vals[1], vals[2]
            if not modul or not excel_baslik or not dia_alan:
                continue
            if excel_baslik.startswith('('):
                continue
            yeni_satirlar.append({
                'modul': modul,
                'excel_baslik': excel_baslik,
                'dia_alan': dia_alan
            })

        if not yeni_satirlar:
            messagebox.showinfo('Bilgi', 'Kaydedilecek esleme yok.')
            return

        df = pd.concat([df, pd.DataFrame(yeni_satirlar)], ignore_index=True)
        df = df.drop_duplicates(subset=['modul', 'excel_baslik'], keep='last')

        try:
            sayfalar = {}
            # Okuma bitince dosya kapanmali; yoksa Windows'ta ayni dosyaya yazilamaz
            with pd.ExcelFile(self.kurallar_yolu) as xl:
                for s in xl.sheet_names:
                    if s == 'ESLESME':
                        sayfalar[s] = df
                    else:
                        sayfalar[s] = pd.read_excel(xl, sheet_name=s)
            with pd.ExcelWriter(self.kurallar_yolu, engine='openpyxl') as w:
                for s, d in sayfalar.items():
                    d.to_excel(w, sheet_name=s, index=False)
            messagebox.showinfo('Basarili',
                                str(len(yeni_satirlar)) + ' esleme kaydedildi.')
            self.pencere.destroy()
        except Exception as e:
            messagebox.showerror('Hata', 'Kaydedilemedi: ' + str(e))


# AYARLAR PENCERESI
class AyarlarPenceresi:
    def __init__(self, parent, musteri_klasoru, program_klasoru):
        self.musteri_klasoru = musteri_klasoru
        self.program_klasoru = program_klasoru

        # kurallar.xlsx musteri klasorunde yoksa program klasorunden kopyala
        self.kurallar_yolu = os.path.join(musteri_klasoru, 'kurallar.xlsx')
        if not os.path.exists(self.kurallar_yolu):
            kaynak = os.path.join(program_klasoru, 'kurallar.xlsx')
            if os.path.exists(kaynak):
                shutil.copy2(kaynak, self.kurallar_yolu)
            else:
                from _ortak import kurallari_excele_yaz
                from kurallar import KURALLAR
                kurallari_excele_yaz(KURALLAR, self.kurallar_yolu)

        self.pencere = tk.Toplevel(parent)
        self.pencere.title('Ayarlar - kurallar.xlsx')
        self.pencere.geometry('1000x700')
        self.pencere.configure(bg='#f0f0f0')
        self.pencere.transient(parent)
        self.pencere.grab_set()

        self.modul_var = tk.StringVar(value='cari')
        self.arayuz_olustur()
        self.veri_yukle()

    def arayuz_olustur(self):
        # Ust: modul secimi
        ust = tk.Frame(self.pencere, bg='#2c3e50', pady=10)
        ust.pack(fill=tk.X)

        tk.Label(ust, text='MODUL:', font=('Arial', 11, 'bold'),
                 bg='#2c3e50', fg='white').pack(side=tk.LEFT, padx=10)

        moduller = ['cari', 'stok', 'banka', 'personel',
                    'demirbas', 'hesapplan', 'muhkod', 'urunagaci']

        combo = ttk.Combobox(ust, textvariable=self.modul_var,
                             values=moduller, state='readonly',
                             width=20, font=('Arial', 11))
        combo.pack(side=tk.LEFT, padx=5)
        combo.bind('<<ComboboxSelected>>', lambda e: self.veri_yukle())

        tk.Button(ust, text='Yukle', command=self.veri_yukle,
                  bg='#3498db', fg='white',
                  font=('Arial', 10, 'bold')).pack(side=tk.LEFT, padx=10)

        # Notebook (sekmeler)
        self.notebook = ttk.Notebook(self.pencere)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.mukerrer_frame = tk.Frame(self.notebook, bg='white')
        self.eslesme_frame = tk.Frame(self.notebook, bg='white')
        self.varsayilan_frame = tk.Frame(self.notebook, bg='white')
        self.ayarlar_frame = tk.Frame(self.notebook, bg='white')

        self.notebook.add(self.mukerrer_frame, text='  Mukerrer  ')
        self.notebook.add(self.eslesme_frame, text='  Eslesme  ')
        self.notebook.add(self.varsayilan_frame, text='  Varsayilan  ')
        self.notebook.add(self.ayarlar_frame, text='  Ayarlar  ')

        self.mukerrer_tablo()
        self.eslesme_tablo()
        self.varsayilan_tablo()
        self.ayarlar_tablo()

        # Alt butonlar
        alt = tk.Frame(self.pencere, bg='#f0f0f0', pady=10)
        alt.pack(fill=tk.X, padx=10)

        tk.Button(alt, text='💾 KAYDET', command=self.kaydet,
                  bg='#27ae60', fg='white', font=('Arial', 12, 'bold'),
                  padx=30, pady=8).pack(side=tk.RIGHT, padx=5)

        tk.Button(alt, text='Iptal', command=self.pencere.destroy,
                  bg='#95a5a6', fg='white', font=('Arial', 11),
                  padx=20, pady=8).pack(side=tk.RIGHT, padx=5)

    def mukerrer_tablo(self):
        cerceve = tk.Frame(self.mukerrer_frame, bg='white')
        cerceve.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        tk.Label(cerceve,
                 text='Mukerrer kontrol: 2+ kolon = sil, 1 kolon = isaretle',
                 font=('Arial', 9, 'italic'), bg='white',
                 fg='#555').pack(anchor='w', pady=5)

        self.muk_tree = ttk.Treeview(cerceve,
                                     columns=('sutunlar', 'aksiyon'),
                                     show='headings', height=15)
        self.muk_tree.heading('sutunlar', text='Sutunlar (virgulle)')
        self.muk_tree.heading('aksiyon', text='Aksiyon')
        self.muk_tree.column('sutunlar', width=600)
        self.muk_tree.column('aksiyon', width=150)
        self.muk_tree.pack(fill=tk.BOTH, expand=True)

        btn = tk.Frame(cerceve, bg='white')
        btn.pack(fill=tk.X, pady=5)

        tk.Button(btn, text='+ Yeni', command=lambda: self.ekle(self.muk_tree),
                  bg='#3498db', fg='white', padx=10).pack(side=tk.LEFT, padx=2)
        tk.Button(btn, text='🗑 Sil', command=lambda: self.sil(self.muk_tree),
                  bg='#e74c3c', fg='white', padx=10).pack(side=tk.LEFT, padx=2)
        tk.Button(btn, text='✏ Duzenle',
                  command=lambda: self.duzenle(self.muk_tree, ['sutunlar', 'aksiyon']),
                  bg='#f39c12', fg='white', padx=10).pack(side=tk.LEFT, padx=2)

    def eslesme_tablo(self):
        cerceve = tk.Frame(self.eslesme_frame, bg='white')
        cerceve.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        tk.Label(cerceve, text='Excel basligi -> DIA alani eslestirmesi',
                 font=('Arial', 9, 'italic'), bg='white',
                 fg='#555').pack(anchor='w', pady=5)

        self.es_tree = ttk.Treeview(cerceve,
                                    columns=('excel', 'dia'),
                                    show='headings', height=15)
        self.es_tree.heading('excel', text='Excel Basligi')
        self.es_tree.heading('dia', text='DIA Alani')
        self.es_tree.column('excel', width=450)
        self.es_tree.column('dia', width=350)
        self.es_tree.pack(fill=tk.BOTH, expand=True)

        btn = tk.Frame(cerceve, bg='white')
        btn.pack(fill=tk.X, pady=5)

        tk.Button(btn, text='+ Yeni', command=lambda: self.ekle(self.es_tree),
                  bg='#3498db', fg='white', padx=10).pack(side=tk.LEFT, padx=2)
        tk.Button(btn, text='🗑 Sil', command=lambda: self.sil(self.es_tree),
                  bg='#e74c3c', fg='white', padx=10).pack(side=tk.LEFT, padx=2)
        tk.Button(btn, text='✏ Duzenle',
                  command=lambda: self.duzenle(self.es_tree, ['excel', 'dia']),
                  bg='#f39c12', fg='white', padx=10).pack(side=tk.LEFT, padx=2)

    def varsayilan_tablo(self):
        cerceve = tk.Frame(self.varsayilan_frame, bg='white')
        cerceve.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        tk.Label(cerceve, text="Excel'de olmayan alanlar icin varsayilan degerler",
                 font=('Arial', 9, 'italic'), bg='white',
                 fg='#555').pack(anchor='w', pady=5)

        self.va_tree = ttk.Treeview(cerceve,
                                    columns=('alan', 'deger'),
                                    show='headings', height=15)
        self.va_tree.heading('alan', text='Alan')
        self.va_tree.heading('deger', text='Deger')
        self.va_tree.column('alan', width=450)
        self.va_tree.column('deger', width=350)
        self.va_tree.pack(fill=tk.BOTH, expand=True)

        btn = tk.Frame(cerceve, bg='white')
        btn.pack(fill=tk.X, pady=5)

        tk.Button(btn, text='+ Yeni', command=lambda: self.ekle(self.va_tree),
                  bg='#3498db', fg='white', padx=10).pack(side=tk.LEFT, padx=2)
        tk.Button(btn, text='🗑 Sil', command=lambda: self.sil(self.va_tree),
                  bg='#e74c3c', fg='white', padx=10).pack(side=tk.LEFT, padx=2)
        tk.Button(btn, text='✏ Duzenle',
                  command=lambda: self.duzenle(self.va_tree, ['alan', 'deger']),
                  bg='#f39c12', fg='white', padx=10).pack(side=tk.LEFT, padx=2)

    def ayarlar_tablo(self):
        cerceve = tk.Frame(self.ayarlar_frame, bg='white')
        cerceve.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        tk.Label(cerceve, text='AYARLAR sayfasi: girdi/cikti dosya, kontrol ayarlari',
                 font=('Arial', 9, 'italic'), bg='white',
                 fg='#555').pack(anchor='w', pady=5)

        ay_frame = tk.Frame(cerceve, bg='white')
        ay_frame.pack(fill=tk.BOTH, expand=True)

        self.ay_tree = ttk.Treeview(ay_frame,
                                    columns=('ayar', 'deger', 'aciklama', 'formul', 'ornek'),
                                    show='headings', height=14)
        self.ay_tree.heading('ayar', text='Ayar')
        self.ay_tree.heading('deger', text='Deger')
        self.ay_tree.heading('aciklama', text='Aciklama (ne ise yarar?)')
        self.ay_tree.heading('formul', text='Formul / Kural')
        self.ay_tree.heading('ornek', text='Ornek')
        self.ay_tree.column('ayar', width=150, minwidth=100)
        self.ay_tree.column('deger', width=180, minwidth=100)
        self.ay_tree.column('aciklama', width=320, minwidth=200)
        self.ay_tree.column('formul', width=320, minwidth=200)
        self.ay_tree.column('ornek', width=200, minwidth=100)

        # Yatay kaydirma
        xscroll = ttk.Scrollbar(ay_frame, orient='horizontal', command=self.ay_tree.xview)
        self.ay_tree.configure(xscrollcommand=xscroll.set)
        xscroll.pack(side=tk.BOTTOM, fill=tk.X)
        self.ay_tree.pack(fill=tk.BOTH, expand=True)
        self.ay_tree.bind('<Double-1>', self._ayar_cift_tikla)
        self.ay_tree.bind('<Return>', self._ayar_cift_tikla)

        # Ipucu
        tk.Label(cerceve, text='Satira cift tiklayin -> detayli aciklama acilir',
                 font=('Arial', 9, 'italic'), bg='white', fg='#7f8c8d'
                 ).pack(anchor='w', pady=(3, 0))

        btn = tk.Frame(cerceve, bg='white')
        btn.pack(fill=tk.X, pady=5)

        tk.Button(btn, text='+ Yeni', command=lambda: self.ekle(self.ay_tree),
                  bg='#3498db', fg='white', padx=10).pack(side=tk.LEFT, padx=2)
        tk.Button(btn, text=' Sil', command=lambda: self.sil(self.ay_tree),
                  bg='#e74c3c', fg='white', padx=10).pack(side=tk.LEFT, padx=2)
        tk.Button(btn, text=' Duzenle',
                  command=lambda: self.duzenle(self.ay_tree, ['ayar', 'deger', 'aciklama', 'formul', 'ornek']),
                  bg='#f39c12', fg='white', padx=10).pack(side=tk.LEFT, padx=2)
        tk.Button(btn, text='ℹ Detay', command=self.ayar_detay_goster,
                  bg='#9b59b6', fg='white', padx=10).pack(side=tk.LEFT, padx=2)
        tk.Button(btn, text='＋ Eksik Ayarlari Ekle', command=self.eksik_ayarlari_ekle,
                  bg='#16a085', fg='white', padx=10).pack(side=tk.LEFT, padx=2)

    def eksik_ayarlari_ekle(self):
        """Yeni parametrik ayarlardan bu modulde olmayanlari varsayilan degerle ekler."""
        from _ortak import AYAR_TANIMLARI
        mevcut = {str(self.ay_tree.item(i, 'values')[0]) for i in self.ay_tree.get_children()}
        eklenen = 0
        for a in AYAR_TANIMLARI:
            if a['ayar'] in mevcut:
                continue
            deger = a['varsayilan']
            if isinstance(deger, bool):
                deger = 'True' if deger else 'False'
            formul = a['formul']
            if a.get('secenekler'):
                formul = 'Secenekler: ' + ' / '.join(a['secenekler']) + '  |  ' + formul
            self.ay_tree.insert('', 'end', values=(a['ayar'], deger, a['aciklama'],
                                                   formul, a['ornek']))
            eklenen += 1
        if eklenen:
            messagebox.showinfo('Bilgi', str(eklenen) + " ayar eklendi. Degistirmek icin satiri secip "
                                "'Duzenle' butonuna basin; kalici olmasi icin KAYDET'e basin.",
                                parent=self.pencere)
        else:
            messagebox.showinfo('Bilgi', 'Tum ayarlar zaten var.', parent=self.pencere)

    def _ayar_cift_tikla(self, event=None):
        """Double-1 veya Return ile cagrilir."""
        try:
            self.ayar_detay_goster()
        except Exception as e:
            import traceback
            print('HATA (popup): ' + str(e))
            traceback.print_exc()

    def ayar_detay_goster(self):
        sec = self.ay_tree.selection()
        if not sec:
            return
        v = self.ay_tree.item(sec[0], 'values')
        if not v:
            return

        ayar_adi = str(v[0]) if len(v) > 0 else ''
        deger = str(v[1]) if len(v) > 1 else ''
        aciklama = str(v[2]) if len(v) > 2 else ''
        formul = str(v[3]) if len(v) > 3 else ''
        ornek = str(v[4]) if len(v) > 4 else ''

        if not (formul and ornek and aciklama):
            try:
                df = pd.read_excel(self.kurallar_yolu, sheet_name='AYARLAR')
                modul = self.modul_var.get()
                satir = df[(df['modul'] == modul) & (df['ayar'] == ayar_adi)]
                if len(satir) > 0:
                    r = satir.iloc[0]
                    if not aciklama and 'aciklama' in df.columns and pd.notna(r.get('aciklama')):
                        aciklama = str(r['aciklama'])
                    if not formul and 'formul' in df.columns and pd.notna(r.get('formul')):
                        formul = str(r['formul'])
                    if not ornek and 'ornek' in df.columns and pd.notna(r.get('ornek')):
                        ornek = str(r['ornek'])
            except Exception as e:
                print('Detay okunamadi: ' + str(e))

        p = tk.Toplevel(self.pencere)
        p.title('Ayar Detayi: ' + ayar_adi)
        p.geometry('720x560')
        p.configure(bg='white')
        p.transient(self.pencere)
        p.grab_set()

        bas = tk.Frame(p, bg='#2c3e50', pady=14)
        bas.pack(fill=tk.X)
        tk.Label(bas, text=ayar_adi, font=('Arial', 15, 'bold'),
                 bg='#2c3e50', fg='white').pack()
        tk.Label(bas, text='Mevcut deger: ' + deger, font=('Arial', 10),
                 bg='#2c3e50', fg='#bdc3c7').pack()

        orta = tk.Frame(p, bg='white')
        orta.pack(fill=tk.BOTH, expand=True, padx=10, pady=8)
        canvas = tk.Canvas(orta, bg='white', highlightthickness=0)
        vscroll = ttk.Scrollbar(orta, orient='vertical', command=canvas.yview)
        canvas.configure(yscrollcommand=vscroll.set)
        vscroll.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        ic = tk.Frame(canvas, bg='white')
        cw = canvas.create_window((0, 0), window=ic, anchor='nw')
        ic.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.bind('<Configure>', lambda e: canvas.itemconfig(cw, width=e.width))

        def blok(baslik, metin, renk):
            tk.Label(ic, text=baslik, font=('Arial', 11, 'bold'),
                     bg='white', fg=renk, anchor='w').pack(fill=tk.X, pady=(10, 3), padx=15)
            f2 = tk.Frame(ic, bg='#f8f9fa', relief=tk.SUNKEN, bd=1)
            f2.pack(fill=tk.X, padx=15)
            tk.Label(f2, text=metin if metin else '(bos)',
                     font=('Consolas', 10), bg='#f8f9fa', fg='#2c3e50',
                     anchor='w', justify='left', wraplength=640,
                     padx=10, pady=8).pack(fill=tk.X)

        blok('[1] ACIKLAMA (Ne ise yarar?)', aciklama, '#2c3e50')
        blok('[2] FORMUL / KURAL', formul, '#8e44ad')
        blok('[3] ORNEK', ornek, '#16a085')

        tk.Button(p, text='Kapat', command=p.destroy,
                  bg='#95a5a6', fg='white', font=('Arial', 11, 'bold'),
                  padx=25, pady=6).pack(pady=10)

    def veri_yukle(self):
        modul = self.modul_var.get()

        self.muk_tree.delete(*self.muk_tree.get_children())
        self.es_tree.delete(*self.es_tree.get_children())
        self.va_tree.delete(*self.va_tree.get_children())
        self.ay_tree.delete(*self.ay_tree.get_children())

        # MUKERRER
        try:
            df = pd.read_excel(self.kurallar_yolu, sheet_name='MUKERRER')
            df = df[df['modul'] == modul]
            for _, r in df.iterrows():
                self.muk_tree.insert('', 'end',
                                     values=(r['sutunlar'], r['aksiyon']))
        except Exception as e:
            print('MUKERRER: ' + str(e))

        # ESLESME
        try:
            df = pd.read_excel(self.kurallar_yolu, sheet_name='ESLESME')
            df = df[df['modul'] == modul]
            for _, r in df.iterrows():
                self.es_tree.insert('', 'end',
                                    values=(r['excel_baslik'], r['dia_alan']))
        except Exception as e:
            print('ESLESME: ' + str(e))

        # VARSAYILAN
        try:
            df = pd.read_excel(self.kurallar_yolu, sheet_name='VARSAYILAN')
            df = df[df['modul'] == modul]
            for _, r in df.iterrows():
                self.va_tree.insert('', 'end', values=(r['alan'], r['deger']))
        except Exception as e:
            print('VARSAYILAN: ' + str(e))

        # AYARLAR
        try:
            df = pd.read_excel(self.kurallar_yolu, sheet_name='AYARLAR')
            df = df[df['modul'] == modul]
            for _, r in df.iterrows():
                deger = r['deger'] if pd.notna(r['deger']) else ''
                ac = r['aciklama'] if 'aciklama' in df.columns and pd.notna(r['aciklama']) else ''
                fo = r['formul'] if 'formul' in df.columns and pd.notna(r['formul']) else ''
                orn = r['ornek'] if 'ornek' in df.columns and pd.notna(r['ornek']) else ''
                self.ay_tree.insert('', 'end', values=(r['ayar'], deger, ac, fo, orn))
        except Exception as e:
            print('AYARLAR: ' + str(e))

    def ekle(self, tree):
        sutunlar = tree['columns']
        self._dialog(tree, sutunlar, [], 'Yeni Kayit')

    def duzenle(self, tree, sutunlar):
        secili = tree.selection()
        if not secili:
            messagebox.showwarning('Uyari', 'Satir secin.')
            return
        if len(secili) > 1:
            messagebox.showwarning('Uyari', 'Tek satir secin.')
            return
        item = secili[0]
        mevcut = list(tree.item(item, 'values'))
        self._dialog(tree, sutunlar, mevcut, 'Duzenle', item)

    def _dialog(self, tree, sutunlar, mevcut, baslik, item=None):
        p = tk.Toplevel(self.pencere)
        p.title(baslik)
        p.geometry('550x' + str(120 + len(sutunlar) * 45))
        p.transient(self.pencere)
        p.grab_set()

        girisler = {}
        for i, s in enumerate(sutunlar):
            tk.Label(p, text=s + ':', font=('Arial', 10, 'bold')).grid(
                row=i, column=0, sticky='w', padx=10, pady=8)
            g = tk.Entry(p, width=45, font=('Arial', 10))
            if i < len(mevcut):
                g.insert(0, str(mevcut[i]))
            g.grid(row=i, column=1, padx=10, pady=8)
            girisler[s] = g

        def kaydet():
            degerler = [girisler[s].get() for s in sutunlar]
            if item:
                tree.item(item, values=degerler)
            else:
                tree.insert('', 'end', values=degerler)
            p.destroy()

        tk.Button(p, text='Kaydet', command=kaydet,
                  bg='#27ae60', fg='white', font=('Arial', 11, 'bold'),
                  padx=25, pady=8).grid(row=len(sutunlar), column=0,
                                        columnspan=2, pady=15)

    def sil(self, tree):
        secili = tree.selection()
        if not secili:
            messagebox.showwarning('Uyari', 'Satir secin.')
            return
        for item in secili:
            tree.delete(item)

    def kaydet(self):
        try:
            sayfalar = {}
            # Okuma bitince dosya kapanmali; yoksa Windows'ta ayni dosyaya yazilamaz
            with pd.ExcelFile(self.kurallar_yolu) as xls:
                for s in xls.sheet_names:
                    sayfalar[s] = pd.read_excel(xls, sheet_name=s)

            modul = self.modul_var.get()

            # MUKERRER
            if 'MUKERRER' in sayfalar:
                df = sayfalar['MUKERRER']
                df = df[df['modul'] != modul]
                yeni = []
                for item in self.muk_tree.get_children():
                    v = self.muk_tree.item(item, 'values')
                    yeni.append({'modul': modul, 'sutunlar': v[0],
                                 'aksiyon': v[1], 'aciklama': ''})
                if yeni:
                    df = pd.concat([df, pd.DataFrame(yeni)], ignore_index=True)
                sayfalar['MUKERRER'] = df

            # ESLESME
            if 'ESLESME' in sayfalar:
                df = sayfalar['ESLESME']
                df = df[df['modul'] != modul]
                yeni = []
                for item in self.es_tree.get_children():
                    v = self.es_tree.item(item, 'values')
                    yeni.append({'modul': modul, 'excel_baslik': v[0],
                                 'dia_alan': v[1]})
                if yeni:
                    df = pd.concat([df, pd.DataFrame(yeni)], ignore_index=True)
                sayfalar['ESLESME'] = df

            # VARSAYILAN
            if 'VARSAYILAN' in sayfalar:
                df = sayfalar['VARSAYILAN']
                df = df[df['modul'] != modul]
                yeni = []
                for item in self.va_tree.get_children():
                    v = self.va_tree.item(item, 'values')
                    yeni.append({'modul': modul, 'alan': v[0], 'deger': v[1]})
                if yeni:
                    df = pd.concat([df, pd.DataFrame(yeni)], ignore_index=True)
                sayfalar['VARSAYILAN'] = df

            # AYARLAR
            if 'AYARLAR' in sayfalar:
                df = sayfalar['AYARLAR']
                df = df[df['modul'] != modul]
                yeni = []
                for item in self.ay_tree.get_children():
                    v = self.ay_tree.item(item, 'values')
                    ac = v[2] if len(v) > 2 else ''
                    fo = v[3] if len(v) > 3 else ''
                    orn = v[4] if len(v) > 4 else ''
                    yeni.append({'modul': modul, 'ayar': v[0], 'deger': v[1],
                                 'aciklama': ac, 'formul': fo, 'ornek': orn})
                if yeni:
                    df = pd.concat([df, pd.DataFrame(yeni)], ignore_index=True)
                sayfalar['AYARLAR'] = df

            with pd.ExcelWriter(self.kurallar_yolu, engine='openpyxl') as w:
                for s, df in sayfalar.items():
                    df.to_excel(w, sheet_name=s, index=False)

            messagebox.showinfo('Basarili', 'Ayarlar kaydedildi.')
            self.pencere.destroy()
        except Exception as e:
            messagebox.showerror('Hata', 'Kaydedilemedi: ' + str(e))


if __name__ == '__main__':
    root = tk.Tk()
    app = DiaGUI(root)
    root.mainloop()
