"""📖 Kullanim Rehberi penceresi (icerik: rehber.py)."""
import os
import tkinter as tk
import webbrowser
from tkinter import ttk, filedialog, messagebox

import rehber


class RehberPenceresi:
    def __init__(self, parent, ara=None, geri_grab=None):
        """ara: acilista bulunup gosterilecek metin (or. bir ayar adi).
        geri_grab: rehber kapaninca tekrar modal yapilacak pencere (Ayarlar)."""
        self.parent = parent
        self.onceki_grab = geri_grab or parent.grab_current()
        self.pencere = tk.Toplevel(parent)
        self.pencere.title('Kullanım Rehberi')
        self.pencere.geometry('1150x760')
        self.pencere.configure(bg='#f0f0f0')
        self.pencere.transient(parent)
        self.pencere.protocol('WM_DELETE_WINDOW', self.kapat)
        self.bolumler = rehber.bolumler()
        self.bulunanlar = []
        self.bulunan_no = -1
        self.arayuz_olustur()
        self.icerik_yaz()
        # Acan pencere (Ayarlar) grab tutuyorsa rehbere tiklanamaz: grab'i rehbere al
        self.pencere.grab_set()
        if ara:
            self.ayar_goster(ara)

    # ------------------------------------------------------------
    def arayuz_olustur(self):
        ust = tk.Frame(self.pencere, bg='#2c3e50', pady=8)
        ust.pack(fill=tk.X)
        tk.Label(ust, text='📖 KULLANIM REHBERİ', font=('Arial', 13, 'bold'),
                 bg='#2c3e50', fg='white').pack(side=tk.LEFT, padx=12)
        tk.Button(ust, text='Kapat', command=self.kapat, bg='#95a5a6', fg='white',
                  padx=12).pack(side=tk.RIGHT, padx=6)
        tk.Button(ust, text='💾 HTML Kaydet / Yazdır', command=self.html_kaydet, bg='#27ae60', fg='white',
                  padx=10).pack(side=tk.RIGHT, padx=6)
        self.bulunan_lbl = tk.Label(ust, text='', bg='#2c3e50', fg='#f1c40f', font=('Arial', 9))
        self.bulunan_lbl.pack(side=tk.RIGHT, padx=6)
        tk.Button(ust, text='Ara / Sonraki', command=self.ara, bg='#3498db', fg='white',
                  padx=8).pack(side=tk.RIGHT)
        self.ara_var = tk.StringVar()
        giris = tk.Entry(ust, textvariable=self.ara_var, width=28, font=('Arial', 11))
        giris.pack(side=tk.RIGHT, padx=6)
        giris.bind('<Return>', lambda e: self.ara())
        self.ara_var.trace_add('write', lambda *a: self._arama_sifirla())
        tk.Label(ust, text='Ara:', bg='#2c3e50', fg='white').pack(side=tk.RIGHT)

        govde = tk.PanedWindow(self.pencere, orient=tk.HORIZONTAL, sashwidth=5, bg='#f0f0f0')
        govde.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        sol = tk.Frame(govde, bg='white')
        self.liste = tk.Listbox(sol, font=('Arial', 10), activestyle='none', bd=0,
                                selectbackground='#3498db', exportselection=False)
        self.liste.pack(fill=tk.BOTH, expand=True)
        for b in self.bolumler:
            self.liste.insert(tk.END, b['baslik'])
        self.liste.bind('<<ListboxSelect>>', self._bolum_sec)
        govde.add(sol, width=290)

        sag = tk.Frame(govde, bg='white')
        kaydirma = ttk.Scrollbar(sag, orient='vertical')
        self.metin = tk.Text(sag, wrap='word', font=('Arial', 10), bg='white', fg='#222', padx=14, pady=8,
                             yscrollcommand=kaydirma.set, bd=0, spacing1=1, spacing3=1)
        kaydirma.config(command=self.metin.yview)
        kaydirma.pack(side=tk.RIGHT, fill=tk.Y)
        self.metin.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        govde.add(sag)

        m = self.metin
        m.tag_configure('b1', font=('Arial', 14, 'bold'), foreground='white', background='#2c3e50',
                        spacing1=14, spacing3=6, lmargin1=4)
        m.tag_configure('b2', font=('Arial', 11, 'bold'), foreground='#2c3e50', spacing1=12, spacing3=3)
        m.tag_configure('p', lmargin1=4, lmargin2=4, spacing3=4)
        m.tag_configure('madde', lmargin1=14, lmargin2=30)
        m.tag_configure('kod', font=('Consolas', 10), background='#f4f6f7', lmargin1=24, lmargin2=24,
                        rmargin=20, spacing1=2, spacing3=2)
        m.tag_configure('not', background='#fff3cd', lmargin1=8, lmargin2=8, rmargin=20, spacing1=4, spacing3=4)
        m.tag_configure('kalin', font=('Arial', 10, 'bold'))
        m.tag_configure('tablo', lmargin1=14, lmargin2=34)
        m.tag_configure('tablo_bas', font=('Arial', 10, 'bold'), foreground='#2c3e50', lmargin1=14, lmargin2=34,
                        spacing1=5)
        m.tag_configure('bul', background='#f9e79f')
        m.tag_configure('bul_aktif', background='#f39c12')

    # ------------------------------------------------------------
    def _yaz(self, metin, *etiketler):
        self.metin.insert(tk.END, metin, etiketler)

    def icerik_yaz(self):
        m = self.metin
        m.config(state='normal')
        m.delete('1.0', tk.END)
        self.isaretler = []
        self.basliklar = {}      # alt baslik metni -> mark adi (mark adinda bosluk olmamali)
        for n, b in enumerate(self.bolumler):
            isaret = 'bolum%d' % n
            m.mark_set(isaret, m.index('end-1c'))
            m.mark_gravity(isaret, 'left')
            self.isaretler.append(isaret)
            self._yaz(' ' + b['baslik'] + '\n', 'b1')
            for tur, icerik in b['parcalar']:
                if tur == 'p':
                    self._yaz(icerik + '\n', 'p')
                elif tur == 'h':
                    if icerik not in self.basliklar:
                        isaret_h = 'baslik%d' % len(self.basliklar)
                        m.mark_set(isaret_h, m.index('end-1c'))
                        m.mark_gravity(isaret_h, 'left')
                        self.basliklar[icerik] = isaret_h
                    self._yaz(icerik + '\n', 'b2')
                elif tur == 'madde':
                    for x in icerik:
                        self._yaz('•  ' + x + '\n', 'madde')
                elif tur == 'adim':
                    for i, x in enumerate(icerik, 1):
                        self._yaz(str(i) + '.  ', 'madde', 'kalin')
                        self._yaz(x + '\n', 'madde')
                elif tur == 'kod':
                    self._yaz(icerik + '\n', 'kod')
                elif tur == 'not':
                    self._yaz('⚠ Dikkat: ', 'not', 'kalin')
                    self._yaz(icerik + '\n', 'not')
                elif tur == 'tablo':
                    self._tablo_yaz(*icerik)
            self._yaz('\n')
        m.config(state='disabled')

    def _tablo_yaz(self, baslik, satirlar):
        """Genis tablolar metin kutusuna sigmaz: her satir bir blok, ilk hucre kalin."""
        anahtar_deger = not any(baslik)
        for s in satirlar:
            s = [str(x) for x in s]
            if anahtar_deger:
                self._yaz(s[0] + ': ', 'tablo', 'kalin')
                self._yaz(s[1] + '\n', 'tablo')
                continue
            self._yaz('▸ ' + s[0] + '\n', 'tablo_bas')
            for h, x in zip(baslik[1:], s[1:]):
                if x:
                    self._yaz(h + ': ', 'tablo', 'kalin')
                    self._yaz(x.replace('\n', ' | ') + '\n', 'tablo')

    # ------------------------------------------------------------
    def _bolum_sec(self, _e=None):
        sec = self.liste.curselection()
        if sec:
            self.metin.yview(self.isaretler[sec[0]])

    def _bolum_isaretle(self, index):
        """index'in bulundugu bolumu soldaki listede secili yapar."""
        satir = int(self.metin.index(index).split('.')[0])
        secili = 0
        for n, isaret in enumerate(self.isaretler):
            if int(self.metin.index(isaret).split('.')[0]) <= satir:
                secili = n
        self.liste.selection_clear(0, tk.END)
        self.liste.selection_set(secili)
        self.liste.see(secili)

    def ayar_goster(self, ad):
        """Ayar referansindaki basliga gider; yoksa metin olarak arar."""
        isaret = self.basliklar.get(ad)
        if isaret:
            self.metin.yview(isaret)
            bitis = self.metin.index(isaret + ' lineend')
            self.metin.tag_add('bul_aktif', isaret, bitis)
            self._bolum_isaretle(isaret)
        else:
            self.ara_var.set(ad)
            self.ara()

    def _arama_sifirla(self):
        self.bulunanlar = []
        self.bulunan_no = -1
        self.metin.tag_remove('bul', '1.0', tk.END)
        self.metin.tag_remove('bul_aktif', '1.0', tk.END)
        self.bulunan_lbl.config(text='')

    @staticmethod
    def _katla(metin):
        """Turkce buyuk/kucuk harf duyarsiz karsilastirma; uzunluk korunur (İ -> i, I -> ı)."""
        sonuc = []
        for c in metin:
            k = 'i' if c == 'İ' else 'ı' if c == 'I' else c.lower()
            sonuc.append(k if len(k) == 1 else c)
        return ''.join(sonuc)

    def ara(self):
        aranan = self._katla(self.ara_var.get().strip())
        if not aranan:
            return
        m = self.metin
        if not self.bulunanlar:
            # Tk'nin nocase aramasi 'İ' iceren metinde coker (Tk 8.6 hatasi): arama Python'da yapilir
            satirlar = m.get('1.0', 'end-1c').split('\n')
            for no, satir in enumerate(satirlar, 1):
                katli = self._katla(satir)
                kol = katli.find(aranan)
                while kol >= 0:
                    bas, son = '%d.%d' % (no, kol), '%d.%d' % (no, kol + len(aranan))
                    m.tag_add('bul', bas, son)
                    self.bulunanlar.append((bas, son))
                    kol = katli.find(aranan, kol + len(aranan))
            if not self.bulunanlar:
                self.bulunan_lbl.config(text='bulunamadı')
                return
        m.tag_remove('bul_aktif', '1.0', tk.END)
        self.bulunan_no = (self.bulunan_no + 1) % len(self.bulunanlar)
        bas, son = self.bulunanlar[self.bulunan_no]
        m.tag_add('bul_aktif', bas, son)
        m.see(bas)
        self._bolum_isaretle(bas)
        self.bulunan_lbl.config(text='%d / %d' % (self.bulunan_no + 1, len(self.bulunanlar)))

    # ------------------------------------------------------------
    def html_kaydet(self):
        yol = filedialog.asksaveasfilename(parent=self.pencere, title='Rehberi kaydet',
                                           defaultextension='.html', initialfile='DIA_Kullanim_Rehberi.html',
                                           filetypes=[('Web sayfası', '*.html')])
        if not yol:
            return
        try:
            with open(yol, 'w', encoding='utf-8') as f:
                f.write(rehber.html_sayfa())
        except OSError as e:
            messagebox.showerror('Hata', 'Kaydedilemedi: ' + str(e), parent=self.pencere)
            return
        webbrowser.open('file:///' + os.path.abspath(yol).replace('\\', '/'))

    def kapat(self):
        try:
            self.pencere.grab_release()
        except tk.TclError:
            pass
        self.pencere.destroy()
        if self.onceki_grab is not None:
            try:
                self.onceki_grab.grab_set()
            except tk.TclError:
                pass
