"""Logo'dan Cek / DIA'ya Gonder pencereleri.

Baglanti ayarlari musteri klasorundeki entegrasyon.json dosyasina yazilir.
Sifreler hicbir zaman dosyaya yazilmaz.
"""
import json
import os
import queue
import shutil
import threading
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext

AYAR_DOSYASI = 'entegrasyon.json'


def ayar_oku(klasor):
    try:
        with open(os.path.join(klasor, AYAR_DOSYASI), encoding='utf-8') as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


def ayar_yaz(klasor, bolum, degerler):
    ayarlar = ayar_oku(klasor)
    ayarlar[bolum] = {k: v for k, v in degerler.items() if 'sifre' not in k}
    with open(os.path.join(klasor, AYAR_DOSYASI), 'w', encoding='utf-8') as f:
        json.dump(ayarlar, f, ensure_ascii=False, indent=2)


def kurallari_getir(klasor):
    import dia_isle
    dia_isle.KLASOR = klasor
    return dia_isle.kurallari_yukle()


class _ArkaPlanPencere:
    """Uzun islemleri ayri thread'de calistirir; log satirlari kuyruk ile pencereye akar."""

    def __init__(self, parent, baslik, boyut, ana_log):
        self.pencere = tk.Toplevel(parent)
        self.pencere.title(baslik)
        self.pencere.geometry(boyut)
        self.pencere.configure(bg='#f0f0f0')
        self.pencere.transient(parent)
        self.ana_log = ana_log
        self._kuyruk = queue.Queue()
        self._calisiyor = False
        self.alanlar = {}

    def form(self, cerceve, tanimlar, kayitli):
        for i, (anahtar, etiket, varsayilan, gizli) in enumerate(tanimlar):
            tk.Label(cerceve, text=etiket, bg='#f0f0f0', font=('Arial', 10)).grid(
                row=i, column=0, sticky='w', padx=8, pady=4)
            var = tk.StringVar(value=str(kayitli.get(anahtar, varsayilan)))
            tk.Entry(cerceve, textvariable=var, width=40, show='•' if gizli else '',
                     font=('Arial', 10)).grid(row=i, column=1, sticky='w', padx=8, pady=4)
            self.alanlar[anahtar] = var

    def degerler(self):
        return {k: v.get().strip() for k, v in self.alanlar.items()}

    def _formu_al_guvenli(self):
        try:
            return self._formu_al()
        except ValueError as e:
            messagebox.showwarning('Eksik bilgi', str(e), parent=self.pencere)
            return None

    def log_alani(self):
        self.log_kutu = scrolledtext.ScrolledText(self.pencere, height=14, font=('Consolas', 9),
                                                  bg='#1e1e1e', fg='#d4d4d4')
        self.log_kutu.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

    def log(self, mesaj):
        self._kuyruk.put(str(mesaj))

    def _kuyrugu_bosalt(self):
        try:
            while True:
                m = self._kuyruk.get_nowait()
                self.log_kutu.insert(tk.END, m + '\n')
                self.log_kutu.see(tk.END)
                self.ana_log(m)
        except queue.Empty:
            pass
        if self._calisiyor:
            self.pencere.after(150, self._kuyrugu_bosalt)

    def arka_planda(self, is_fonksiyonu):
        if self._calisiyor:
            messagebox.showinfo('Bilgi', 'Bir islem zaten calisiyor.', parent=self.pencere)
            return

        def calistir():
            try:
                is_fonksiyonu()
            except Exception as e:
                self.log('HATA: ' + str(e))
            finally:
                self._calisiyor = False

        self._calisiyor = True
        threading.Thread(target=calistir, daemon=True).start()
        self.pencere.after(150, self._kuyrugu_bosalt)


# ============================================================
# LOGO'DAN CEK
# ============================================================
class LogoPenceresi(_ArkaPlanPencere):
    def __init__(self, parent, musteri_klasoru, ana_log):
        super().__init__(parent, "Logo'dan Cek", '720x620', ana_log)
        self.klasor = musteri_klasoru
        kayitli = ayar_oku(musteri_klasoru).get('logo', {})

        ust = tk.Frame(self.pencere, bg='#c0392b', pady=8)
        ust.pack(fill=tk.X)
        tk.Label(ust, text='LOGO (TIGER / GO) VERITABANINDAN CEK', font=('Arial', 12, 'bold'),
                 bg='#c0392b', fg='white').pack()
        tk.Label(self.pencere, text='Sadece okuma yapilir, Logo verisi degistirilmez. '
                                    'Kullanici bos birakilirsa Windows oturumuyla baglanilir.',
                 font=('Arial', 9, 'italic'), bg='#f0f0f0', fg='#555').pack(anchor='w', padx=10, pady=4)

        cerceve = tk.Frame(self.pencere, bg='#f0f0f0')
        cerceve.pack(fill=tk.X, padx=10)
        self.form(cerceve, [
            ('sunucu', 'SQL Sunucu (ör. SUNUCU\\LOGO):', '', False),
            ('veritabani', 'Veritabani (ör. TIGERDB):', '', False),
            ('kullanici', 'SQL Kullanici (bos = Windows):', '', False),
            ('sifre', 'SQL Sifre (kaydedilmez):', '', True),
            ('firma', 'Logo Firma No:', '1', False),
        ], kayitli)
        self.sadece_aktif = tk.BooleanVar(value=kayitli.get('sadece_aktif', True))
        tk.Checkbutton(cerceve, text='Sadece aktif kartlar', variable=self.sadece_aktif,
                       bg='#f0f0f0').grid(row=5, column=1, sticky='w', padx=8)

        btn = tk.Frame(self.pencere, bg='#f0f0f0', pady=8)
        btn.pack(fill=tk.X, padx=10)
        tk.Button(btn, text='Baglanti Testi', command=self.test, bg='#7f8c8d', fg='white',
                  padx=12).pack(side=tk.LEFT, padx=3)
        tk.Button(btn, text='Cari Cek', command=lambda: self.cek('cari'), bg='#27ae60', fg='white',
                  font=('Arial', 10, 'bold'), padx=14).pack(side=tk.LEFT, padx=3)
        tk.Button(btn, text='Stok Cek', command=lambda: self.cek('stok'), bg='#27ae60', fg='white',
                  font=('Arial', 10, 'bold'), padx=14).pack(side=tk.LEFT, padx=3)
        tk.Button(btn, text='Kapat', command=self.pencere.destroy, bg='#95a5a6', fg='white',
                  padx=12).pack(side=tk.RIGHT, padx=3)
        self.log_alani()

    def _formu_al(self):
        """Tk degiskenleri sadece ana thread'de okunabilir."""
        d = self.degerler()
        d['sadece_aktif'] = self.sadece_aktif.get()
        if not d['sunucu'] or not d['veritabani']:
            raise ValueError('Sunucu ve veritabani girin.')
        ayar_yaz(self.klasor, 'logo', d)
        return d

    def _baglan(self, d):
        from logo_kaynak import LogoKaynak
        self.log('Baglaniliyor: ' + d['sunucu'] + ' / ' + d['veritabani'])
        return LogoKaynak.baglan(d['sunucu'], d['veritabani'], d['kullanici'], d['sifre'])

    def test(self):
        d = self._formu_al_guvenli()
        if d is None:
            return

        def isi():
            kaynak = self._baglan(d)
            firmalar = kaynak.firmalar()
            self.log('Baglanti OK. Bulunan Logo firmalari: ' + (', '.join(firmalar) or 'YOK'))
        self.arka_planda(isi)

    def cek(self, modul_adi):
        d = self._formu_al_guvenli()
        if d is None:
            return

        def isi():
            from logo_kaynak import excele_yaz
            kaynak = self._baglan(d)
            kurallar = kurallari_getir(self.klasor)
            dosya = kurallar[modul_adi]['girdi_dosya']
            if modul_adi == 'cari':
                df = kaynak.cari_cek(d['firma'], d['sadece_aktif'])
            else:
                df = kaynak.stok_cek(d['firma'], d['sadece_aktif'])
            yol = os.path.join(self.klasor, dosya)
            if os.path.exists(yol):
                shutil.copy2(yol, yol + '.bak')
                self.log('  Eski dosya yedeklendi: ' + dosya + '.bak')
            excele_yaz(df, self.klasor, dosya)
            self.log(str(len(df)) + ' ' + modul_adi + ' karti yazildi: ' + dosya)
            self.log(">> Simdi ana ekranda modulu secip CALISTIR'a basin.")
        self.arka_planda(isi)


# ============================================================
# DIA'YA GONDER
# ============================================================
class DiaPenceresi(_ArkaPlanPencere):
    def __init__(self, parent, musteri_klasoru, ana_log):
        super().__init__(parent, "DIA'ya Gonder", '760x700', ana_log)
        self.klasor = musteri_klasoru
        kayitli = ayar_oku(musteri_klasoru).get('dia', {})

        ust = tk.Frame(self.pencere, bg='#2980b9', pady=8)
        ust.pack(fill=tk.X)
        tk.Label(ust, text="DIA WEB SERVIS AKTARIMI", font=('Arial', 12, 'bold'),
                 bg='#2980b9', fg='white').pack()
        tk.Label(self.pencere, text="Temizlenmis *_dia_hazir.xlsx dosyasi gonderilir. Kod DIA'da yoksa "
                                    "eklenir, varsa guncellenir. Her cagri 0,0125 kontor harcar.",
                 font=('Arial', 9, 'italic'), bg='#f0f0f0', fg='#555', wraplength=720,
                 justify='left').pack(anchor='w', padx=10, pady=4)

        cerceve = tk.Frame(self.pencere, bg='#f0f0f0')
        cerceve.pack(fill=tk.X, padx=10)
        self.form(cerceve, [
            ('sunucu_kodu', 'Sunucu Kodu (ör. racabilisim):', '', False),
            ('kullanici', 'Kullanici:', '', False),
            ('sifre', 'Sifre (kaydedilmez):', '', True),
            ('apikey', 'API Key:', '', True),
            ('firma_kodu', 'Firma Kodu:', '1', False),
            ('donem_kodu', 'Donem Kodu:', '1', False),
        ], kayitli)
        if not kayitli.get('apikey_hatirla'):
            self.alanlar['apikey'].set('')
        satir = 6
        tk.Label(cerceve, text='Modul:', bg='#f0f0f0').grid(row=satir, column=0, sticky='w', padx=8)
        self.modul = tk.StringVar(value=kayitli.get('modul', 'cari'))
        ttk.Combobox(cerceve, textvariable=self.modul, values=['cari', 'stok'], state='readonly',
                     width=12).grid(row=satir, column=1, sticky='w', padx=8, pady=4)
        tk.Label(cerceve, text="DIA'da varsa:", bg='#f0f0f0').grid(row=satir + 1, column=0, sticky='w', padx=8)
        self.mevcut_ise = tk.StringVar(value=kayitli.get('mevcut_ise', 'guncelle'))
        ttk.Combobox(cerceve, textvariable=self.mevcut_ise, values=['guncelle', 'atla'], state='readonly',
                     width=12).grid(row=satir + 1, column=1, sticky='w', padx=8, pady=4)
        self.kuru = tk.BooleanVar(value=True)
        tk.Checkbutton(cerceve, text='Kuru calistir (gondermeden rapor uret)', variable=self.kuru,
                       bg='#f0f0f0').grid(row=satir + 2, column=1, sticky='w', padx=8)
        self.apikey_hatirla = tk.BooleanVar(value=kayitli.get('apikey_hatirla', False))
        tk.Checkbutton(cerceve, text="API Key'i bu klasorde hatirla", variable=self.apikey_hatirla,
                       bg='#f0f0f0').grid(row=satir + 3, column=1, sticky='w', padx=8)

        btn = tk.Frame(self.pencere, bg='#f0f0f0', pady=8)
        btn.pack(fill=tk.X, padx=10)
        tk.Button(btn, text='Baglanti Testi', command=self.test, bg='#7f8c8d', fg='white',
                  padx=12).pack(side=tk.LEFT, padx=3)
        tk.Button(btn, text='▶ GONDER', command=self.gonder, bg='#27ae60', fg='white',
                  font=('Arial', 11, 'bold'), padx=18).pack(side=tk.LEFT, padx=3)
        tk.Button(btn, text='Alan Eslemesi', command=self.eslesme_ac, bg='#8e44ad', fg='white',
                  padx=12).pack(side=tk.LEFT, padx=3)
        tk.Button(btn, text='Kapat', command=self.pencere.destroy, bg='#95a5a6', fg='white',
                  padx=12).pack(side=tk.RIGHT, padx=3)
        self.log_alani()

    def _formu_al(self):
        """Tk degiskenleri sadece ana thread'de okunabilir."""
        d = self.degerler()
        eksik = [k for k in ('sunucu_kodu', 'kullanici', 'sifre', 'apikey', 'firma_kodu', 'donem_kodu')
                 if not d[k]]
        if eksik:
            raise ValueError('Eksik bilgi: ' + ', '.join(eksik))
        kayit = dict(d, modul=self.modul.get(), mevcut_ise=self.mevcut_ise.get(),
                     apikey_hatirla=self.apikey_hatirla.get())
        if not self.apikey_hatirla.get():
            kayit.pop('apikey')
        ayar_yaz(self.klasor, 'dia', kayit)
        return kayit if self.apikey_hatirla.get() else dict(kayit, apikey=d['apikey'])

    @staticmethod
    def _istemci(d):
        from dia_ws import DiaIstemci
        return DiaIstemci(d['sunucu_kodu'], d['kullanici'], d['sifre'], d['apikey'],
                          d['firma_kodu'], d['donem_kodu'])

    def test(self):
        d = self._formu_al_guvenli()
        if d is None:
            return

        def isi():
            ist = self._istemci(d)
            self.log('Giris yapiliyor: ' + ist.temel_url)
            ist.login()
            self.log('Giris OK.')
            try:
                self.log('Kalan kontor: ' + str(ist.kontor_sorgula()))
            except Exception as e:
                self.log('Kontor sorgulanamadi: ' + str(e))
            ist.logout()
        self.arka_planda(isi)

    def eslesme_ac(self):
        from dia_aktar import eslesme_yukle, ESLESME_DOSYASI
        eslesme_yukle(self.klasor)
        yol = os.path.join(self.klasor, ESLESME_DOSYASI)
        try:
            os.startfile(yol)
        except (AttributeError, OSError):
            messagebox.showinfo('Alan Eslemesi', 'Dosya: ' + yol, parent=self.pencere)

    def gonder(self):
        d = self._formu_al_guvenli()
        if d is None:
            return
        kuru = self.kuru.get()
        modul_adi = self.modul.get()
        mevcut_ise = self.mevcut_ise.get()
        if not kuru and not messagebox.askyesno(
                'Onay', modul_adi.upper() + " kartlari DIA'ya GERCEKTEN gonderilecek.\n"
                "Once kuru calistirma raporunu kontrol ettiniz mi?\n\nDevam edilsin mi?", parent=self.pencere):
            return

        def isi():
            from dia_aktar import dosyadan_aktar, eslesme_yukle
            ist = self._istemci(d)
            kurallar = kurallari_getir(self.klasor)
            eslesme = eslesme_yukle(self.klasor)
            eslesme[modul_adi]['mevcut_ise'] = mevcut_ise
            with open(os.path.join(self.klasor, 'dia_eslesme.json'), 'w', encoding='utf-8') as f:
                json.dump(eslesme, f, ensure_ascii=False, indent=2)
            self.log(('[KURU] ' if kuru else '') + modul_adi.upper() + " DIA'ya aktariliyor...")
            try:
                dosyadan_aktar(ist, self.klasor, modul_adi, kurallar, kuru, self.log)
            finally:
                ist.logout()
            self.log('Bitti. Cagri sayisi: ' + str(ist.cagri_sayisi))
        self.arka_planda(isi)
