"""Logo Tiger / Go veritabanindan (MS SQL) cari ve stok kartlarini okur.

Sadece SELECT calistirir, Logo verisini degistirmez.
Cikti kolon adlari kurallar.py ESLESME basliklarina uygundur; uretilen
cari_listesi.xlsx / stok_listesi.xlsx dogrudan temizleme akisina girer.
Tablo yapisi: docs/ENTEGRASYON.md (LG_<firma>_CLCARD, LG_<firma>_ITEMS ...)
"""
import os
import re

import pandas as pd

SURUCU_ONCELIK = ['ODBC Driver 18 for SQL Server', 'ODBC Driver 17 for SQL Server',
                  'ODBC Driver 13 for SQL Server', 'SQL Server Native Client 11.0', 'SQL Server']

# (Logo alani, cikti basligi, zorunlu mu)  - zorunlu olmayanlar surumde varsa eklenir
CARI_ALANLARI = [
    ('CODE', 'Cari Kodu', True),
    ('DEFINITION_', 'Unvan', True),
    ('TAXNR', 'Vergi No', True),
    ('TAXOFFICE', 'Vergi Dairesi', True),
    ('TCKNO', 'TC Kimlik', False),
    ('ADDR1', 'Adres 1', True),
    ('ADDR2', 'Adres 2', True),
    ('CITY', 'Il', True),
    ('TOWN', 'Ilce', True),
    ('POSTCODE', 'postakodu', True),
    ('COUNTRY', 'Ulke', True),
    ('TELNRS1', 'Telefon 1', True),
    ('TELNRS2', 'Telefon 2', True),
    ('CELLPHONE', 'Cep Telefonu', False),
    ('FAXNR', 'Fax', True),
    ('EMAILADDR', 'Email', True),
    ('WEBADDR', 'Web Adresi', True),
    ('SPECODE', 'Ozel Kod 1', True),
    ('CARDTYPE', 'LOGO_KART_TURU', True),
    ('ISPERSCOMP', 'LOGO_SAHIS_SIRKETI', False),
]


def firma_no_dogrula(firma):
    """'1' -> '001'. Tablo adina girdigi icin sadece rakam kabul edilir."""
    s = str(firma).strip()
    if not re.fullmatch(r'\d{1,3}', s):
        raise ValueError('Firma numarasi 1-3 haneli sayi olmali: ' + s)
    return s.zfill(3)


def donem_no_dogrula(donem):
    s = str(donem).strip()
    if not re.fullmatch(r'\d{1,2}', s):
        raise ValueError('Donem numarasi 1-2 haneli sayi olmali: ' + s)
    return s.zfill(2)


class LogoKaynak:
    """conn: DB-API baglantisi. lehce: 'mssql' (gercek Logo) veya 'sqlite' (testler)."""

    def __init__(self, conn, lehce='mssql'):
        self.conn = conn
        self.lehce = lehce

    # ---------------- baglanti ----------------
    @staticmethod
    def surucu_bul():
        import pyodbc
        mevcut = pyodbc.drivers()
        for s in SURUCU_ONCELIK:
            if s in mevcut:
                return s
        raise RuntimeError('SQL Server ODBC surucusu bulunamadi. "ODBC Driver 17 for SQL Server" kurun. '
                           'Bulunanlar: ' + ', '.join(mevcut))

    @classmethod
    def baglan(cls, sunucu, veritabani, kullanici='', sifre='', zaman_asimi=15):
        """kullanici bos ise Windows kimlik dogrulamasi kullanilir."""
        import pyodbc
        surucu = cls.surucu_bul()
        parcalar = ['DRIVER={' + surucu + '}', 'SERVER=' + sunucu, 'DATABASE=' + veritabani]
        if kullanici:
            parcalar += ['UID=' + kullanici, 'PWD=' + sifre]
        else:
            parcalar.append('Trusted_Connection=yes')
        if surucu.startswith('ODBC Driver 18'):
            parcalar.append('TrustServerCertificate=yes')
        conn = pyodbc.connect(';'.join(parcalar) + ';', timeout=zaman_asimi, readonly=True)
        return cls(conn, 'mssql')

    # ---------------- yardimcilar ----------------
    def _oku(self, sql):
        cur = self.conn.cursor()
        cur.execute(sql)
        kolonlar = [c[0] for c in cur.description]
        satirlar = [tuple(r) for r in cur.fetchall()]
        return pd.DataFrame(satirlar, columns=kolonlar)

    def tablolar(self):
        if self.lehce == 'sqlite':
            sql = "SELECT name FROM sqlite_master WHERE type IN ('table', 'view')"
        else:
            sql = 'SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES'
        return [str(t).upper() for t in self._oku(sql).iloc[:, 0]]

    def kolonlar(self, tablo):
        if self.lehce == 'sqlite':
            return [str(r[1]).upper() for r in self.conn.execute('PRAGMA table_info(' + tablo + ')')]
        df = self._oku("SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '"
                       + tablo + "'")
        return [str(c).upper() for c in df.iloc[:, 0]]

    def _ilk(self, kolon, kaynak, kosul, sirala):
        """Ilk eslesen degeri getiren alt sorgu (TOP 1 / LIMIT 1)."""
        if self.lehce == 'sqlite':
            return ('(SELECT ' + kolon + ' FROM ' + kaynak + ' WHERE ' + kosul +
                    ' ORDER BY ' + sirala + ' LIMIT 1)')
        return '(SELECT TOP 1 ' + kolon + ' FROM ' + kaynak + ' WHERE ' + kosul + ' ORDER BY ' + sirala + ')'

    def firmalar(self):
        """Veritabanindaki Logo firma numaralari (LG_xxx_CLCARD tablolarindan)."""
        nolar = sorted({m.group(1) for t in self.tablolar()
                        for m in [re.fullmatch(r'LG_(\d{3})_CLCARD', t)] if m})
        return nolar

    # ---------------- sorgular ----------------
    def cari_sorgusu(self, firma, sadece_aktif=True):
        f = firma_no_dogrula(firma)
        tablo = 'LG_' + f + '_CLCARD'
        mevcut = set(self.kolonlar(tablo))
        if not mevcut:
            raise RuntimeError(tablo + ' tablosu bulunamadi. Firma numarasini kontrol edin.')
        secim = []
        for alan, baslik, zorunlu in CARI_ALANLARI:
            if alan in mevcut:
                secim.append('C.' + alan + ' AS "' + baslik + '"')
            elif zorunlu:
                raise RuntimeError(tablo + ' tablosunda ' + alan + ' alani yok.')
        sql = 'SELECT ' + ', '.join(secim) + ' FROM ' + tablo + ' C'
        if sadece_aktif:
            sql += ' WHERE C.ACTIVE = 0'
        return sql + ' ORDER BY C.CODE'

    def stok_sorgusu(self, firma, sadece_aktif=True):
        f = firma_no_dogrula(firma)
        p = 'LG_' + f + '_'
        if not self.kolonlar(p + 'ITEMS'):
            raise RuntimeError(p + 'ITEMS tablosu bulunamadi. Firma numarasini kontrol edin.')
        barkod = self._ilk('B.BARCODE', p + 'ITMUNITA B',
                           "B.ITEMREF = I.LOGICALREF AND B.BARCODE <> ''", 'B.LINENR')
        satis = self._ilk('P.PRICE', p + 'PRCLIST P', 'P.CARDREF = I.LOGICALREF AND P.PTYPE = 2', 'P.PRIORITY')
        alis = self._ilk('P.PRICE', p + 'PRCLIST P', 'P.CARDREF = I.LOGICALREF AND P.PTYPE = 1', 'P.PRIORITY')
        sql = ('SELECT I.CODE AS "Stok Kodu", I.NAME AS "Stok Adi", U.CODE AS "Ana Birim", '
               'I.VAT AS "KDV", I.STGRPCODE AS "Grup Kodu", I.SPECODE AS "Ozel Kod 1", '
               + barkod + ' AS "Barkod No", ' + satis + ' AS "Satis Fiyati", '
               + alis + ' AS "Alis Fiyati", I.CARDTYPE AS "LOGO_KART_TURU" '
               'FROM ' + p + 'ITEMS I '
               'LEFT JOIN ' + p + 'UNITSETL U ON U.UNITSETREF = I.UNITSETREF AND U.MAINUNIT = 1 '
               'WHERE I.CARDTYPE NOT IN (20, 21)')
        if sadece_aktif:
            sql += ' AND I.ACTIVE = 0'
        return sql + ' ORDER BY I.CODE'

    def cari_cek(self, firma, sadece_aktif=True):
        df = self._oku(self.cari_sorgusu(firma, sadece_aktif))
        return _metinlestir(df, ['Vergi No', 'TC Kimlik', 'Telefon 1', 'Telefon 2', 'Cep Telefonu',
                                 'Fax', 'postakodu', 'Cari Kodu'])

    def stok_cek(self, firma, sadece_aktif=True):
        df = self._oku(self.stok_sorgusu(firma, sadece_aktif))
        return _metinlestir(df, ['Stok Kodu', 'Barkod No'])


def _metinlestir(df, kolonlar):
    """Logo CHAR alanlari: sag bosluklari al, None -> '' ; VKN/telefon metin kalsin."""
    for k in df.columns:
        df[k] = df[k].map(lambda v: v.strip() if isinstance(v, str) else v)
    for k in kolonlar:
        if k in df.columns:
            df[k] = df[k].map(lambda v: '' if v is None or (not isinstance(v, str) and pd.isna(v)) else str(v))
    return df


def excele_yaz(df, klasor, dosya_adi):
    """Metin kolonlari Excel'de metin hucresi olarak yazilir (bastaki 0 korunur)."""
    yol = os.path.join(klasor, dosya_adi)
    df.to_excel(yol, index=False)
    return yol
