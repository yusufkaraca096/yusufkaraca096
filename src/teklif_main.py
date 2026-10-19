"""RACA_Teklif.exe giris noktasi: sadece teklif motoru (veri temizleme programi olmadan).

--oz-test CIKTI.pdf : arayuz acmadan ornek veriyle PDF + Excel uretir (derlenen EXE'nin CI'da sinanmasi icin).
"""
import sys


def oz_test(pdf_yolu):
    import teklif_cikti
    from teklif_hesap import hesapla
    from teklif_veri import AYAR_VARSAYILAN
    paketler = {'P1': {'kod': 'P1', 'tur': 'Masaüstü', 'ad': 'Örnek Paket Şçğ', 'std_fiyat': 1000, 'ent_fiyat': 0,
                       'kull_yuzde': 10, 'std_saat': 2, 'ent_saat': 3}}
    girdi = {'teklif_no': 'TEST-1', 'musteri': 'Öz Test İşletmesi', 'masaustu': 3, 'mobil': 0,
             'kalemler': [{'kod': 'P1', 'ind_tutar': 10}]}
    sonuc = hesapla(girdi, paketler, dict(AYAR_VARSAYILAN, saatlik_ucret=100))
    teklif_cikti.pdf_olustur(pdf_yolu, girdi, sonuc, AYAR_VARSAYILAN)
    teklif_cikti.excel_olustur(pdf_yolu + '.xlsx', girdi, sonuc, AYAR_VARSAYILAN)


if __name__ == '__main__':
    if len(sys.argv) > 2 and sys.argv[1] == '--oz-test':
        oz_test(sys.argv[2])
    else:
        from teklif_gui import main
        main()
