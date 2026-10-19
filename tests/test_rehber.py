"""Kullanim rehberi testleri."""
import os
import sys
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

import rehber  # noqa: E402
from _ortak import AYAR_TANIMLARI  # noqa: E402

KOK = os.path.join(os.path.dirname(__file__), '..')


class RehberTest(unittest.TestCase):
    def test_her_ayar_rehberde(self):
        basliklar = {icerik for b in rehber.bolumler() for tur, icerik in b['parcalar'] if tur == 'h'}
        eksik = [a['ayar'] for a in AYAR_TANIMLARI if a['ayar'] not in basliklar]
        self.assertEqual(eksik, [])

    def test_bmp_disi_karakter_yok(self):
        # Tk 8.6 metin kutusunda emoji (U+10000 ustu) indeksleri kaydirir
        metin = rehber.markdown()
        self.assertEqual([c for c in metin if ord(c) > 0xFFFF], [])

    def test_html(self):
        sayfa = rehber.html_sayfa()
        self.assertIn('<h2 id="b0">', sayfa)
        self.assertIn('birim_eslesme', sayfa)

    def test_docs_guncel(self):
        with open(os.path.join(KOK, 'docs', 'KULLANIM_REHBERI.md'), encoding='utf-8') as f:
            self.assertEqual(f.read(), rehber.markdown(),
                             'docs/KULLANIM_REHBERI.md eski: python src/rehber.py calistirin')


if __name__ == '__main__':
    unittest.main()
