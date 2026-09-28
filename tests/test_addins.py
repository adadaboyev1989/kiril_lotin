# -*- coding: utf-8 -*-
"""installer/build_addins.py generatorini tekshiradi (Office'siz)."""
import os
import random
import sys
import tempfile
import unittest
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "installer"))

import build_addins as b  # noqa: E402

try:
    import olefile
except ImportError:          # ixtiyoriy: pip install olefile
    olefile = None


class TestCompression(unittest.TestCase):
    def test_round_trip(self):
        rng = random.Random(0)
        samples = [b"", b"a", b"abcabcabcabc" * 50,
                   bytes(rng.randrange(256) for _ in range(9000))]
        for name in ("KLCore.bas", "KLWord.bas", "KLExcel.bas"):
            samples.append(b.read_src(name))
        for data in samples:
            self.assertEqual(b.ovba_decompress(b.ovba_compress(data)), data)


class TestPackages(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.mkdtemp()
        cls.word = os.path.join(cls.tmp, "KirillLotin.dotm")
        cls.excel = os.path.join(cls.tmp, "KirillLotin.xlam")
        cls.word_modules = b.build_word(cls.word)
        cls.excel_modules = b.build_excel(cls.excel)

    def _check(self, path, part, modules):
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
            self.assertIn("customUI/customUI.xml", names)
            self.assertIn("customUI/customUI.xml", z.read("_rels/.rels").decode())
            vba = z.read(part)
        if olefile is None:
            self.skipTest("olefile o'rnatilmagan")
        ole = olefile.OleFileIO(vba, raise_defects=olefile.DEFECT_INCORRECT)
        self.assertEqual(ole.openstream("VBA/_VBA_PROJECT").read(), b"\xCC\x61\xFF\xFF\x00\x00\x00")
        project = ole.openstream("PROJECT").read().decode("cp1252")
        for name, _, code in modules:
            self.assertIn(name, project)
            stream = ole.openstream("VBA/" + name).read()
            self.assertEqual(b.ovba_decompress(stream), code)

    def test_word(self):
        self._check(self.word, "word/vbaProject.bin", self.word_modules)

    def test_excel(self):
        self._check(self.excel, "xl/vbaProject.bin", self.excel_modules)

    def test_reproducible(self):
        other = os.path.join(self.tmp, "again.xlam")
        b.build_excel(other)
        with open(self.excel, "rb") as f1, open(other, "rb") as f2:
            self.assertEqual(f1.read(), f2.read())


if __name__ == "__main__":
    unittest.main()
