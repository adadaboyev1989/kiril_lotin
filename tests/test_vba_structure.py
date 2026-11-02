# -*- coding: utf-8 -*-
"""VBA modullarining tuzilishini Office'siz tekshiradi.

Bu xatolar VBA da butun modulni kompilyatsiya qilib bo'lmaydigan qiladi
(qo'shimcha umuman ishlamay qoladi), shuning uchun reliz oldidan ushlanishi shart:

* Declare / Const / modul o'zgaruvchilari birinchi Sub/Function dan keyin
  ("Only comments may appear after End Sub, End Function, or End Property");
* bitta buyruqda 24 tadan ortiq qator davomi (" _");
* yopilmagan Sub/Function/Property yoki #If bloklari;
* ASCII bo'lmagan belgilar (VBA muharriri UTF-8 ni tushunmaydi) va CRLF bo'lmagan qatorlar.
"""
import glob
import os
import re
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FILES = sorted(glob.glob(os.path.join(ROOT, "src", "*.bas")) + [os.path.join(ROOT, "src", "ThisWorkbook.vba")])

PROC_START = re.compile(r"^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+\w+", re.I)
PROC_END = re.compile(r"^\s*End\s+(Sub|Function|Property)\b", re.I)


def logical_lines(text):
    """(boshlanish qatori raqami, qator davomlari soni, birlashtirilgan matn)"""
    lines = text.split("\r\n")
    i = 0
    while i < len(lines):
        start, parts, cont = i + 1, [lines[i]], 0
        while parts[-1].rstrip().endswith(" _") and not parts[-1].lstrip().startswith("'"):
            i += 1
            cont += 1
            parts.append(lines[i] if i < len(lines) else "")
        yield start, cont, " ".join(p.rstrip()[:-1] if p.rstrip().endswith(" _") else p for p in parts)
        i += 1


def strip_comment(line):
    out, in_str = [], False
    for ch in line:
        if ch == '"':
            in_str = not in_str
        elif ch == "'" and not in_str:
            break
        out.append(ch)
    return "".join(out).strip()


class TestVbaStructure(unittest.TestCase):
    def test_files(self):
        self.assertTrue(FILES)
        for path in FILES:
            with self.subTest(file=os.path.basename(path)):
                self.check(path)

    def check(self, path):
        name = os.path.basename(path)
        with open(path, "rb") as f:
            raw = f.read()
        raw.decode("ascii")                                     # faqat ASCII
        self.assertNotIn(b"\n", raw.replace(b"\r\n", b""), "%s: CRLF bo'lmagan qator" % name)
        text = raw.decode("ascii")

        in_proc, seen_proc, if_depth = False, False, 0
        for no, cont, line in logical_lines(text):
            self.assertLessEqual(cont, 24, "%s:%d: 24 tadan ortiq qator davomi" % (name, no))
            code = strip_comment(line)
            if not code or code.lower().startswith("attribute vb_"):
                continue
            low = code.lower()
            if low.startswith("#if "):
                if_depth += 1
            elif low.startswith("#end if"):
                if_depth -= 1
            if PROC_END.match(code):
                self.assertTrue(in_proc, "%s:%d: ortiqcha End" % (name, no))
                in_proc = False
                continue
            if PROC_START.match(code) and not low.startswith(("private declare", "public declare", "declare")):
                self.assertFalse(in_proc, "%s:%d: protsedura yopilmagan" % (name, no))
                in_proc, seen_proc = True, True
                continue
            if not in_proc and seen_proc:
                self.fail("%s:%d: birinchi protseduradan keyin modul darajasida kod bor "
                          "(Declare/Const/o'zgaruvchilar modul boshida bo'lishi kerak): %s" % (name, no, code))
        self.assertFalse(in_proc, "%s: oxirgi protsedura yopilmagan" % name)
        self.assertEqual(if_depth, 0, "%s: #If/#End If muvozanatda emas" % name)

    def test_detects_misplaced_declare(self):
        """Test o'zi ishlayaptimi: v1.5.1 dagi xato albatta ushlanishi kerak."""
        bad = "Option Explicit\r\nSub A()\r\nEnd Sub\r\n#If VBA7 Then\r\nPrivate Declare PtrSafe Function F Lib \"x\" () As Long\r\n#End If\r\n"
        tmp = os.path.join(ROOT, "tests", "_bad_module.bas")
        with open(tmp, "wb") as f:
            f.write(bad.encode("ascii"))
        try:
            with self.assertRaises(AssertionError):
                self.check(tmp)
        finally:
            os.remove(tmp)


if __name__ == "__main__":
    unittest.main()
