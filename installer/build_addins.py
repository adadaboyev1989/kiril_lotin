#!/usr/bin/env python3
"""
Word (.dotm) va Excel (.xlam) qo'shimchalarini src/ dagi VBA manbalaridan
Office'siz yasaydi.

    python3 installer/build_addins.py          ->  addins/KirillLotin.dotm
                                                    addins/KirillLotin.xlam

Qo'shimcha kutubxona kerak emas (faqat Python 3 standart kutubxonasi).

Qanday ishlaydi
---------------
Office fayli - ZIP arxiv. Makroslar uning ichidagi vbaProject.bin
faylida saqlanadi. vbaProject.bin - OLE (Compound File, [MS-CFB]) fayli
bo'lib, ichida VBA loyihasi [MS-OVBA] formatida yoziladi:

    PROJECT          - loyiha xossalari (matn)
    PROJECTwm        - modul nomlari (Unicode)
    VBA/dir          - loyiha va modullar ro'yxati (siqilgan)
    VBA/_VBA_PROJECT - versiya: 0xFFFF (kompilyatsiya keshi yo'q)
    VBA/<modul>      - modul matni (siqilgan)

_VBA_PROJECT versiyasi 0xFFFF bo'lgani uchun Office faylni birinchi marta
ochganda VBA kodini manba matnidan o'zi kompilyatsiya qiladi. Shuning
uchun fayl Office'ning 32 va 64 bitli barcha versiyalarida ishlaydi.
"""
import os
import random
import struct
import sys
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "src")
OUT = os.path.join(ROOT, "addins")

CODEPAGE = 1252
CP = "cp1252"
LCID = 0x0409
PROJECT_NAME = "KirillLotin"
PROJECT_ID = "{6B1E6F3A-3C2D-4B7E-9A51-2F0D8C4E7A10}"

STDOLE_LIBID = ("*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#"
                "C:\\Windows\\System32\\stdole2.tlb#OLE Automation")


# =====================================================================
#  [MS-OVBA] 2.4.1  Siqish (compression)
# =====================================================================
def _compress_chunk(data):
    out = bytearray()
    pos, n = 0, len(data)
    while pos < n:
        flag_pos = len(out)
        out.append(0)
        flags = 0
        for bit in range(8):
            if pos >= n:
                break
            bit_count = 4
            while (1 << bit_count) < pos:
                bit_count += 1
            max_len = (0xFFFF >> bit_count) + 3
            best_len, best_off = 0, 0
            for cand in range(pos - 1, -1, -1):
                length = 0
                while (pos + length < n and length < max_len
                       and data[cand + length] == data[pos + length]):
                    length += 1
                if length > best_len:
                    best_len, best_off = length, pos - cand
                    if length == max_len:
                        break
            if best_len >= 3:
                token = ((best_off - 1) << (16 - bit_count)) | (best_len - 3)
                out += struct.pack("<H", token)
                flags |= 1 << bit
                pos += best_len
            else:
                out.append(data[pos])
                pos += 1
        out[flag_pos] = flags
    return bytes(out)


def ovba_compress(data):
    result = bytearray(b"\x01")
    for start in range(0, len(data), 4096):
        chunk = data[start:start + 4096]
        body = _compress_chunk(chunk)
        if len(body) > 4096 and len(chunk) == 4096:
            header = 0x3000 | (4096 + 2 - 3)            # siqilmagan blok
            result += struct.pack("<H", header) + chunk
        else:
            assert len(body) <= 4096, "siqilgan blok juda katta"
            header = 0x8000 | 0x3000 | (len(body) + 2 - 3)
            result += struct.pack("<H", header) + body
    return bytes(result)


def ovba_decompress(data):
    """Tekshirish uchun: [MS-OVBA] 2.4.1.3.1"""
    assert data[0] == 1
    out = bytearray()
    pos = 1
    while pos < len(data):
        header = struct.unpack_from("<H", data, pos)[0]
        size = (header & 0x0FFF) + 3
        compressed = header & 0x8000
        chunk_end = pos + size
        pos += 2
        start = len(out)
        if not compressed:
            out += data[pos:pos + 4096]
            pos += 4096
            continue
        while pos < chunk_end:
            flags = data[pos]
            pos += 1
            for bit in range(8):
                if pos >= chunk_end:
                    break
                if flags & (1 << bit):
                    token = struct.unpack_from("<H", data, pos)[0]
                    pos += 2
                    diff = len(out) - start
                    bit_count = 4
                    while (1 << bit_count) < diff:
                        bit_count += 1
                    length = (token & (0xFFFF >> bit_count)) + 3
                    offset = (token >> (16 - bit_count)) + 1
                    for _ in range(length):
                        out.append(out[-offset])
                else:
                    out.append(data[pos])
                    pos += 1
    return bytes(out)


# =====================================================================
#  [MS-OVBA] 2.4.3  PROJECT oqimidagi CMG / DPB / GC shifrlash
# =====================================================================
def ovba_encrypt(project_id, data, rng):
    seed = rng.randrange(256)
    version = 2
    version_enc = seed ^ version
    proj_key = sum(ord(c) for c in project_id) & 0xFF
    proj_key_enc = seed ^ proj_key
    unenc1, enc1, enc2 = proj_key, proj_key_enc, version_enc
    out = bytearray([seed, version_enc, proj_key_enc])
    for _ in range((seed & 6) // 2):
        temp = rng.randrange(256)
        b = ((enc2 + unenc1) & 0xFF) ^ temp
        out.append(b)
        enc2, enc1, unenc1 = enc1, b, temp
    for byte in struct.pack("<I", len(data)) + data:
        b = ((enc2 + unenc1) & 0xFF) ^ byte
        out.append(b)
        enc2, enc1, unenc1 = enc1, b, byte
    return out.hex().upper()


# =====================================================================
#  [MS-OVBA] 2.3.4.2  dir oqimi
# =====================================================================
def _rec(rec_id, data):
    return struct.pack("<HI", rec_id, len(data)) + data


def _mbcs(s):
    return s.encode(CP)


def _utf16(s):
    return s.encode("utf-16-le")


def build_dir_stream(modules):
    d = bytearray()
    d += _rec(0x0001, struct.pack("<I", 1))                 # SYSKIND: Win32
    d += _rec(0x0002, struct.pack("<I", LCID))              # LCID
    d += _rec(0x0014, struct.pack("<I", LCID))              # LCIDINVOKE
    d += _rec(0x0003, struct.pack("<H", CODEPAGE))          # CODEPAGE
    d += _rec(0x0004, _mbcs(PROJECT_NAME))                  # NAME
    d += _rec(0x0005, b"") + _rec(0x0040, b"")              # DOCSTRING
    d += _rec(0x0006, b"") + _rec(0x003D, b"")              # HELPFILEPATH
    d += _rec(0x0007, struct.pack("<I", 0))                 # HELPCONTEXT
    d += _rec(0x0008, struct.pack("<I", 0))                 # LIBFLAGS
    d += struct.pack("<HIIH", 0x0009, 4, 1, 0)              # VERSION 1.0
    d += _rec(0x000C, b"") + _rec(0x003C, b"")              # CONSTANTS

    # Havola: stdole (OLE Automation)
    d += _rec(0x0016, _mbcs("stdole")) + _rec(0x003E, _utf16("stdole"))
    libid = _mbcs(STDOLE_LIBID)
    body = struct.pack("<I", len(libid)) + libid + struct.pack("<IH", 0, 0)
    d += _rec(0x000D, body)

    d += struct.pack("<HIH", 0x000F, 2, len(modules))       # MODULES count
    d += struct.pack("<HIH", 0x0013, 2, 0xFFFF)             # PROJECTCOOKIE
    for name, is_doc, _ in modules:
        d += _rec(0x0019, _mbcs(name)) + _rec(0x0047, _utf16(name))
        d += _rec(0x001A, _mbcs(name)) + _rec(0x0032, _utf16(name))
        d += _rec(0x001C, b"") + _rec(0x0048, b"")
        d += _rec(0x0031, struct.pack("<I", 0))             # MODULEOFFSET
        d += _rec(0x001E, struct.pack("<I", 0))             # HELPCONTEXT
        d += _rec(0x002C, struct.pack("<H", 0xFFFF))        # MODULECOOKIE
        d += struct.pack("<HI", 0x0022 if is_doc else 0x0021, 0)
        d += struct.pack("<HI", 0x002B, 0)                  # terminator
    d += struct.pack("<HI", 0x0010, 0)                      # dir terminator
    return bytes(d)


def build_project_stream(modules, rng):
    lines = ['ID="%s"' % PROJECT_ID]
    for name, is_doc, _ in modules:
        lines.append(("Document=%s/&H00000000" if is_doc else "Module=%s") % name)
    lines += [
        'Name="%s"' % PROJECT_NAME,
        'HelpContextID="0"',
        'VersionCompatible32="393222000"',
        'CMG="%s"' % ovba_encrypt(PROJECT_ID, b"\x00\x00\x00\x00", rng),
        'DPB="%s"' % ovba_encrypt(PROJECT_ID, b"\x00", rng),
        'GC="%s"' % ovba_encrypt(PROJECT_ID, b"\xFF", rng),
        "",
        "[Host Extender Info]",
        "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000",
        "",
        "[Workspace]",
    ]
    for name, _, _ in modules:
        lines.append("%s=0, 0, 0, 0, C" % name)
    return ("\r\n".join(lines) + "\r\n").encode(CP)


def build_projectwm(modules):
    out = bytearray()
    for name, _, _ in modules:
        out += _mbcs(name) + b"\x00" + _utf16(name) + b"\x00\x00"
    return bytes(out + b"\x00\x00")


# =====================================================================
#  [MS-CFB]  Compound File yozuvchi (versiya 3, 512 baytli sektorlar)
# =====================================================================
ENDOFCHAIN, FREESECT, FATSECT, NOSTREAM = 0xFFFFFFFE, 0xFFFFFFFF, 0xFFFFFFFD, 0xFFFFFFFF
SECTOR, MINI, CUTOFF = 512, 64, 4096


class _Entry:
    def __init__(self, name, kind, data=b""):
        self.name, self.kind, self.data = name, kind, data   # kind: 1 storage, 2 stream, 5 root
        self.children = []
        self.left = self.right = self.child = NOSTREAM
        self.color = 1
        self.start, self.size = ENDOFCHAIN, 0
        self.sid = None


def _cfb_key(e):
    return (len(e.name), e.name.upper())


def _build_tree(children):
    """Muvozanatli ikkilik daraxt; oxirgi to'lmagan qatlam qizil."""
    items = sorted(children, key=_cfb_key)
    depth_of = {}

    def build(lo, hi, depth):
        if lo >= hi:
            return NOSTREAM
        mid = (lo + hi) // 2
        e = items[mid]
        depth_of[id(e)] = depth
        e.left = build(lo, mid, depth + 1)
        e.right = build(mid + 1, hi, depth + 1)
        return e.sid

    root = build(0, len(items), 0)
    if items:
        max_depth = max(depth_of.values())
        full = (1 << (max_depth + 1)) - 1 == len(items)
        for e in items:
            e.color = 0 if (depth_of[id(e)] == max_depth and not full and max_depth > 0) else 1
    return root


def write_cfb(root_children):
    """root_children: [(nom, bytes) | (nom, [ichki...])]"""
    root = _Entry("Root Entry", 5)
    entries = [root]

    def add(parent, spec):
        for name, content in spec:
            if isinstance(content, list):
                e = _Entry(name, 1)
                e.sid = len(entries)
                entries.append(e)
                parent.children.append(e)
                add(e, content)
            else:
                e = _Entry(name, 2, content)
                e.sid = len(entries)
                entries.append(e)
                parent.children.append(e)
    root.sid = 0
    add(root, root_children)
    for e in entries:
        if e.children:
            e.child = _build_tree(e.children)
    root.color = 1

    # Kichik oqimlar mini-oqimga, kattalari oddiy sektorlarga
    mini_data = bytearray()
    minifat = []
    big = []
    for e in entries:
        if e.kind != 2:
            continue
        e.size = len(e.data)
        if e.size == 0:
            e.start = ENDOFCHAIN
        elif e.size < CUTOFF:
            first = len(mini_data) // MINI
            count = (e.size + MINI - 1) // MINI
            e.start = first
            minifat += [first + i + 1 for i in range(count - 1)] + [ENDOFCHAIN]
            mini_data += e.data + b"\x00" * (count * MINI - e.size)
        else:
            big.append(e)

    sectors = []   # har biri 512 bayt
    fat = []

    def alloc(data):
        if not data:
            return ENDOFCHAIN
        first = len(sectors)
        count = (len(data) + SECTOR - 1) // SECTOR
        padded = data + b"\x00" * (count * SECTOR - len(data))
        for i in range(count):
            sectors.append(padded[i * SECTOR:(i + 1) * SECTOR])
            fat.append(first + i + 1 if i < count - 1 else ENDOFCHAIN)
        return first

    for e in big:
        e.start = alloc(e.data)
    root.start = alloc(bytes(mini_data))
    root.size = len(mini_data)
    if not mini_data:
        root.start = ENDOFCHAIN

    minifat_bytes = b"".join(struct.pack("<I", v) for v in minifat)
    first_minifat = alloc(minifat_bytes) if minifat else ENDOFCHAIN
    n_minifat = (len(minifat_bytes) + SECTOR - 1) // SECTOR

    # Katalog
    dir_bytes = bytearray()
    for e in entries:
        name16 = _utf16(e.name) + b"\x00\x00"
        assert len(name16) <= 64
        dir_bytes += name16 + b"\x00" * (64 - len(name16))
        dir_bytes += struct.pack("<HBB", len(name16), e.kind, e.color)
        dir_bytes += struct.pack("<III", e.left, e.right, e.child)
        dir_bytes += b"\x00" * 16 + struct.pack("<I", 0)          # CLSID, state
        dir_bytes += b"\x00" * 16                                 # vaqtlar
        start = e.start if e.kind in (2, 5) else 0
        size = e.size if e.kind in (2, 5) else 0
        dir_bytes += struct.pack("<IQ", start, size)
    while len(dir_bytes) % SECTOR:
        free = b"\x00" * 64 + struct.pack("<HBB", 0, 0, 0) + struct.pack("<III", NOSTREAM, NOSTREAM, NOSTREAM)
        dir_bytes += free + b"\x00" * (128 - len(free))
    first_dir = alloc(bytes(dir_bytes))

    # FAT sektorlari (o'zlarini ham hisobga olgan holda)
    n_fat = 1
    while (len(fat) + n_fat) > n_fat * (SECTOR // 4):
        n_fat += 1
    assert n_fat <= 109
    fat_start = len(fat)
    fat += [FATSECT] * n_fat
    fat += [FREESECT] * (n_fat * (SECTOR // 4) - len(fat))
    fat_bytes = b"".join(struct.pack("<I", v) for v in fat)
    for i in range(n_fat):
        sectors.append(fat_bytes[i * SECTOR:(i + 1) * SECTOR])

    header = bytearray()
    header += bytes.fromhex("D0CF11E0A1B11AE1") + b"\x00" * 16
    header += struct.pack("<HHHHH", 0x003E, 0x0003, 0xFFFE, 9, 6)
    header += b"\x00" * 6
    header += struct.pack("<IIIIIIIII", 0, n_fat, first_dir, 0, CUTOFF,
                          first_minifat, n_minifat, ENDOFCHAIN, 0)
    difat = [fat_start + i for i in range(n_fat)] + [FREESECT] * (109 - n_fat)
    header += b"".join(struct.pack("<I", v) for v in difat)
    assert len(header) == 512
    return bytes(header) + b"".join(sectors)


# =====================================================================
#  VBA loyihasi
# =====================================================================
DOC_ATTRS = (
    'Attribute VB_Name = "{name}"\r\n'
    'Attribute VB_Base = "0{{{guid}}}"\r\n'
    "Attribute VB_GlobalNameSpace = False\r\n"
    "Attribute VB_Creatable = False\r\n"
    "Attribute VB_PredeclaredId = True\r\n"
    "Attribute VB_Exposed = True\r\n"
    "Attribute VB_TemplateDerived = False\r\n"
    "Attribute VB_Customizable = True\r\n"
)

GUID_WORD_DOCUMENT = "00020906-0000-0000-C000-000000000046"
GUID_EXCEL_WORKBOOK = "00020819-0000-0000-C000-000000000046"
GUID_EXCEL_WORKSHEET = "00020820-0000-0000-C000-000000000046"


def read_src(name):
    with open(os.path.join(SRC, name), "rb") as f:
        data = f.read()
    data.decode("ascii")                        # faqat ASCII bo'lishi shart
    data = data.replace(b"\r\n", b"\n").replace(b"\n", b"\r\n")
    return data


def doc_module(name, guid, code=b""):
    return (name, True, DOC_ATTRS.format(name=name, guid=guid).encode(CP) + code)


def std_module(file_name):
    code = read_src(file_name)
    name = os.path.splitext(file_name)[0]
    assert code.startswith(('Attribute VB_Name = "%s"' % name).encode()), file_name
    return (name, False, code)


def build_vba_project(modules, seed):
    rng = random.Random(seed)                   # takrorlanuvchan natija
    vba = []
    for name, _, code in modules:
        vba.append((name, ovba_compress(code)))
    vba.append(("dir", ovba_compress(build_dir_stream(modules))))
    vba.append(("_VBA_PROJECT", b"\xCC\x61\xFF\xFF\x00\x00\x00"))
    return write_cfb([
        ("VBA", vba),
        ("PROJECT", build_project_stream(modules, rng)),
        ("PROJECTwm", build_projectwm(modules)),
    ])


# =====================================================================
#  Office Open XML paketlari
# =====================================================================
XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
R_OFFICEDOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
R_UI = "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"
R_VBA = "http://schemas.microsoft.com/office/2006/relationships/vbaProject"
CT_VBA = "application/vnd.ms-office.vbaProject"


def rels(items):
    body = "".join('<Relationship Id="%s" Type="%s" Target="%s"/>' % it for it in items)
    return XML_HEAD + '<Relationships xmlns="%s">%s</Relationships>' % (REL_NS, body)


def content_types(overrides):
    body = ('<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Default Extension="bin" ContentType="%s"/>' % CT_VBA)
    body += "".join('<Override PartName="%s" ContentType="%s"/>' % o for o in overrides)
    return XML_HEAD + ('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                       '%s</Types>' % body)


def write_package(path, parts):
    fixed = (2024, 1, 1, 0, 0, 0)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts:
            if isinstance(data, str):
                data = data.encode("utf-8")
            info = zipfile.ZipInfo(name, fixed)
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, data)


def custom_ui():
    with open(os.path.join(SRC, "customUI.xml"), "rb") as f:
        return f.read()


def build_word(path):
    modules = [
        doc_module("ThisDocument", GUID_WORD_DOCUMENT),
        std_module("KLCore.bas"),
        std_module("KLWord.bas"),
    ]
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    document = XML_HEAD + (
        '<w:document xmlns:w="%s"><w:body><w:p/>'
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1701" '
        'w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>'
        '</w:body></w:document>' % W)
    vba_data = XML_HEAD + '<wne:vbaSuppData xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"/>'
    write_package(path, [
        ("[Content_Types].xml", content_types([
            ("/word/document.xml", "application/vnd.ms-word.template.macroEnabledTemplate.main+xml"),
            ("/word/vbaData.xml", "application/vnd.ms-word.vbaData+xml"),
        ])),
        ("_rels/.rels", rels([
            ("rId1", R_OFFICEDOC, "word/document.xml"),
            ("rIdKL", R_UI, "customUI/customUI.xml"),
        ])),
        ("word/document.xml", document),
        ("word/_rels/document.xml.rels", rels([("rId1", R_VBA, "vbaProject.bin")])),
        ("word/vbaProject.bin", build_vba_project(modules, seed=1)),
        ("word/_rels/vbaProject.bin.rels", rels([
            ("rId1", "http://schemas.microsoft.com/office/2006/relationships/wordVbaData", "vbaData.xml"),
        ])),
        ("word/vbaData.xml", vba_data),
        ("customUI/customUI.xml", custom_ui()),
    ])
    return modules


def build_excel(path):
    modules = [
        doc_module("ThisWorkbook", GUID_EXCEL_WORKBOOK, read_src("ThisWorkbook.vba")),
        doc_module("Sheet1", GUID_EXCEL_WORKSHEET),
        std_module("KLCore.bas"),
        std_module("KLExcel.bas"),
    ]
    S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    workbook = XML_HEAD + (
        '<workbook xmlns="%s" xmlns:r="%s"><workbookPr codeName="ThisWorkbook"/>'
        '<bookViews><workbookView/></bookViews>'
        '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>' % (S, R))
    sheet = XML_HEAD + '<worksheet xmlns="%s"><sheetPr codeName="Sheet1"/><sheetData/></worksheet>' % S
    styles = XML_HEAD + (
        '<styleSheet xmlns="%s">'
        '<fonts count="1"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts>'
        '<fills count="2"><fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="gray125"/></fill></fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        '</styleSheet>' % S)
    SS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    write_package(path, [
        ("[Content_Types].xml", content_types([
            ("/xl/workbook.xml", "application/vnd.ms-excel.addin.macroEnabled.main+xml"),
            ("/xl/worksheets/sheet1.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"),
            ("/xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"),
        ])),
        ("_rels/.rels", rels([
            ("rId1", R_OFFICEDOC, "xl/workbook.xml"),
            ("rIdKL", R_UI, "customUI/customUI.xml"),
        ])),
        ("xl/workbook.xml", workbook),
        ("xl/_rels/workbook.xml.rels", rels([
            ("rId1", SS + "/worksheet", "worksheets/sheet1.xml"),
            ("rId2", SS + "/styles", "styles.xml"),
            ("rId3", R_VBA, "vbaProject.bin"),
        ])),
        ("xl/worksheets/sheet1.xml", sheet),
        ("xl/styles.xml", styles),
        ("xl/vbaProject.bin", build_vba_project(modules, seed=2)),
        ("customUI/customUI.xml", custom_ui()),
    ])
    return modules


def main():
    os.makedirs(OUT, exist_ok=True)
    word = os.path.join(OUT, "KirillLotin.dotm")
    excel = os.path.join(OUT, "KirillLotin.xlam")
    build_word(word)
    build_excel(excel)
    for p in (word, excel):
        print("yasaldi: %s (%d bayt)" % (os.path.relpath(p, ROOT), os.path.getsize(p)))


if __name__ == "__main__":
    sys.exit(main())
