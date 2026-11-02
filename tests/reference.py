# -*- coding: utf-8 -*-
"""Python reference implementation of the transliteration rules.

It mirrors src/KLCore.bas line by line so that the rules can be tested
without Microsoft Office. If you change a rule here, change it in
KLCore.bas as well (and vice versa).
"""

CYR_LOWER = "абвгдеёжзийклмнопрстуфхцчшщъыьэюяўқғҳ"
CYR_UPPER = "АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯЎҚҒҲ"
CYR_VOWELS = "аеёиоуэюяўы"

OKINA = "ʻ"   # ʻ  (o‘, g‘)
TUTUQ = "ʼ"   # ʼ  (tutuq belgisi)
APOSTROPHES = "'`‘’ʻʼ´"

# Yangi lotin alifbosi (2021): sh -> ş, ch -> ç, g' -> ğ, o' -> ö
NEW_LOWER = "\u015f\u00e7\u011f\u00f6"   # ş ç ğ ö
NEW_UPPER = "\u015e\u00c7\u011e\u00d6"   # Ş Ç Ğ Ö

LAT_VOWELS = "aeiou\u00f6"


def lat_upper(t):
    """ş ç ğ ö ni ham bosh harfga aylantiradi (VBA: LatUpper)."""
    return "".join(NEW_UPPER[NEW_LOWER.index(c)] if c in NEW_LOWER else c.upper() for c in t)


def lat_lower(c):
    return NEW_LOWER[NEW_UPPER.index(c)] if c in NEW_UPPER else c.lower()


def cyr_base(lower, ascii_apos, new_alphabet=True):
    ok = "'" if ascii_apos else OKINA
    if new_alphabet:
        special = {"ч": "\u00e7", "ш": "\u015f", "щ": "\u015f", "ў": "\u00f6", "ғ": "\u011f"}
        if lower in special:
            return special[lower]
    table = {
        "а": "a", "б": "b", "в": "v", "г": "g", "д": "d", "ё": "yo",
        "ж": "j", "з": "z", "и": "i", "й": "y", "к": "k", "л": "l",
        "м": "m", "н": "n", "о": "o", "п": "p", "р": "r", "с": "s",
        "т": "t", "у": "u", "ф": "f", "х": "x", "ч": "ch", "ш": "sh",
        "щ": "sh", "ы": "i", "ь": "", "э": "e", "ю": "yu", "я": "ya",
        "ў": "o" + ok, "қ": "q", "ғ": "g" + ok, "ҳ": "h",
    }
    return table[lower]


def cyr_lower(ch):
    p = CYR_UPPER.find(ch)
    return CYR_LOWER[p] if p >= 0 else ch


def is_cyr(ch):
    return ch != "" and (ch in CYR_LOWER or ch in CYR_UPPER)


def is_cyr_upper(ch):
    return ch != "" and ch in CYR_UPPER


def cyr_to_lat(s, ascii_apos=False, new_alphabet=True):
    out = []
    n = len(s)
    for i in range(n):
        c = s[i]
        if not is_cyr(c):
            out.append(c)
            continue
        prev = s[i - 1] if i > 0 else ""
        nxt = s[i + 1] if i + 1 < n else ""
        lc = cyr_lower(c)
        lp = cyr_lower(prev)
        if lc == "е":
            if not is_cyr(prev) or lp in CYR_VOWELS or lp in "ъь":
                t = "ye"
            else:
                t = "e"
        elif lc == "ц":
            t = "ts" if (is_cyr(prev) and lp in CYR_VOWELS) else "s"
        elif lc == "ъ":
            if cyr_lower(nxt) in ("е", "ё", "ю", "я") and nxt != "":
                t = ""
            else:
                t = "'" if ascii_apos else TUTUQ
        else:
            t = cyr_base(lc, ascii_apos, new_alphabet)
        if is_cyr_upper(c) and t:
            if len(t) > 1 and (is_cyr_upper(nxt) or (not is_cyr(nxt) and is_cyr_upper(prev))):
                t = lat_upper(t)
            else:
                t = lat_upper(t[0]) + t[1:]
        out.append(t)
    return "".join(out)


def is_lat(ch):
    return ch != "" and ((("a" <= ch.lower() <= "z") and ch.isascii()) or ch in NEW_LOWER or ch in NEW_UPPER)


def is_lat_upper(ch):
    return ch != "" and (("A" <= ch <= "Z") or ch in NEW_UPPER)


def is_apos(ch):
    return ch != "" and ch in APOSTROPHES


LAT_BASE = {
    "a": "а", "b": "б", "c": "ц", "d": "д", "f": "ф", "g": "г", "h": "ҳ",
    "i": "и", "j": "ж", "k": "к", "l": "л", "m": "м", "n": "н", "o": "о",
    "p": "п", "q": "қ", "r": "р", "s": "с", "t": "т", "u": "у", "v": "в",
    "w": "в", "x": "х", "y": "й", "z": "з",
}


def to_upper_cyr(ch):
    p = CYR_LOWER.find(ch)
    return CYR_UPPER[p] if p >= 0 else ch


def lat_to_cyr_word(s):
    out = []
    n = len(s)
    i = 0
    while i < n:
        c = s[i]
        prev = s[i - 1] if i > 0 else ""
        nxt = s[i + 1] if i + 1 < n else ""
        nxt2 = s[i + 2] if i + 2 < n else ""
        step = 1
        if is_apos(c):
            if is_lat(prev) and is_lat(nxt):
                if prev.lower() == "s" and nxt.lower() == "h":
                    t = ""                      # s'h -> сҳ
                elif is_lat_upper(prev) and is_lat_upper(nxt):
                    t = "Ъ"
                else:
                    t = "ъ"
            else:
                t = c
            out.append(t)
            i += 1
            continue
        if not is_lat(c):
            out.append(c)
            i += 1
            continue
        lc = lat_lower(c)
        ln = lat_lower(nxt)
        if lc in NEW_LOWER:                      # yangi alifbo: ş ç ğ ö
            t = {"\u015f": "ш", "\u00e7": "ч", "\u011f": "ғ", "\u00f6": "ў"}[lc]
        elif lc in "og" and is_apos(nxt):
            t = "ў" if lc == "o" else "ғ"
            step = 2
        elif lc == "s" and ln == "h":
            t = "ш"; step = 2
        elif lc == "c" and ln == "h":
            t = "ч"; step = 2
        elif lc == "y" and ln in ("a", "e", "u", "o") and ln != "" and not (ln == "o" and is_apos(nxt2)):
            t = {"a": "я", "e": "е", "u": "ю", "o": "ё"}[ln]; step = 2
        elif lc == "e":
            if not (is_lat(prev) or is_apos(prev)) or lat_lower(prev) in LAT_VOWELS:
                t = "э"
            else:
                t = "е"
        else:
            t = LAT_BASE[lc]
        if is_lat_upper(c):
            t = to_upper_cyr(t)
        out.append(t)
        i += step
    return "".join(out)


def is_url_like(tok):
    t = tok.lower()
    return "://" in t or "www." in t or "@" in t


def lat_to_cyr(s):
    # Split on whitespace, keep URL/e-mail tokens unchanged.
    out = []
    i = 0
    n = len(s)
    while i < n:
        if s[i].isspace():
            out.append(s[i]); i += 1; continue
        j = i
        while j < n and not s[j].isspace():
            j += 1
        tok = s[i:j]
        out.append(tok if is_url_like(tok) else lat_to_cyr_word(tok))
        i = j
    return "".join(out)


# ---------------------------------------------------------------------
#  Word'dagi ishlash tartibini takrorlash (src/KLWord.bas: ConvertRange)
#  Word matnni so'zma-so'z o'giradi: so'zning birinchi harfi topiladi,
#  keyin so'z tarkibidagi belgilar bo'yicha oxirigacha kengaytiriladi.
# ---------------------------------------------------------------------
WORD_LAT_CHARS = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz" + NEW_LOWER + NEW_UPPER + APOSTROPHES
WORD_CYR_CHARS = CYR_LOWER + CYR_UPPER


def _word_link_token(s, start, end):
    """IsInsideLink: bo'shliqlar bilan chegaralangan butun bo'lak."""
    delims = " \t\r\n\x0b\xa0"
    a = start
    while a > 0 and s[a - 1] not in delims:
        a -= 1
    b = end
    while b < len(s) and s[b] not in delims:
        b += 1
    return s[a:b]


def word_convert(s, to_latin, ascii_apos=False, new_alphabet=True):
    charset = WORD_CYR_CHARS if to_latin else WORD_LAT_CHARS
    out = []
    i, n = 0, len(s)
    while i < n:
        if s[i] not in charset:
            out.append(s[i])
            i += 1
            continue
        j = i
        while j < n and s[j] in charset:        # MoveEndWhile
            j += 1
        word = s[i:j]
        if to_latin:
            out.append(cyr_to_lat(word, ascii_apos, new_alphabet))
        elif is_url_like(_word_link_token(s, i, j)):
            out.append(word)
        else:
            out.append(lat_to_cyr_word(word))
        i = j
    return "".join(out)
