"""src/KLCore.bas ni LibreOffice Basic (VBA moslik rejimi) da haqiqatan
ishga tushirib, natijalarni tests/reference.py (Python nusxasi) bilan
solishtiradi. Office kerak emas, LibreOffice kerak:

    sudo apt install libreoffice-core
    python3 tests/vba_libreoffice.py

Chiqish kodi: 0 - hammasi mos, 1 - farqlar bor, 2 - ishga tushmadi.
"""
import html
import os
import re
import subprocess
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, 'tests'))
import reference as R  # noqa: E402
import test_reference as T  # noqa: E402

W = tempfile.mkdtemp(prefix='kl_vba_')
PROFILE = W + '/profile'
URL = 'file://' + PROFILE

# 1) profil yaratish
if not os.path.isdir(PROFILE + '/user/basic/Standard'):
    subprocess.run(['soffice', '-env:UserInstallation=' + URL, '--headless', '--terminate_after_init'],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=120)

std = PROFILE + '/user/basic/Standard'
core = open(ROOT + '/src/KLCore.bas', encoding='ascii').read().replace('\r\n', '\n')
core = '\n'.join(l for l in core.split('\n') if not l.startswith('Attribute VB_'))
# Windows API (MessageBoxW) LibreOffice'da yo'q - bu qism test paytida olib tashlanadi
core = re.sub(r"' --- WINAPI BEGIN.*?' --- WINAPI END[^\n]*\n", '', core, flags=re.S)
core = re.sub(r"' --- WINAPI-DECL BEGIN.*?' --- WINAPI-DECL END[^\n]*\n", '', core, flags=re.S)
core = 'Option VBASupport 1\n' + core.replace('Option Explicit', '')

harness = r'''Option VBASupport 1
Function GetSetting(a, b, c, d)
    GetSetting = d
End Function
Sub SaveSetting(a, b, c, d)
End Sub
Sub KL_UnicodeMsg(t, c, f)
End Sub
Sub Main
    Dim sfa As Object, inp As Object, outp As Object, ti As Object, tout As Object
    Dim line As String, mode As String, txt As String, res As String, p As Long
    sfa = createUnoService("com.sun.star.ucb.SimpleFileAccess")
    inp = sfa.openFileRead("file://W/cases.txt")
    ti = createUnoService("com.sun.star.io.TextInputStream")
    ti.setInputStream(inp)
    ti.setEncoding("UTF-8")
    If sfa.exists("file://W/out.txt") Then sfa.kill("file://W/out.txt")
    outp = sfa.openFileWrite("file://W/out.txt")
    tout = createUnoService("com.sun.star.io.TextOutputStream")
    tout.setOutputStream(outp)
    tout.setEncoding("UTF-8")
    Do While Not ti.isEOF()
        line = ti.readLine()
        p = InStr(line, Chr(9))
        If p > 0 Then
            mode = Left(line, p - 1)
            txt = Mid(line, p + 1)
            Select Case mode
                Case "cyr_new": KL_SetNewAlphabet True: KL_SetAsciiApostrophe False: res = KL_CyrToLat(txt)
                Case "cyr_old": KL_SetNewAlphabet False: KL_SetAsciiApostrophe False: res = KL_CyrToLat(txt)
                Case "cyr_old_ascii": KL_SetNewAlphabet False: KL_SetAsciiApostrophe True: res = KL_CyrToLat(txt)
                Case "lat": res = KL_LatToCyr(txt)
                Case "latword": res = KL_LatToCyrWord(txt)
                Case "ping": res = KL_Ping()
                Case "author": res = Replace(KL_Author(), Chr(13) & Chr(10), " / ")
            End Select
            tout.writeString(res & Chr(10))
        End If
    Loop
    tout.closeOutput()
    ti.closeInput()
End Sub
'''.replace('file://W', 'file://' + W)

def xba(name, code):
    return ('<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">\n'
            '<script:module xmlns:script="http://openoffice.org/2000/script" script:name="%s" script:language="StarBasic">%s</script:module>\n'
            % (name, html.escape(code, quote=True)))
open(std + '/KLCore.xba', 'w').write(xba('KLCore', core))
open(std + '/Harness.xba', 'w').write(xba('Harness', harness))
open(std + '/script.xlb', 'w').write('<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE library:library PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "library.dtd">\n'
    '<library:library xmlns:library="http://openoffice.org/2000/library" library:name="Standard" library:readonly="false" library:passwordprotected="false">\n'
    ' <library:element library:name="KLCore"/>\n <library:element library:name="Harness"/>\n</library:library>\n')

# 2) test holatlari
cases = []
for src, _ in T.CYR_LAT + T.WORD_CYR_LAT + T.NEW_CYR_LAT:
    cases += [('cyr_new', src, R.cyr_to_lat(src)),
              ('cyr_old', src, R.cyr_to_lat(src, new_alphabet=False)),
              ('cyr_old_ascii', src, R.cyr_to_lat(src, ascii_apos=True, new_alphabet=False))]
for src, _ in T.LAT_CYR + T.WORD_LAT_CYR + T.NEW_LAT_CYR:
    cases.append(('lat', src, R.lat_to_cyr(src)))
    for w in src.split():
        cases.append(('latword', w, R.lat_to_cyr_word(w)))
version = re.search(r'KL_VERSION As String = "([^"]+)"', open(ROOT + '/src/KLCore.bas').read()).group(1)
cases.append(('ping', 'x', 'OK %s Shahar' % version))
cases.append(('author', 'x', 'Namangan \u015fahar ta\u02bclim b\u00f6limi AKT mutaxassisi / Dadaboyev Abdurahmon Abduqahhorovich'))
with open(W + '/cases.txt', 'w', encoding='utf-8') as f:
    for m, src, _ in cases:
        f.write(m + '\t' + src + '\n')

# 3) ishga tushirish
if os.path.exists(W + '/out.txt'):
    os.remove(W + '/out.txt')
r = subprocess.run(['soffice', '-env:UserInstallation=' + URL, '--headless', '--invisible',
                    'vnd.sun.star.script:Standard.Harness.Main?language=Basic&location=application'],
                   capture_output=True, text=True, timeout=180)
if not os.path.exists(W + '/out.txt'):
    print('NATIJA YO\'Q', r.stdout, r.stderr); sys.exit(2)
got = open(W + '/out.txt', encoding='utf-8').read().split('\n')
bad = 0
for (m, src, exp), g in zip(cases, got):
    if g != exp:
        bad += 1
        print('FARQ [%s] %r: VBA=%r Python=%r' % (m, src, g, exp))
print('VBA (LibreOffice): holatlar %d, natijalar %d, farqlar %d' % (len(cases), len(got) - 1, bad))
sys.exit(1 if bad or len(got) - 1 != len(cases) else 0)
