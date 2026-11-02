Attribute VB_Name = "KLCore"
' =====================================================================
'  Kirill-Lotin: o'zbek kirill <-> lotin transliteratsiya yadrosi
'  Word va Excel uchun umumiy modul.
'
'  DIQQAT: bu fayl faqat ASCII belgilardan iborat bo'lishi kerak
'  (VBA muharriri UTF-8 ni tushunmaydi). Barcha kirill va maxsus
'  belgilar ChrW() orqali yasaladi.
'
'  Qoidalar tests/reference.py dagi Python nusxasi bilan bir xil.
'  Bittasini o'zgartirsangiz, ikkinchisini ham o'zgartiring.
' =====================================================================
Option Explicit

Private Const APP_NAME As String = "KirillLotin"
Public Const KL_VERSION As String = "1.5.1"

Private mReady As Boolean
Private mAscii As Boolean            ' True: ' ishlatiladi, False: U+02BB / U+02BC
Private mNew As Boolean              ' True: yangi lotin alifbosi (sh->s-cedilla, ch->c-cedilla, g'->g-breve, o'->o-umlaut)
Private mNewLower As String          ' yangi alifbo harflari (kichik): s-cedilla c-cedilla g-breve o-umlaut
Private mNewUpper As String          ' ularning bosh harflari
Private mCyrLower As String          ' kichik kirill harflar
Private mCyrUpper As String          ' ularga mos bosh harflar
Private mCyrVowels As String         ' kirill unlilar
Private mApostrophes As String       ' lotin matnida apostrof sifatida qabul qilinadigan belgilar
Private mLatBase(0 To 25) As String  ' a..z -> kirill harf

' ---------------------------------------------------------------------
'  Jadvallar
' ---------------------------------------------------------------------
Private Sub InitTables()
    Dim i As Long
    If mReady Then Exit Sub

    ' U+0430..U+044F: a..ya (yo dan tashqari), bosh harflari U+0410..U+042F
    For i = &H430 To &H44F
        mCyrLower = mCyrLower & ChrW(i)
        mCyrUpper = mCyrUpper & ChrW(i - &H20)
    Next i
    ' yo, o', q, g', h
    mCyrLower = mCyrLower & ChrW(&H451) & ChrW(&H45E) & ChrW(&H49B) & ChrW(&H493) & ChrW(&H4B3)
    mCyrUpper = mCyrUpper & ChrW(&H401) & ChrW(&H40E) & ChrW(&H49A) & ChrW(&H492) & ChrW(&H4B2)

    ' a e yo i o u e(oborot) yu ya o' y(i)
    mCyrVowels = ChrW(&H430) & ChrW(&H435) & ChrW(&H451) & ChrW(&H438) & ChrW(&H43E) & _
                 ChrW(&H443) & ChrW(&H44D) & ChrW(&H44E) & ChrW(&H44F) & ChrW(&H45E) & ChrW(&H44B)

    ' ' ` (chap/o'ng qo'shtirnoq) okina tutuq akut
    mApostrophes = "'" & "`" & ChrW(&H2018) & ChrW(&H2019) & ChrW(&H2BB) & ChrW(&H2BC) & ChrW(&HB4)

    mLatBase(0) = ChrW(&H430)   ' a
    mLatBase(1) = ChrW(&H431)   ' b
    mLatBase(2) = ChrW(&H446)   ' c -> ts
    mLatBase(3) = ChrW(&H434)   ' d
    mLatBase(4) = ChrW(&H435)   ' e (kontekstga qarab alohida ishlanadi)
    mLatBase(5) = ChrW(&H444)   ' f
    mLatBase(6) = ChrW(&H433)   ' g
    mLatBase(7) = ChrW(&H4B3)   ' h
    mLatBase(8) = ChrW(&H438)   ' i
    mLatBase(9) = ChrW(&H436)   ' j
    mLatBase(10) = ChrW(&H43A)  ' k
    mLatBase(11) = ChrW(&H43B)  ' l
    mLatBase(12) = ChrW(&H43C)  ' m
    mLatBase(13) = ChrW(&H43D)  ' n
    mLatBase(14) = ChrW(&H43E)  ' o
    mLatBase(15) = ChrW(&H43F)  ' p
    mLatBase(16) = ChrW(&H49B)  ' q
    mLatBase(17) = ChrW(&H440)  ' r
    mLatBase(18) = ChrW(&H441)  ' s
    mLatBase(19) = ChrW(&H442)  ' t
    mLatBase(20) = ChrW(&H443)  ' u
    mLatBase(21) = ChrW(&H432)  ' v
    mLatBase(22) = ChrW(&H432)  ' w -> v
    mLatBase(23) = ChrW(&H445)  ' x
    mLatBase(24) = ChrW(&H439)  ' y
    mLatBase(25) = ChrW(&H437)  ' z

    ' Yangi lotin alifbosi harflari
    mNewLower = ChrW(&H15F) & ChrW(&HE7) & ChrW(&H11F) & ChrW(&HF6)
    mNewUpper = ChrW(&H15E) & ChrW(&HC7) & ChrW(&H11E) & ChrW(&HD6)

    mAscii = (GetSetting(APP_NAME, "Options", "AsciiApostrophe", "0") = "1")
    mNew = (GetSetting(APP_NAME, "Options", "NewAlphabet", "1") = "1")
    mReady = True
End Sub

' ---------------------------------------------------------------------
'  Sozlamalar
' ---------------------------------------------------------------------
Public Function KL_AsciiApostrophe() As Boolean
    InitTables
    KL_AsciiApostrophe = mAscii
End Function

Public Sub KL_SetAsciiApostrophe(ByVal value As Boolean)
    InitTables
    mAscii = value
    SaveSetting APP_NAME, "Options", "AsciiApostrophe", IIf(value, "1", "0")
End Sub

Public Function KL_NewAlphabet() As Boolean
    InitTables
    KL_NewAlphabet = mNew
End Function

Public Sub KL_SetNewAlphabet(ByVal value As Boolean)
    InitTables
    mNew = value
    SaveSetting APP_NAME, "Options", "NewAlphabet", IIf(value, "1", "0")
End Sub

Public Function KL_Title() As String
    KL_Title = "Kirill-Lotin " & KL_VERSION
End Function

' --- WINAPI BEGIN -----------------------------------------------------
' VBA ning MsgBox oynasi matnni ANSI kodlashga o'tkazadi, shuning uchun
' s-cedilla, o-umlaut, tutuq kabi harflar "?" bo'lib qolishi mumkin.
' Windows'ning Unicode oynasi (MessageBoxW) ishlatiladi; VBA7 (Office
' 2010+, 32/64 bit) va VBA6 (Office 2003/2007) uchun alohida e'lon.
#If VBA7 Then
Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hWnd As LongPtr, ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal uType As Long) As Long
#Else
Private Declare Function MessageBoxW Lib "user32" (ByVal hWnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal uType As Long) As Long
#End If

Public Sub KL_UnicodeMsg(ByVal text As String, ByVal title As String, ByVal flags As Long)
    On Error GoTo Fallback
    ' &H2000 = MB_TASKMODAL, &H10000 = MB_SETFOREGROUND
    MessageBoxW 0, StrPtr(text), StrPtr(title), flags Or &H2000 Or &H10000
    Exit Sub
Fallback:
    MsgBox text, flags, title
End Sub
' --- WINAPI END -------------------------------------------------------

' Muallif (Yordam oynasida ko'rsatiladi)
Public Function KL_Author() As String
    KL_Author = "Namangan " & ChrW(&H15F) & "ahar ta" & ChrW(&H2BC) & "lim b" & ChrW(&HF6) & _
                "limi AKT mutaxassisi" & vbCrLf & "Dadaboyev Abdurahmon Abduqahhorovich"
End Function

Public Sub KL_ShowAbout()
    Dim t As String
    t = "Kirill-Lotin " & KL_VERSION & vbCrLf & vbCrLf & _
        "O'zbek tilidagi matnni kirill alifbosidan lotin alifbosiga va aksincha o'giradi." & vbCrLf & vbCrLf & _
        "Foydalanish:" & vbCrLf & _
        "  1. Matnni (Word) yoki kataklarni (Excel) belgilang." & vbCrLf & _
        "     Hech narsa belgilanmasa, butun hujjat/varaq o'giriladi." & vbCrLf & _
        "  2. 'Kirill-Lotin' lentasidagi tugmani bosing." & vbCrLf & vbCrLf & _
        "Tezkor tugmalar:" & vbCrLf & _
        "  Alt+Shift+L  -  Kirill -> Lotin" & vbCrLf & _
        "  Alt+Shift+K  -  Lotin -> Kirill" & vbCrLf & vbCrLf & _
        "Excel formulalari:" & vbCrLf & _
        "  =LOTINGA(A1)   =KIRILGA(A1)" & vbCrLf & vbCrLf
    t = t & "Lotin alifbosi:" & vbCrLf & _
        "  'Yangi alifbo' belgilangan (sukut bo'yicha): sh -> " & ChrW(&H15F) & ", ch -> " & ChrW(&HE7) & _
        ", g' -> " & ChrW(&H11F) & ", o' -> " & ChrW(&HF6) & vbCrLf & _
        "  Belgi olib tashlansa: eski alifbo (sh, ch, g" & ChrW(&H2BB) & ", o" & ChrW(&H2BB) & ")." & vbCrLf & _
        "  Eski alifboda 'Oddiy apostrof' belgilansa, o' va g' uchun oddiy ' ishlatiladi." & vbCrLf & _
        "  Lotindan kirillga o'girishda ikkala yozuv ham tushuniladi." & vbCrLf & vbCrLf
    t = t & "Muallif:" & vbCrLf & KL_Author()
    KL_UnicodeMsg t, KL_Title(), vbInformation
End Sub

' O'rnatuvchi qo'shimcha ishlayotganini shu funksiya orqali tekshiradi
Public Function KL_Ping() As String
    ' Sozlamalardan qat'i nazar eski alifbo bilan: natija har doim "Shahar"
    KL_Ping = "OK " & KL_VERSION & " " & CyrToLatEx(ChrW(&H428) & ChrW(&H430) & ChrW(&H4B3) & ChrW(&H430) & ChrW(&H440), False, True)
End Function

' Office 2003 va undan eski versiyalar uchun asboblar paneli tugmasi
' (lenta yo'q). Hamma chaqiruvlar kechiktirilgan bog'lanish (Object)
' orqali - Office kutubxonasiga havola kerak emas.
Public Sub KL_AddToolbarButton(ByVal bar As Object, ByVal caption As String, ByVal macroName As String)
    Dim btn As Object
    Set btn = bar.Controls.Add(Type:=1)     ' msoControlButton
    btn.Caption = caption
    btn.Style = 2                           ' msoButtonCaption
    btn.OnAction = macroName
End Sub

' Lenta (Ribbon) uchun umumiy chaqiruvlar.
' control ning turi Object: Office kutubxonasiga havola kerak bo'lmasin.
Public Sub KL_RibbonGetAscii(control As Object, ByRef returnedVal)
    returnedVal = KL_AsciiApostrophe()
End Sub

Public Sub KL_RibbonGetNew(control As Object, ByRef returnedVal)
    returnedVal = KL_NewAlphabet()
End Sub

Public Sub KL_RibbonSetNew(control As Object, pressed As Boolean)
    KL_SetNewAlphabet pressed
End Sub

Public Sub KL_RibbonSetAscii(control As Object, pressed As Boolean)
    KL_SetAsciiApostrophe pressed
End Sub

Public Sub KL_RibbonAbout(control As Object)
    KL_ShowAbout
End Sub

' ---------------------------------------------------------------------
'  Yordamchi funksiyalar
' ---------------------------------------------------------------------
Private Function IsCyr(ByVal ch As String) As Boolean
    If Len(ch) <> 1 Then Exit Function
    IsCyr = (InStr(1, mCyrLower, ch, vbBinaryCompare) > 0) Or (InStr(1, mCyrUpper, ch, vbBinaryCompare) > 0)
End Function

Private Function IsCyrUpper(ByVal ch As String) As Boolean
    If Len(ch) <> 1 Then Exit Function
    IsCyrUpper = (InStr(1, mCyrUpper, ch, vbBinaryCompare) > 0)
End Function

Private Function CyrLower(ByVal ch As String) As String
    Dim p As Long
    CyrLower = ch
    If Len(ch) <> 1 Then Exit Function
    p = InStr(1, mCyrUpper, ch, vbBinaryCompare)
    If p > 0 Then CyrLower = Mid$(mCyrLower, p, 1)
End Function

Private Function CyrUpper(ByVal ch As String) As String
    Dim p As Long
    CyrUpper = ch
    If Len(ch) <> 1 Then Exit Function
    p = InStr(1, mCyrLower, ch, vbBinaryCompare)
    If p > 0 Then CyrUpper = Mid$(mCyrUpper, p, 1)
End Function

Private Function IsLat(ByVal ch As String) As Boolean
    Dim code As Long
    If Len(ch) <> 1 Then Exit Function
    code = AscW(ch)
    IsLat = (code >= 65 And code <= 90) Or (code >= 97 And code <= 122) Or _
            (InStr(1, mNewLower & mNewUpper, ch, vbBinaryCompare) > 0)
End Function

Private Function IsLatUpper(ByVal ch As String) As Boolean
    Dim code As Long
    If Len(ch) <> 1 Then Exit Function
    code = AscW(ch)
    IsLatUpper = (code >= 65 And code <= 90) Or (InStr(1, mNewUpper, ch, vbBinaryCompare) > 0)
End Function

' Lotin harfni kichik/bosh harfga o'tkazish (yangi alifbo harflari ham).
' UCase$/LCase$ ga tayanmaymiz: Windows tiliga qarab natija farq qilishi mumkin.
Private Function LatLower(ByVal ch As String) As String
    Dim p As Long
    LatLower = LCase$(ch)
    If Len(ch) <> 1 Then Exit Function
    p = InStr(1, mNewUpper, ch, vbBinaryCompare)
    If p > 0 Then LatLower = Mid$(mNewLower, p, 1)
End Function

Private Function LatUpper(ByVal t As String) As String
    Dim i As Long, ch As String, p As Long, res As String
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        p = InStr(1, mNewLower, ch, vbBinaryCompare)
        If p > 0 Then
            res = res & Mid$(mNewUpper, p, 1)
        Else
            res = res & UCase$(ch)
        End If
    Next i
    LatUpper = res
End Function

Private Function IsApos(ByVal ch As String) As Boolean
    If Len(ch) <> 1 Then Exit Function
    IsApos = (InStr(1, mApostrophes, ch, vbBinaryCompare) > 0)
End Function

Private Function IsSpaceChar(ByVal ch As String) As Boolean
    Select Case AscW(ch)
        Case 7, 9, 10, 11, 12, 13, 32, 160
            IsSpaceChar = True
    End Select
End Function

Private Function CharAt(ByRef s As String, ByVal i As Long) As String
    If i >= 1 And i <= Len(s) Then CharAt = Mid$(s, i, 1)
End Function

' Kichik kirill harfning asosiy lotincha ko'rinishi
Private Function CyrBase(ByVal code As Long, ByVal okina As String, ByVal newAlpha As Boolean) As String
    If newAlpha Then
        Select Case code
            Case &H447: CyrBase = ChrW(&HE7): Exit Function            ' ch -> c-cedilla
            Case &H448, &H449: CyrBase = ChrW(&H15F): Exit Function    ' sh -> s-cedilla
            Case &H45E: CyrBase = ChrW(&HF6): Exit Function            ' o' -> o-umlaut
            Case &H493: CyrBase = ChrW(&H11F): Exit Function           ' g' -> g-breve
        End Select
    End If
    Select Case code
        Case &H430: CyrBase = "a"
        Case &H431: CyrBase = "b"
        Case &H432: CyrBase = "v"
        Case &H433: CyrBase = "g"
        Case &H434: CyrBase = "d"
        Case &H451: CyrBase = "yo"
        Case &H436: CyrBase = "j"
        Case &H437: CyrBase = "z"
        Case &H438: CyrBase = "i"
        Case &H439: CyrBase = "y"
        Case &H43A: CyrBase = "k"
        Case &H43B: CyrBase = "l"
        Case &H43C: CyrBase = "m"
        Case &H43D: CyrBase = "n"
        Case &H43E: CyrBase = "o"
        Case &H43F: CyrBase = "p"
        Case &H440: CyrBase = "r"
        Case &H441: CyrBase = "s"
        Case &H442: CyrBase = "t"
        Case &H443: CyrBase = "u"
        Case &H444: CyrBase = "f"
        Case &H445: CyrBase = "x"
        Case &H447: CyrBase = "ch"
        Case &H448: CyrBase = "sh"
        Case &H449: CyrBase = "sh"
        Case &H44B: CyrBase = "i"
        Case &H44C: CyrBase = ""
        Case &H44D: CyrBase = "e"
        Case &H44E: CyrBase = "yu"
        Case &H44F: CyrBase = "ya"
        Case &H45E: CyrBase = "o" & okina
        Case &H49B: CyrBase = "q"
        Case &H493: CyrBase = "g" & okina
        Case &H4B3: CyrBase = "h"
    End Select
End Function

Public Function KL_IsUrlLike(ByVal token As String) As Boolean
    Dim t As String
    t = LCase$(token)
    KL_IsUrlLike = (InStr(1, t, "://", vbBinaryCompare) > 0) Or _
                   (InStr(1, t, "www.", vbBinaryCompare) > 0) Or _
                   (InStr(1, t, "@", vbBinaryCompare) > 0)
End Function

' Word "Find" uchun wildcard shablon: so'zning BITTA harfi.
' Diqqat: Word'da "[...]@" butun so'zni emas, ko'pincha bitta harfni
' topadi. Shuning uchun so'z boshini shu shablon bilan topib, keyin
' KL_WordCharset belgilari bo'yicha MoveEndWhile bilan kengaytiramiz.
Public Function KL_WordPattern(ByVal toLatin As Boolean) As String
    InitTables
    If toLatin Then
        KL_WordPattern = "[" & mCyrLower & mCyrUpper & "]"
    Else
        KL_WordPattern = "[A-Za-z" & mNewLower & mNewUpper & mApostrophes & "]"
    End If
End Function

' So'z tarkibiga kiradigan belgilar (MoveEndWhile uchun)
Public Function KL_WordCharset(ByVal toLatin As Boolean) As String
    InitTables
    If toLatin Then
        KL_WordCharset = mCyrLower & mCyrUpper
    Else
        KL_WordCharset = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz" & mNewLower & mNewUpper & mApostrophes
    End If
End Function

' ---------------------------------------------------------------------
'  KIRILL -> LOTIN
' ---------------------------------------------------------------------
Public Function KL_CyrToLat(ByVal s As String) As String
    InitTables
    KL_CyrToLat = CyrToLatEx(s, mNew, mAscii)
End Function

Private Function CyrToLatEx(ByVal s As String, ByVal newAlpha As Boolean, ByVal asciiApos As Boolean) As String
    Dim i As Long, n As Long
    Dim c As String, prevC As String, nextC As String
    Dim lc As String, lp As String, ln As String, t As String, res As String
    Dim okina As String, tutuq As String

    InitTables
    If asciiApos Then
        okina = "'"
        tutuq = "'"
    Else
        okina = ChrW(&H2BB)
        tutuq = ChrW(&H2BC)
    End If

    n = Len(s)
    For i = 1 To n
        c = Mid$(s, i, 1)
        If Not IsCyr(c) Then
            res = res & c
        Else
            prevC = CharAt(s, i - 1)
            nextC = CharAt(s, i + 1)
            lc = CyrLower(c)
            lp = CyrLower(prevC)

            Select Case AscW(lc)
                Case &H435  ' e: so'z boshida, unli, ' yoki ' dan keyin "ye"
                    If Not IsCyr(prevC) Then
                        t = "ye"
                    ElseIf InStr(1, mCyrVowels & ChrW(&H44A) & ChrW(&H44C), lp, vbBinaryCompare) > 0 Then
                        t = "ye"
                    Else
                        t = "e"
                    End If
                Case &H446  ' ts: unlidan keyin "ts", aks holda "s"
                    If IsCyr(prevC) And InStr(1, mCyrVowels, lp, vbBinaryCompare) > 0 Then
                        t = "ts"
                    Else
                        t = "s"
                    End If
                Case &H44A  ' ayirish belgisi: e/yo/yu/ya oldidan tushib qoladi
                    ln = CyrLower(nextC)
                    If ln = ChrW(&H435) Or ln = ChrW(&H451) Or ln = ChrW(&H44E) Or ln = ChrW(&H44F) Then
                        t = ""
                    Else
                        t = tutuq
                    End If
                Case Else
                    t = CyrBase(AscW(lc), okina, newAlpha)
            End Select

            If IsCyrUpper(c) And Len(t) > 0 Then
                If Len(t) > 1 And (IsCyrUpper(nextC) Or (Not IsCyr(nextC) And IsCyrUpper(prevC))) Then
                    t = LatUpper(t)                            ' SH, CH, YO ... (bosh harfli so'z)
                Else
                    t = LatUpper(Left$(t, 1)) & Mid$(t, 2)     ' Sh, Ch, Yo, S-cedilla ...
                End If
            End If
            res = res & t
        End If
    Next i
    CyrToLatEx = res
End Function

' ---------------------------------------------------------------------
'  LOTIN -> KIRILL
' ---------------------------------------------------------------------
' Bitta so'z (bo'shliqsiz bo'lak) uchun
Public Function KL_LatToCyrWord(ByVal s As String) As String
    Dim i As Long, n As Long, stepN As Long
    Dim c As String, prevC As String, nextC As String, next2 As String
    Dim lc As String, ln As String, t As String, res As String

    InitTables
    n = Len(s)
    i = 1
    Do While i <= n
        c = Mid$(s, i, 1)
        prevC = CharAt(s, i - 1)
        nextC = CharAt(s, i + 1)
        next2 = CharAt(s, i + 2)
        stepN = 1

        If IsApos(c) Then
            If IsLat(prevC) And IsLat(nextC) Then
                If LCase$(prevC) = "s" And LCase$(nextC) = "h" Then
                    t = ""                        ' s'h -> sh emas, s + h
                ElseIf IsLatUpper(prevC) And IsLatUpper(nextC) Then
                    t = ChrW(&H42A)               ' katta ayirish belgisi
                Else
                    t = ChrW(&H44A)               ' kichik ayirish belgisi
                End If
            Else
                t = c                             ' qo'shtirnoq va h.k. o'zgarmaydi
            End If
        ElseIf Not IsLat(c) Then
            t = c
        Else
            lc = LatLower(c)
            ln = LatLower(nextC)
            If InStr(1, mNewLower, lc, vbBinaryCompare) > 0 And Len(lc) = 1 Then
                ' yangi alifbo: s-cedilla -> sh, c-cedilla -> ch, g-breve -> g', o-umlaut -> o'
                Select Case AscW(lc)
                    Case &H15F: t = ChrW(&H448)
                    Case &HE7: t = ChrW(&H447)
                    Case &H11F: t = ChrW(&H493)
                    Case &HF6: t = ChrW(&H45E)
                End Select
            ElseIf (lc = "o" Or lc = "g") And IsApos(nextC) Then
                If lc = "o" Then t = ChrW(&H45E) Else t = ChrW(&H493)
                stepN = 2
            ElseIf lc = "s" And ln = "h" Then
                t = ChrW(&H448)
                stepN = 2
            ElseIf lc = "c" And ln = "h" Then
                t = ChrW(&H447)
                stepN = 2
            ElseIf lc = "y" And (ln = "a" Or ln = "e" Or ln = "u" Or (ln = "o" And Not IsApos(next2))) Then
                Select Case ln
                    Case "a": t = ChrW(&H44F)
                    Case "e": t = ChrW(&H435)
                    Case "u": t = ChrW(&H44E)
                    Case "o": t = ChrW(&H451)
                End Select
                stepN = 2
            ElseIf lc = "e" Then
                If Not (IsLat(prevC) Or IsApos(prevC)) Then
                    t = ChrW(&H44D)               ' so'z boshida
                ElseIf InStr(1, "aeiou" & ChrW(&HF6), LatLower(prevC), vbBinaryCompare) > 0 Then
                    t = ChrW(&H44D)               ' unlidan keyin
                Else
                    t = ChrW(&H435)
                End If
            Else
                t = mLatBase(AscW(lc) - 97)
            End If
            If IsLatUpper(c) Then t = CyrUpper(t)
        End If

        res = res & t
        i = i + stepN
    Loop
    KL_LatToCyrWord = res
End Function

' Butun matn uchun: havolalar va e-pochta manzillari o'zgarmaydi
Public Function KL_LatToCyr(ByVal s As String) As String
    Dim i As Long, j As Long, n As Long
    Dim ch As String, tok As String, res As String

    n = Len(s)
    i = 1
    Do While i <= n
        ch = Mid$(s, i, 1)
        If IsSpaceChar(ch) Then
            res = res & ch
            i = i + 1
        Else
            j = i
            Do While j <= n
                If IsSpaceChar(Mid$(s, j, 1)) Then Exit Do
                j = j + 1
            Loop
            tok = Mid$(s, i, j - i)
            If KL_IsUrlLike(tok) Then
                res = res & tok
            Else
                res = res & KL_LatToCyrWord(tok)
            End If
            i = j
        End If
    Loop
    KL_LatToCyr = res
End Function
