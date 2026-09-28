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
Public Const KL_VERSION As String = "1.4.0"

Private mReady As Boolean
Private mAscii As Boolean            ' True: ' ishlatiladi, False: U+02BB / U+02BC
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

    mAscii = (GetSetting(APP_NAME, "Options", "AsciiApostrophe", "0") = "1")
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

Public Function KL_Title() As String
    KL_Title = "Kirill-Lotin " & KL_VERSION
End Function

Public Sub KL_ShowAbout()
    MsgBox "Kirill-Lotin " & KL_VERSION & vbCrLf & vbCrLf & _
        "O'zbek tilidagi matnni kirill alifbosidan lotin alifbosiga va aksincha o'giradi." & vbCrLf & vbCrLf & _
        "Foydalanish:" & vbCrLf & _
        "  1. Matnni (Word) yoki kataklarni (Excel) belgilang." & vbCrLf & _
        "     Hech narsa belgilanmasa, butun hujjat/varaq o'giriladi." & vbCrLf & _
        "  2. 'Kirill-Lotin' lentasidagi tugmani bosing." & vbCrLf & vbCrLf & _
        "Tezkor tugmalar:" & vbCrLf & _
        "  Alt+Shift+L  -  Kirill -> Lotin" & vbCrLf & _
        "  Alt+Shift+K  -  Lotin -> Kirill" & vbCrLf & vbCrLf & _
        "Excel formulalari:" & vbCrLf & _
        "  =LOTINGA(A1)   =KIRILGA(A1)" & vbCrLf & vbCrLf & _
        "'Oddiy apostrof' belgilansa, o' va g' uchun oddiy ' belgisi," & vbCrLf & _
        "aks holda rasmiy imlodagi " & ChrW(&H2BB) & " va " & ChrW(&H2BC) & " belgilari ishlatiladi.", _
        vbInformation, KL_Title()
End Sub

' O'rnatuvchi qo'shimcha ishlayotganini shu funksiya orqali tekshiradi
Public Function KL_Ping() As String
    KL_Ping = "OK " & KL_VERSION & " " & KL_CyrToLat(ChrW(&H428) & ChrW(&H430) & ChrW(&H4B3) & ChrW(&H430) & ChrW(&H440))
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
    IsLat = (code >= 65 And code <= 90) Or (code >= 97 And code <= 122)
End Function

Private Function IsLatUpper(ByVal ch As String) As Boolean
    Dim code As Long
    If Len(ch) <> 1 Then Exit Function
    code = AscW(ch)
    IsLatUpper = (code >= 65 And code <= 90)
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
Private Function CyrBase(ByVal code As Long, ByVal okina As String) As String
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
        KL_WordPattern = "[A-Za-z" & mApostrophes & "]"
    End If
End Function

' So'z tarkibiga kiradigan belgilar (MoveEndWhile uchun)
Public Function KL_WordCharset(ByVal toLatin As Boolean) As String
    InitTables
    If toLatin Then
        KL_WordCharset = mCyrLower & mCyrUpper
    Else
        KL_WordCharset = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz" & mApostrophes
    End If
End Function

' ---------------------------------------------------------------------
'  KIRILL -> LOTIN
' ---------------------------------------------------------------------
Public Function KL_CyrToLat(ByVal s As String) As String
    Dim i As Long, n As Long
    Dim c As String, prevC As String, nextC As String
    Dim lc As String, lp As String, ln As String, t As String, res As String
    Dim okina As String, tutuq As String

    InitTables
    If mAscii Then
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
                    t = CyrBase(AscW(lc), okina)
            End Select

            If IsCyrUpper(c) And Len(t) > 0 Then
                If Len(t) > 1 And (IsCyrUpper(nextC) Or (Not IsCyr(nextC) And IsCyrUpper(prevC))) Then
                    t = UCase$(t)                              ' SH, CH, YO ... (bosh harfli so'z)
                Else
                    t = UCase$(Left$(t, 1)) & Mid$(t, 2)       ' Sh, Ch, Yo ...
                End If
            End If
            res = res & t
        End If
    Next i
    KL_CyrToLat = res
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
            lc = LCase$(c)
            ln = LCase$(nextC)
            If (lc = "o" Or lc = "g") And IsApos(nextC) Then
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
                ElseIf InStr(1, "aeiou", LCase$(prevC), vbBinaryCompare) > 0 Then
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
