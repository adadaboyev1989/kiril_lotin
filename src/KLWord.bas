Attribute VB_Name = "KLWord"
' =====================================================================
'  Kirill-Lotin: Microsoft Word uchun buyruqlar
'  Matn "Find" orqali so'zma-so'z almashtiriladi, shuning uchun
'  shrift, rang, jadval va boshqa formatlash saqlanib qoladi.
' =====================================================================
Option Explicit

' --- Makroslar (tezkor tugmalar shularga bog'langan) ------------------
Public Sub KL_WordToLatin()
    ConvertWord True
End Sub

Public Sub KL_WordToCyrillic()
    ConvertWord False
End Sub

' --- Lenta (Ribbon) chaqiruvlari -------------------------------------
Public Sub KL_RibbonToLatin(control As IRibbonControl)
    ConvertWord True
End Sub

Public Sub KL_RibbonToCyrillic(control As IRibbonControl)
    ConvertWord False
End Sub

' --- Asosiy ish -------------------------------------------------------
Private Sub ConvertWord(ByVal toLatin As Boolean)
    Dim doc As Document
    Dim useSelection As Boolean
    Dim undoRec As Object
    Dim story As Range, r As Range
    Dim total As Long

    If Documents.Count = 0 Then
        MsgBox "Ochiq hujjat yo'q.", vbExclamation, KL_Title()
        Exit Sub
    End If
    Set doc = ActiveDocument
    If doc.ProtectionType <> wdNoProtection Then
        MsgBox "Hujjat himoyalangan. Avval himoyani olib tashlang.", vbExclamation, KL_Title()
        Exit Sub
    End If

    useSelection = (Selection.Type <> wdSelectionIP) And (Len(Selection.Range.Text) > 0)
    If Not useSelection Then
        If MsgBox("Matn belgilanmagan." & vbCrLf & _
                  "Butun hujjat " & IIf(toLatin, "lotin", "kirill") & " alifbosiga o'girilsinmi?", _
                  vbYesNo + vbQuestion, KL_Title()) <> vbYes Then Exit Sub
    End If

    ' Bitta Ctrl+Z bilan bekor qilish uchun (Word 2010+)
    On Error Resume Next
    Set undoRec = Application.UndoRecord
    If Not undoRec Is Nothing Then undoRec.StartCustomRecord IIf(toLatin, "Kirill -> Lotin", "Lotin -> Kirill")
    On Error GoTo Fail

    Application.ScreenUpdating = False
    Application.StatusBar = "Kirill-Lotin: o'girilmoqda..."

    If useSelection Then
        total = ConvertRange(Selection.Range, toLatin)
    Else
        ' Asosiy matn, kolontitullar, izohlar, matn maydonlari ...
        For Each story In doc.StoryRanges
            Set r = story
            Do While Not r Is Nothing
                total = total + ConvertRange(r, toLatin)
                Set r = r.NextStoryRange
            Loop
        Next story
    End If

Done:
    On Error Resume Next
    If Not undoRec Is Nothing Then undoRec.EndCustomRecord
    Application.ScreenUpdating = True
    Application.StatusBar = "Kirill-Lotin: " & total & " ta so'z o'girildi."
    Exit Sub

Fail:
    MsgBox "Xatolik: " & Err.Description, vbCritical, KL_Title()
    Resume Done
End Sub

' Diapazondagi har bir so'zni topib, o'giradi. O'zgargan so'zlar sonini qaytaradi.
Private Function ConvertRange(ByVal scope As Range, ByVal toLatin As Boolean) As Long
    Dim rng As Range
    Dim oldText As String, newText As String
    Dim n As Long

    Set rng = scope.Duplicate
    rng.Collapse wdCollapseStart
    With rng.Find
        .ClearFormatting
        .Text = KL_WordPattern(toLatin)
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .MatchWildcards = True
    End With

    Do While rng.Find.Execute
        If rng.Start >= scope.End Then Exit Do
        If rng.End > scope.End Then rng.End = scope.End

        oldText = rng.Text
        If toLatin Then
            newText = KL_CyrToLat(oldText)
        ElseIf IsInsideLink(rng) Then
            newText = oldText
        Else
            newText = KL_LatToCyrWord(oldText)
        End If

        If newText <> oldText Then
            rng.Text = newText
            n = n + 1
        End If

        rng.Collapse wdCollapseEnd
        If rng.End >= scope.End Then Exit Do
    Loop
    ConvertRange = n
End Function

' So'z havola (http://, www.) yoki e-pochta manzilining bir qismimi?
Private Function IsInsideLink(ByVal rng As Range) As Boolean
    Dim t As Range
    Dim delims As String
    On Error GoTo Quit
    delims = " " & vbTab & vbCr & vbLf & Chr(11) & ChrW(160)
    Set t = rng.Duplicate
    t.MoveStartUntil Cset:=delims, Count:=wdBackward
    t.MoveEndUntil Cset:=delims, Count:=wdForward
    IsInsideLink = KL_IsUrlLike(t.Text)
Quit:
End Function
