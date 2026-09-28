Attribute VB_Name = "KLExcel"
' =====================================================================
'  Kirill-Lotin: Microsoft Excel uchun buyruqlar va formulalar
' =====================================================================
Option Explicit

' Oxirgi o'girishni bekor qilish (Ctrl+Z) uchun ma'lumotlar
Private mUndoSheet As Object
Private mUndoAddr() As String
Private mUndoVal() As String
Private mUndoCount As Long

' --- Formulalar -------------------------------------------------------
' =LOTINGA(A1)  - kirill matnni lotinga o'giradi
Public Function LOTINGA(ByVal matn As Variant) As Variant
    LOTINGA = Translate(matn, True)
End Function

' =KIRILGA(A1)  - lotin matnni kirillga o'giradi
Public Function KIRILGA(ByVal matn As Variant) As Variant
    KIRILGA = Translate(matn, False)
End Function

Private Function Translate(ByVal v As Variant, ByVal toLatin As Boolean) As Variant
    If TypeName(v) = "Range" Then v = v.Cells(1, 1).Value
    If IsError(v) Then
        Translate = v
    ElseIf toLatin Then
        Translate = KL_CyrToLat(CStr(v))
    Else
        Translate = KL_LatToCyr(CStr(v))
    End If
End Function

' --- Makroslar (tezkor tugmalar shularga bog'langan) ------------------
Public Sub KL_ExcelToLatin()
    ConvertExcel True
End Sub

Public Sub KL_ExcelToCyrillic()
    ConvertExcel False
End Sub

' --- Lenta (Ribbon) chaqiruvlari -------------------------------------
Public Sub KL_RibbonToLatin(control As Object)
    ConvertExcel True
End Sub

Public Sub KL_RibbonToCyrillic(control As Object)
    ConvertExcel False
End Sub

' --- Asosiy ish -------------------------------------------------------
Private Sub ConvertExcel(ByVal toLatin As Boolean)
    Dim target As Range, txtCells As Range, c As Range
    Dim oldText As String, newText As String
    Dim calc As Long, n As Long

    If ActiveWorkbook Is Nothing Then
        MsgBox "Ochiq kitob yo'q.", vbExclamation, KL_Title()
        Exit Sub
    End If

    If TypeName(Selection) <> "Range" Then
        ConvertShapes toLatin
        Exit Sub
    End If

    If ActiveSheet.ProtectContents Then
        MsgBox "Varaq himoyalangan. Avval himoyani olib tashlang.", vbExclamation, KL_Title()
        Exit Sub
    End If

    Set target = Selection
    If target.CountLarge = 1 Then
        Select Case MsgBox("Faqat bitta katak belgilangan." & vbCrLf & vbCrLf & _
                           "Ha - butun varaqni o'girish" & vbCrLf & _
                           "Yo'q - faqat shu katakni o'girish", _
                           vbYesNoCancel + vbQuestion, KL_Title())
            Case vbYes: Set target = ActiveSheet.UsedRange
            Case vbNo
            Case Else: Exit Sub
        End Select
    End If

    ' Faqat matnli qiymatlar (formulalar va sonlarga tegilmaydi)
    If target.CountLarge = 1 Then
        Set txtCells = target
    Else
        On Error Resume Next
        Set txtCells = target.SpecialCells(xlCellTypeConstants, xlTextValues)
        On Error GoTo 0
    End If
    If txtCells Is Nothing Then
        MsgBox "O'giriladigan matn topilmadi.", vbInformation, KL_Title()
        Exit Sub
    End If

    Set mUndoSheet = ActiveSheet
    mUndoCount = 0
    ReDim mUndoAddr(1 To txtCells.CountLarge)
    ReDim mUndoVal(1 To txtCells.CountLarge)

    calc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    On Error GoTo Fail

    For Each c In txtCells.Cells
        If Not c.HasFormula And VarType(c.Value) = vbString Then
            oldText = c.Value
            If toLatin Then
                newText = KL_CyrToLat(oldText)
            Else
                newText = KL_LatToCyr(oldText)
            End If
            If newText <> oldText Then
                n = n + 1
                mUndoAddr(n) = c.Address(False, False)
                mUndoVal(n) = c.PrefixCharacter & oldText
                mUndoCount = n
                WriteText c, newText, c.PrefixCharacter
            End If
        End If
    Next c

Done:
    On Error Resume Next
    Application.Calculation = calc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If mUndoCount > 0 Then
        Application.OnUndo "Kirill-Lotin: bekor qilish", "'" & ThisWorkbook.Name & "'!KL_ExcelUndo"
    End If
    Exit Sub

Fail:
    MsgBox "Xatolik: " & Err.Description, vbCritical, KL_Title()
    Resume Done
End Sub

' Matnni katakka yozadi; son, sana yoki formula bo'lib qolmasligi uchun ' qo'yadi
Private Sub WriteText(ByVal c As Range, ByVal t As String, ByVal prefix As String)
    Dim first As String
    first = Left$(t, 1)
    If prefix = "'" Or IsNumeric(t) Or IsDate(t) Or first = "=" Or first = "+" Or first = "-" Or first = "@" Then
        c.Value = "'" & t
    Else
        c.Value = t
    End If
End Sub

' Belgilangan shakllar (matn maydoni, figura) ichidagi matn
Private Sub ConvertShapes(ByVal toLatin As Boolean)
    Dim shp As Object, n As Long
    Dim oldText As String, newText As String
    On Error Resume Next
    For Each shp In Selection.ShapeRange
        If shp.TextFrame2.HasText Then
            oldText = shp.TextFrame2.TextRange.Text
            If toLatin Then newText = KL_CyrToLat(oldText) Else newText = KL_LatToCyr(oldText)
            If newText <> oldText Then
                shp.TextFrame2.TextRange.Text = newText
                n = n + 1
            End If
        End If
    Next shp
    If Err.Number <> 0 And n = 0 Then
        MsgBox "Kataklarni yoki matnli shaklni belgilang.", vbExclamation, KL_Title()
    End If
End Sub

' Application.OnUndo orqali chaqiriladi
Public Sub KL_ExcelUndo()
    Dim i As Long
    If mUndoSheet Is Nothing Then Exit Sub
    On Error Resume Next
    Application.ScreenUpdating = False
    For i = 1 To mUndoCount
        mUndoSheet.Range(mUndoAddr(i)).Value = mUndoVal(i)
    Next i
    Application.ScreenUpdating = True
    mUndoCount = 0
    Set mUndoSheet = Nothing
End Sub
