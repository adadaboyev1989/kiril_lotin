Option Explicit

' Kirill-Lotin qo'shimchasi yuklanganda tezkor tugmalarni o'rnatadi
Private Sub Workbook_Open()
    On Error Resume Next
    Application.OnKey "%+l", "'" & ThisWorkbook.Name & "'!KL_ExcelToLatin"
    Application.OnKey "%+k", "'" & ThisWorkbook.Name & "'!KL_ExcelToCyrillic"
    KL_ExcelCreateToolbar
    Application.MacroOptions Macro:="LOTINGA", Description:="Kirill matnni lotin alifbosiga o'giradi", Category:=7
    Application.MacroOptions Macro:="KIRILGA", Description:="Lotin matnni kirill alifbosiga o'giradi", Category:=7
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    Application.OnKey "%+l"
    Application.OnKey "%+k"
    KL_ExcelDeleteToolbar
End Sub
