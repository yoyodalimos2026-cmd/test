Attribute VB_Name = "Module14"
Sub New_Xcomparer_Ref()
'Macro enregistrer le 21/06/2018 par Mr Yassine HASSEN
Dim wb As Workbook
Dim ws As Worksheet
Dim Message, Alerte, DEQ, RefInit, RefEquiv As String
Dim lin As Integer
For lin = 2 To Range("A65536").End(xlUp).Row
RefInit = ActiveWorkbook.Sheets(1).Cells(lin, 5).Value
RefEquiv = ActiveWorkbook.Sheets(1).Cells(lin, 11).Value
Message = ""
Alerte = ""
If Trim(RefEquiv) = Trim(RefInit) Then
Cells(lin, 1).Interior.Color = RGB(0, 255, 255)
End If
Application.Wait Now + TimeValue("0:00:01")
Next lin
MsgBox ("OK")
End Sub
