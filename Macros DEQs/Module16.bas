Attribute VB_Name = "Module16"
Sub New_Import_Ref()


'Macro enregistrer le 15/01/2019 par Mr Yassine HASSEN


Dim wb As Workbook
Dim ws As Worksheet
Dim Message, Alerte, DEQ, RefInit, RefEquiv, FabInit, FabEquiv, Design, Num As String
Dim lin As Integer



For lin = 2 To Range("A65536").End(xlUp).Row

DEQ = ActiveWorkbook.Sheets(1).Cells(lin, 1).Value
RefInit = ActiveWorkbook.Sheets(1).Cells(lin, 5).Value
FabInit = ActiveWorkbook.Sheets(1).Cells(lin, 6).Value
RefEquiv = ActiveWorkbook.Sheets(1).Cells(lin, 11).Value
FabEquiv = ActiveWorkbook.Sheets(1).Cells(lin, 12).Value
'Design = ActiveWorkbook.Sheets(1).Cells(lin, 7).Value
Num = ActiveWorkbook.Sheets(1).Cells(lin, 1).Value
Message = ""
Alerte = ""

Set wb = Workbooks.Open(ActiveWorkbook.Path & "\" & DEQ & "\COMP.xlsx")

Set ws = wb.Worksheets(1)
Application.Wait Now + TimeValue("0:00:01")

ws.Cells(16, 3).Value = RefInit
ws.Cells(18, 3).Value = FabInit
ws.Cells(16, 4).Value = RefEquiv
ws.Cells(18, 4).Value = FabEquiv
'ws.Cells(17, 3).Value = Design
ws.Cells(1, 4).Value = Num
wb.Save
wb.Close

Application.Wait Now + TimeValue("0:00:01")

Next lin
MsgBox ("les INIT,EQUIV & les Fabricants sont importés avec succés")

End Sub










