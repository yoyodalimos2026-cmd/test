Attribute VB_Name = "Module8"
Sub TabVerif()

Dim wb As Workbook
Dim ws As Worksheet
Dim Message, Alerte, DEQ, RefInit, RefEquiv, FabInit, FabEquiv As String


For lin = 2 To 3 'ActiveWorkbook.Sheets(1).Columns(1).Cells.Find("*", , , , , xlPrevious).Row

DEQ = ActiveWorkbook.Sheets(1).Cells(lin, 1).Value
RefInit = ActiveWorkbook.Sheets(1).Cells(lin, 5).Value
FabInit = ActiveWorkbook.Sheets(1).Cells(lin, 6).Value
RefEquiv = ActiveWorkbook.Sheets(1).Cells(lin, 10).Value
FabEquiv = ActiveWorkbook.Sheets(1).Cells(lin, 11).Value
Message = ""
Alerte = ""


Set wb = Workbooks.Open(ActiveWorkbook.Path & "\" & DEQ & "\COMP.xls")
Set ws = wb.Worksheets(1)
Application.Wait Now + TimeValue("0:00:01")
If ws.Cells(1, 3).Value <> DEQ Then
Message = " N° DEQ /"
End If
MsgBox (KillSpace(ws.Cells(16, 2).Value) & RefInit)
If KillSpace(ws.Cells(16, 2).Value) <> Trim(RefInit) And Trim(ws.Cells(16, 2).Value) <> Trim(RefInit) Then
Message = Message & " RefINIT /"
End If
If Trim(ws.Cells(18, 2).Value) <> Trim(FabInit) Then
Alerte = Alerte & " FabInit /"
End If



If RefEquiv = "SANS" Then

If ws.Cells(16, 2).MergeArea.Columns.Count <> 2 Then
Message = "ERREUR !!"
MsgBox (ws.Cells(16, 2).Value)
End If
Else

If Trim(ws.Cells(18, 3).Value) <> Trim(FabEquiv) Then
Alerte = Alerte & " FabEquiv /"
End If

If KillSpace(ws.Cells(16, 3).Value) <> Trim(RefEquiv) And Trim(ws.Cells(16, 3).Value) <> Trim(RefEquiv) Then
Message = Message & " RefEQUIV /"
End If

End If

wb.Close
ActiveWorkbook.Sheets(1).Cells(lin, 31) = Message
ActiveWorkbook.Sheets(1).Cells(lin, 32) = Alerte
Application.Wait Now + TimeValue("0:00:01")

Next lin

End Sub

Function KillSpace(ByVal FileName As String) As String

   

 
          
     KillSpace = Left(FileName, Len(FileName) - 2)

    

End Function

