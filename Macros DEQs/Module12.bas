Attribute VB_Name = "Module12"
Sub New_tab_Verif()
'
'
'
Dim wb As Workbook
Dim ws As Worksheet
Dim Message, Alerte, DEQ, RefInit, RefEquiv, FabInit, FabEquiv As String
Dim lin As Integer



For lin = 2 To Range("A65536").End(xlUp).Row

DEQ = ActiveWorkbook.Sheets(1).Cells(lin, 1).Value
RefInit = ActiveWorkbook.Sheets(1).Cells(lin, 5).Value
FabInit = ActiveWorkbook.Sheets(1).Cells(lin, 6).Value
RefEquiv = ActiveWorkbook.Sheets(1).Cells(lin, 11).Value
FabEquiv = ActiveWorkbook.Sheets(1).Cells(lin, 12).Value
Message = ""
Alerte = ""

Set wb = Workbooks.Open(ActiveWorkbook.Path & "\" & DEQ & "\COMP.xlsx")
Set ws = wb.Worksheets(1)
Application.Wait Now + TimeValue("0:00:01")

' Tester le numéro de DEQ

If ws.Cells(1, 4).Value <> DEQ Then
Message = " N° DEQ /"
End If


If KillSpace(ws.Cells(16, 3).Value) <> Application.Trim(RefInit) And Application.Trim(ws.Cells(16, 3).Value) <> Application.Trim(RefInit) Then

Message = Message & " RefINIT /"

End If
If Application.Trim(ws.Cells(18, 3).Value) <> Application.Trim(FabInit) Then

Alerte = Alerte & " FabInit /"
End If


If RefEquiv = "SANS" Then

    If ws.Cells(16, 3).MergeArea.Columns.Count <> 2 Then
        Message = "ERREUR !! il faut enlever l'équivalent du comp !"
        'MsgBox (ws.Cells(16, 2).Value)
    End If
    Else
        If ws.Cells(16, 3).MergeArea.Columns.Count <> 1 Then
            Message = "ERREUR !! Il manque la colonne équivalent dans le Comp!"
        Else
            If Application.Trim(ws.Cells(18, 4).Value) <> Application.Trim(FabEquiv) Then
                Alerte = Alerte & " FabEquiv /"
            End If
            
                
           If KillSpace(ws.Cells(16, 4).Value) <> Application.Trim(RefEquiv) <> Application.Trim(RefEquiv) And KillSpace(KillSpace(ws.Cells(16, 4).Value)) <> Application.Trim(RefEquiv) And Application.Trim(ws.Cells(16, 4).Value) <> Application.Trim(RefEquiv) Then
               Message = Message & " RefEQUIV /"
            End If
        End If
End If

wb.Close
ActiveWorkbook.Sheets(1).Cells(lin, 33) = Message
If Message <> "" Then
ActiveWorkbook.Sheets(1).Cells(lin, 33).Hyperlinks.Add Anchor:=Cells(lin, 33), Address:=ActiveWorkbook.Path & "\" & DEQ & "\COMP.xlsx"
End If
ActiveWorkbook.Sheets(1).Cells(lin, 34) = Alerte
Application.Wait Now + TimeValue("0:00:01")

Next lin
MsgBox ("OK")
End Sub

 
Function KillSpace(ByVal FileName As String) As String

     KillSpace = Left(FileName, Len(FileName) - 2)
     
End Function


