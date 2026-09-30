Attribute VB_Name = "Module1"
Sub EnregistrementMDS()

Dim ws1, ws2 As Worksheet
Dim LineNB As Double
Dim str1 As String

Set ws1 = ActiveWorkbook.Sheets("Input")
Set ws2 = ActiveWorkbook.Sheets("Home")

Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oMDS = oFSO.GetFolder(ws2.Cells(5, "D"))
Set ws1 = ActiveWorkbook.Sheets("Input")

LineNB = ws1.Range("A65536").End(xlUp).Row

ws2.Cells(6, "D") = LineNB

For i = 2 To LineNB

    ws1.Cells(i, "G") = ws1.Cells(i, "A") & "_" & ws1.Cells(i, "B") & "_" & ws1.Cells(i, "C") & "_MDS_" & ws1.Cells(i, "E")
    str1 = ws1.Cells(i, "D")
    ws1.Cells(i, "H") = ExtractID(str1)
    
Next i

x = 2

For Each oDoc In oMDS.Files
    TMP = Split(oDoc.Name, "_")(1)
    TMP = Mid(TMP, 1, Len(TMP) - 4)
    '''ws1.Cells(x, "H") = TMP'''
    x = x + 1
    For y = 2 To LineNB
        If StrComp(ws1.Cells(y, "H"), TMP) = 0 Then
            oDoc.Name = ws1.Cells(y, "G") & ".pdf"
        End If
    Next y
Next

ws2.Cells(7, "D") = x - 2

End Sub
Public Function ExtractID(SourceString As String)

Fposition = InStr(1, SourceString, "/") + 1
Lposition = InStrRev(SourceString, "/")

ExtractID = Mid(SourceString, Fposition, Lposition - Fposition)

End Function

