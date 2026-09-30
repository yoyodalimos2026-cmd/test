Attribute VB_Name = "Module14"
Sub Delete_Doublon_BOM()


Nb_Lgn_BOM_a_analyser = ActiveWorkbook.Sheets("BOM_a_analyser").Range("D65536").End(xlUp).Row

For c = 2 To Nb_Lgn_BOM_a_analyser
'For c = 2 To 200
    ActiveWorkbook.Sheets("BOM_a_analyser").Range("A" & c & ":P" & Nb_Lgn_BOM_a_analyser).Borders.LineStyle = xlContinuous
    ActiveWorkbook.Sheets("BOM_a_analyser").Cells(c, "E").Interior.ColorIndex = 2
    ActiveWorkbook.Sheets("BOM_a_analyser").Cells(c, "E") = ""
    ActiveWorkbook.Sheets("BOM_a_analyser").Cells(c, "F").Interior.ColorIndex = 2
    ActiveWorkbook.Sheets("BOM_a_analyser").Cells(c, "F") = ""
    ActiveWorkbook.Sheets("BOM_a_analyser").Cells(c, "G").Interior.ColorIndex = 2
    ActiveWorkbook.Sheets("BOM_a_analyser").Cells(c, "G") = ""
Next

'======================================================Eliminer les doublons=========================================================
    ActiveWorkbook.Sheets("BOM_a_analyser").Activate
    Columns("D:D").Select
    ActiveWorkbook.Sheets("BOM_a_analyser").Range("$A$1:$P$1038475").RemoveDuplicates Columns:=Array(1, 2, 3, 4, _
        5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16), Header:=xlYes
'======================================================'======================================================'=======================

For d = 2 To Nb_Lgn_BOM_a_analyser
    For y = 8 To 16
        'If ActiveWorkbook.Sheets("BOM_a_analyser").Cells(c, y) <> "NA" Then
        ActiveWorkbook.Sheets("BOM_a_analyser").Cells(d, y) = Replace(ActiveWorkbook.Sheets("BOM_a_analyser").Cells(d, y), " ", "")
        'End If
    Next
Next

'======================================================Eliminer les doublons=========================================================
    ActiveWorkbook.Sheets("BOM_a_analyser").Activate
    Columns("D:D").Select
    ActiveWorkbook.Sheets("BOM_a_analyser").Range("$A$1:$P$1038475").RemoveDuplicates Columns:=Array(1, 2, 3, 4, _
        5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16), Header:=xlYes
'======================================================'======================================================'=======================

Nb_Lgn_BOM_a_analyser = ActiveWorkbook.Sheets("BOM_a_analyser").Range("D65536").End(xlUp).Row
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(28, "I") = Nb_Lgn_BOM_a_analyser

ActiveWorkbook.Save

End Sub
