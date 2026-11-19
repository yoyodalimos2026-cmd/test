Attribute VB_Name = "Module8"
Sub Ajout_FMD()

Nb_Ligne_BDD_FMD = ActiveWorkbook.Sheets("BDD_FMD").Range("F65536").End(xlUp).Row
'Nb_Ligne_BDD_FMD = ActiveWorkbook.Sheets("BDD_FMD").UsedRange.Rows.Count
Nb_Lignes_Data = ActiveWorkbook.Sheets("LU_FMD").Range("A65536").End(xlUp).Row
lgn = Nb_Lignes_Data + 1
Ligne_debut = ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(15, "B")
Ligne_Fin = ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(16, "B")
For t = Ligne_debut To Ligne_Fin
'For t = 2 To Nb_Ligne_BDD_FMD

    ActiveWorkbook.Sheets("LU_FMD").Range("A" & t & ":K" & Ligne_Fin).Borders.LineStyle = xlContinuous
    
    If ActiveWorkbook.Sheets("BDD_FMD").Cells(t, "M") = "Pas encore ajouté" Then
    'InStr(ActiveWorkbook.Sheets("BDD_FMD").Cells(t, "J"), "FMD introuvable sous le dossier FMD") <> 0 Or _
    'ActiveWorkbook.Sheets("BDD_FMD").Cells(t, "J") <> "Pas de FMD" Then
    
    'ActiveWorkbook.Sheets("LU_FMD").Cells(lgn, "A").Value2 = ActiveWorkbook.Sheets("BDD_FMD").Cells(t, "I").Value2
    ActiveWorkbook.Sheets("BDD_FMD").Cells(t, "I").Copy
    ActiveWorkbook.Sheets("LU_FMD").Cells(lgn, "A").PasteSpecial Paste:=xlPasteFormulas
    ActiveWorkbook.Sheets("LU_FMD").Cells(lgn, "A").PasteSpecial Paste:=xlPasteFormats
    ActiveWorkbook.Sheets("BDD_FMD").Cells(t, "J").Copy
    ActiveWorkbook.Sheets("LU_FMD").Cells(lgn, "B").PasteSpecial Paste:=xlPasteFormulas
    ActiveWorkbook.Sheets("LU_FMD").Cells(lgn, "B").PasteSpecial Paste:=xlPasteFormats
    
        If InStr(ActiveWorkbook.Sheets("BDD_FMD").Cells(t, "I"), "MDS") = 0 Then
        ActiveWorkbook.Sheets("LU_FMD").Cells(lgn, "C").Value = Mid(ActiveWorkbook.Sheets("BDD_FMD").Cells(t, "K").Value, 1, 4)
        ActiveWorkbook.Sheets("LU_FMD").Cells(lgn, "C").PasteSpecial Paste:=xlPasteFormats
        ElseIf InStr(ActiveWorkbook.Sheets("BDD_FMD").Cells(t, "I"), "MDS") <> 0 Then
        ActiveWorkbook.Sheets("LU_FMD").Cells(lgn, "C").Value = ActiveWorkbook.Sheets("BDD_FMD").Cells(t, "K").Value
        'ActiveWorkbook.Sheets("LU_FMD").Cells(lgn, "C").PasteSpecial Paste:=xlPasteFormats
        End If
        
    lgn = lgn + 1
    ActiveWorkbook.Sheets("BDD_FMD").Cells(t, "M") = "Pas encore analysé"
    End If

Next

    ActiveWorkbook.Sheets("LU_FMD").Activate
    Columns("A:A").Select
    ActiveSheet.Range("$A$1:$K$20000").RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5, 6 _
        , 7, 8, 9, 10, 11), Header:=xlYes
    ActiveWorkbook.Sheets("Macro_MAJ_Document").Activate
    ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(1, 1).Select
    
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(28, "B") = ActiveWorkbook.Sheets("LU_FMD").Range("A65536").End(xlUp).Row
End Sub
