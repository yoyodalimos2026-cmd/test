Attribute VB_Name = "Module9"
Sub MAJ_Etat_FMD_BDD()

Nb_Ligne_BDD_FMD = ActiveWorkbook.Sheets("BDD_FMD").Range("F65536").End(xlUp).Row
Ligne_debut = ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(15, "B")
Ligne_Fin = ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(16, "B")
'For i = 2 To Nb_Ligne_BDD_FMD
For i = Ligne_debut To Ligne_Fin
    
    
    If ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "I") <> Empty Then
    ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "M").FormulaLocal = "=RECHERCHEV(I" & i & ";'" & "LU_FMD" & "'!A:K;7;FAUX)"
    ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "M").NumberFormat = "@"
    ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "M").NumberFormat = "dd/mm/yyyy"
        
        If Application.IsNA(ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "M")) Then
        ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "M") = "Pas encore ajouté"
        
        ElseIf ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "M") = 0 Then
        ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "M").FormulaLocal = "=RECHERCHEV(I" & i & ";'" & "LU_FMD" & "'!A:K;2;FAUX)"
        
            If ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "M") = "FMD Format IPC" Then
            ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "M") = "FMD Format IPC"
            Else
            ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "M") = "Pas encore analysé"
            ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "M").NumberFormat = "@"
            ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "M").NumberFormat = "dd/mm/yyyy"
            End If
        
        End If
        
    ElseIf ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "I") = Empty Then
    ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "M") = "Pas de FMD"
    End If
Next

    ActiveWorkbook.Sheets("BDD_FMD").Activate
    Columns("M:M").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    ActiveWorkbook.Sheets("BDD_FMD").Cells(1, "I").Select


End Sub
