Attribute VB_Name = "Module7"
Sub MAJ_Matrce_List_06_02_2024()
Attribute MAJ_Matrce_List_06_02_2024.VB_ProcData.VB_Invoke_Func = " \n14"

Temps_debut = Time

'ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(28, "D").Interior.ColorIndex = 2
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(21, "D") = "Execution En cours"
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(21, "D").Interior.ColorIndex = 4
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(22, "D") = Date & " a " & Time


For u = 1 To Sheets.Count
Onglet_Liste_subs = ActiveWorkbook.Sheets(u).Name
ActiveWorkbook.Sheets(Sheets.Count - 1).Name = "Matrice_" & Onglet_Liste_subs
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(22, "B") = "Matrice_" & Onglet_Liste_subs
Next

    Nb_Lgn_CAS = ActiveWorkbook.Sheets("Matrice_" & Onglet_Liste_subs).Range("A200000").End(xlUp).Row
    If Nb_Lgn_CAS > 1 Then
    ActiveWorkbook.Sheets("Matrice_" & Onglet_Liste_subs).Rows(2 & ":" & ActiveWorkbook.Sheets("Matrice_" & Onglet_Liste_subs).Rows.Count).Delete
    End If

'   '=========================================Identifier les doublons===============================================================
    ActiveWorkbook.Sheets(Onglet_Liste_subs).Activate
    Columns("A:A").Select
    Selection.Copy
    Sheets("Matrice_" & Onglet_Liste_subs).Select
    Columns("A:A").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveSheet.Range("$A$1:$A$200000").RemoveDuplicates Columns:=1, Header:=xlYes
    '=========================================Retourner les listes pour chaque substance========================================
    
    Nb_Lgn_CAS = ActiveWorkbook.Sheets("Matrice_" & Onglet_Liste_subs).Range("A65536").End(xlUp).Row
    'Temps_fin = Time
    For e = 2 To Nb_Lgn_CAS
    'For e = 2 To 10
    b = 4
    CAS_Nbr_Fltr = ActiveWorkbook.Sheets("Matrice_" & Onglet_Liste_subs).Cells(e, "A")
    ActiveWorkbook.Sheets(Onglet_Liste_subs).Range("A1:D65536").AutoFilter Field:=1, Criteria1:=CAS_Nbr_Fltr
    Nb_Lgn_liste = ActiveWorkbook.Sheets(Onglet_Liste_subs).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count
    ActiveWorkbook.Sheets("Matrice_" & Onglet_Liste_subs).Cells(e, "B").FormulaLocal = "=RECHERCHEV(A" & e & ";'" & CStr(Onglet_Liste_subs) & "'!A:D;4;FAUX)"
    Set plage_cas_list = ActiveWorkbook.Sheets(Onglet_Liste_subs).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible)
    Debug.Print Nb_Lgn_liste
    
        For Each cas_list In plage_cas_list
        x3 = cas_list.Row
            If x3 <> 1 Then
            ActiveWorkbook.Sheets("Matrice_" & Onglet_Liste_subs).Cells(e, b) = ActiveWorkbook.Sheets(Onglet_Liste_subs).Cells(x3, "C") & "/"
            ActiveWorkbook.Sheets("Matrice_" & Onglet_Liste_subs).Cells(e, "C") = ActiveWorkbook.Sheets("Matrice_" & Onglet_Liste_subs).Cells(e, "C") _
            & ActiveWorkbook.Sheets("Matrice_" & Onglet_Liste_subs).Cells(e, b)
            b = b + 1
            End If
        Next
    
    ActiveWorkbook.Sheets("Matrice_" & Onglet_Liste_subs).Cells(e, "C") = "[" & ActiveWorkbook.Sheets("Matrice_" & Onglet_Liste_subs).Cells(e, "C") & "]"
    ActiveWorkbook.Sheets("Matrice_" & Onglet_Liste_subs).Cells(e, "C") = Replace(ActiveWorkbook.Sheets("Matrice_" & Onglet_Liste_subs).Cells(e, "C"), "/]", "]")
    Next
    
    
ActiveWorkbook.Sheets(Onglet_Liste_subs).Range("A1:D200000").AutoFilter Field:=1
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(21, "B") = Nb_Lgn_CAS


'For u = 1 To Sheets.Count
'Onglet_Liste_subs = ActiveWorkbook.Sheets(Sheets.Count).Name
'ActiveWorkbook.Sheets(Sheets.Count - 1).Name = "Matrice_" & Onglet_Liste_subs & "[" & Nb_Lgn_CAS & "]"
'Next

ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(22, "B") = "Matrice_" & Onglet_Liste_subs & "[" & Nb_Lgn_CAS & "]"

Temps_fin = Time


ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(21, "D") = "Execution Terminé"
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(21, "D").Interior.ColorIndex = 9
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(22, "E") = Temps_fin - Temps_debut


End Sub

