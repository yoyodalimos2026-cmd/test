Attribute VB_Name = "Module17"
Sub Transfert_data()

'==========================================Remplacer les espaces==============================================================
    ActiveWorkbook.Sheets("BOM_a_analyser").Activate
    Columns("H:R").Select
    Selection.Replace What:=" ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

'=======================================================================================================================
'==========================================Remplacer les "."==============================================================
    ActiveWorkbook.Sheets("BOM_a_analyser").Activate
    Columns("H:R").Select
    Selection.Replace What:=".", Replacement:=",", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

'=======================================================================================================================
Nb_Lignes_BOM = ActiveWorkbook.Sheets("BOM_a_analyser").Range("D200000").End(xlUp).Row
For g = 2 To Nb_Lignes_BOM


ActiveWorkbook.Sheets("BOM_a_analyser").Range("H" & g & ":R" & Nb_Lignes_BOM).Borders.LineStyle = xlContinuous
    
    For f = 8 To 18
    
        ActiveWorkbook.Sheets("BOM_a_analyser").Cells(g, f).Interior.ColorIndex = 2
        
        Len_Cell_CAS = Len(ActiveWorkbook.Sheets("BOM_a_analyser").Cells(g, f))
        Len_Cell_CAS_x = Len(Replace(ActiveWorkbook.Sheets("BOM_a_analyser").Cells(g, f), "-", ""))
        Nb_CAS = (Len_Cell_CAS - Len_Cell_CAS_x) / 2 'Clacul le nombre de CAS Number par cellule

        Len_Cell_percent = Len(ActiveWorkbook.Sheets("BOM_a_analyser").Cells(g, f))
        Len_Cell_percent_x = Len(Replace(ActiveWorkbook.Sheets("BOM_a_analyser").Cells(g, f), "%", ""))
        Nb_percent = (Len_Cell - Len_Cell_x) / 2    'Clacul le nombre de % par cellule
        
        
        If Nb_CAS <> Nb_percent Then
        ActiveWorkbook.Sheets("BOM_a_analyser").Cells(g, f).Interior.ColorIndex = 3
        
        
        
        End If


    Next f
Next g


End Sub
