Attribute VB_Name = "Module13"
Sub Init_BOM_14_03_2024()
Attribute Init_BOM_14_03_2024.VB_ProcData.VB_Invoke_Func = " \n14"

    ActiveWorkbook.Sheets("BOM_a_analyser").Activate
    Columns("H:R").Select
    Selection.Replace What:=" ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        
        
    ActiveWorkbook.Sheets("BOM_a_analyser").Activate
    Columns("D:D").Select
    ActiveWorkbook.Sheets("BOM_a_analyser").Range("$A$1:$R$1038475").RemoveDuplicates Columns:=Array(1, 2, 3, 4, _
        5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18), Header:=xlYes
        

Nb_Lignes_BOM = ActiveWorkbook.Sheets("BOM_a_analyser").Range("D200000").End(xlUp).Row
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(28, "I") = Nb_Lignes_BOM

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
