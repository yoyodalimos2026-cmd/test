Attribute VB_Name = "Module15"
Sub Indicateur_Iper3()

'=======================Verifie si la colonne "Nb de Jour" existe==============================
For x = 1 To 30

If Cells(1, x) = "Nb de Jour" Then

Columns("J:J").Delete
Columns("J:J").Delete

End If


Next
'==================================================================================================
Columns("J:J").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
Columns("J:J").Select
Selection.NumberFormat = "0"

Nb_Ligne = Range("F65536").End(xlUp).Row

Cells(1, 10) = "Nb de Jour"
Cells(1, 11) = "Etat"

'===================================Mise en forme du rapport=========================================
Cells(Nb_Ligne + 1, 1) = "Rouge"
Cells(Nb_Ligne + 1, 2).FormulaLocal = "=NB.SI(P:P" & ";" & "A" & Nb_Ligne + 1 & ")"
Cells(Nb_Ligne + 1, 3) = "Nombre de références Rouge"
Cells(Nb_Ligne + 1, 1).Interior.ColorIndex = 3
Cells(Nb_Ligne + 1, 2).Interior.ColorIndex = 3

Cells(Nb_Ligne + 2, 1) = "Jaune"
Cells(Nb_Ligne + 2, 2).FormulaLocal = "=NB.SI(P:P" & ";" & "A" & Nb_Ligne + 2 & ")"
Cells(Nb_Ligne + 2, 3) = "Nombre de références Jaune"
Cells(Nb_Ligne + 2, 1).Interior.ColorIndex = 6
Cells(Nb_Ligne + 2, 2).Interior.ColorIndex = 6

Cells(Nb_Ligne + 3, 1) = "Retard Rouge"
Cells(Nb_Ligne + 3, 2).FormulaLocal = "=NB.SI(k:k" & ";" & "A" & Nb_Ligne + 3 & ")"
Cells(Nb_Ligne + 3, 3) = "Nombre de DEQ Rouge qui depassent 10 Jours"
Cells(Nb_Ligne + 3, 1).Interior.ColorIndex = 3
Cells(Nb_Ligne + 3, 2).Interior.ColorIndex = 3

Cells(Nb_Ligne + 4, 1) = "Retard Jaune"
Cells(Nb_Ligne + 4, 2).FormulaLocal = "=NB.SI(k:k" & ";" & "A" & Nb_Ligne + 4 & ")"
Cells(Nb_Ligne + 4, 3) = "Nombre de DEQ Rouge qui depassent 20 Jours"
Cells(Nb_Ligne + 4, 1).Interior.ColorIndex = 3
Cells(Nb_Ligne + 4, 2).Interior.ColorIndex = 3

Cells(Nb_Ligne + 5, 1) = "A Validé"
Cells(Nb_Ligne + 5, 2).FormulaLocal = "=NB.SI(k:k" & ";" & "A" & Nb_Ligne + 5 & ")"
Cells(Nb_Ligne + 5, 3) = "Nombre de DEQ (Rouge>8J,Jaune>15J)"
Cells(Nb_Ligne + 5, 1).Interior.ColorIndex = 45
Cells(Nb_Ligne + 5, 2).Interior.ColorIndex = 45

Cells(Nb_Ligne + 6, 1) = "OK"
Cells(Nb_Ligne + 6, 2).FormulaLocal = "=NB.SI(k:k" & ";" & "A" & Nb_Ligne + 6 & ")"
Cells(Nb_Ligne + 6, 3) = "Nombre de DEQ OK"
Cells(Nb_Ligne + 6, 1).Interior.ColorIndex = 10
Cells(Nb_Ligne + 6, 2).Interior.ColorIndex = 10
'=========================================================================================================


For i = 2 To Nb_Ligne

Cells(i, 10).FormulaLocal = "=AUJOURDHUI()-H" & i

Columns("J:J").Select
Selection.NumberFormat = "0"

Cells(i, 11).FormulaLocal = "=SI(p" & i & "=" & "A" & Nb_Ligne + 2 & ";" & "SI(j" & i & " >10" & ";" & _
"A" & Nb_Ligne + 4 & ";" & "SI(j" & i & ">8" & ";" & "A" & Nb_Ligne + 5 & ";" & "A" & Nb_Ligne + 6 & "))" & ";" & "SI(p" & i & _
"=" & "A" & Nb_Ligne + 1 & ";" & "SI(j" & i & ">20" & ";" & "A" & Nb_Ligne + 3 & ";" _
& "SI(j" & i & ">15" & ";" & "A" & Nb_Ligne + 5 & ";" & "A" & Nb_Ligne + 6 & "))))"


Next


'====================================Mise en forme==============================================================

Columns("K:K").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""A Validé"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 49407
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = True
    
    Columns("K:K").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""Retard Rouge"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = True
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""Retard Jaune"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = True
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""OK"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 5287936
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = True
    ActiveWindow.SmallScroll Down:=-3
    
 '============================Ecadrer le rapport====================================================
 
 Range("A61:C66").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("E66").Select
    


End Sub

Sub Macro16()
'
' Macro16 Macro
'

'
    Range("A61:C66").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("E66").Select
End Sub


