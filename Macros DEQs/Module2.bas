Attribute VB_Name = "Module2"
Sub SILICON()

    'Sub CompSilicon()
'
    Dim shcomp As Worksheet
    Dim Chemin1, Chemin2, fichier1, fichier2, Cellfin As String
    Dim nbretemp, nbrepar, nbredim As Integer
    Dim Suppressed(50) As String
    
    Suppressed(1) = "Speed Grade"
    Suppressed(2) = "Temperature Flag"
    Suppressed(3) = "Package Description"
    Suppressed(4) = "Lead Shape"
    Suppressed(5) = "PCB"
    Suppressed(6) = "Package Outline"
    Suppressed(7) = "Jedec"
    Suppressed(8) = "Propagation Delay Test Condition (pF)"
    Suppressed(9) = "Seated Plane Height (mm)"
    Suppressed(10) = "Maximum Propagation Delay Time @ Maximum CL (ns)"
    Suppressed(11) = "Typical Quiescent Current (uA)"
    Suppressed(12) = "Supplier Temperature Grade"
    Suppressed(13) = "Quantity per Packaging (Pieces)"
    Suppressed(14) = "Basic Package Type"
    Suppressed(15) = "Standard Package Name"
    Suppressed(16) = "Special Features"
    Suppressed(17) = "Maximum High Level Output Current (mA)"
    Suppressed(18) = "Maximum Low Level Output Current (mA)"
    Suppressed(19) = "Number of Selection Inputs per Element"
    Suppressed(20) = "Number of Count Input Enables"
    Suppressed(21) = "Parallel Enable Input"
    Suppressed(22) = "Number of Stages"
    Suppressed(23) = "Maximum Operating Supply Voltage (V)"
    Suppressed(24) = "Maximum Single Supply Voltage (V)"
    Suppressed(25) = "Seated Plane Height (mm)"
    Suppressed(26) = "Maximum Propagation Delay Time @ Maximum CL (ns)"
    Suppressed(27) = "Pin Pitch (mm)"
    Suppressed(28) = "Typical Operating Supply Voltage (V)"
    Suppressed(29) = " Number of Output Enables per Element"
    Suppressed(30) = "Manufacturer Type"
    Suppressed(31) = "Typical Dual Supply Voltage (V)"
    Suppressed(32) = "Shut Down Support"
    Suppressed(33) = "Rail to Rail"
    Suppressed(34) = "Number of Inputs per Chip"
    Suppressed(35) = "Number of Channels per Chip"
    Suppressed(36) = "Number of Outputs per Chip"
    Suppressed(37) = " Number of Output Enables per Element"
    Suppressed(38) = "Number of Elements per Chip"
    Suppressed(39) = "Number of Input Enables per Chip"
    Suppressed(40) = "Number of Output Enables per Chip"
    Suppressed(41) = "Rail to Rail"
    Suppressed(42) = "Strobe Capability"
    Suppressed(43) = "Typical Single Supply Voltage (V)"
    Suppressed(44) = "Minimum Operating Supply Voltage (V)"
    Suppressed(45) = "Minimum Single Supply Voltage (V)"
Chemin1 = ActiveWorkbook.Path
NumDEQ = Right(Chemin1, 14)
fichier1 = ActiveWorkbook.Name

Chemin2 = "U:\Disque_Hassen_Yassine\Validation des équivalents\Trame.xls"
Workbooks.Open Chemin2, 0, ReadOnly:=False
fichier2 = ActiveWorkbook.Name
Workbooks(fichier2).Sheets("Comparatif").Copy Before:=Workbooks(fichier1).Sheets(1)

Workbooks(fichier2).Close
Workbooks(fichier1).Activate

Cellfin = "C"

'information P/N + Mfr + Description

Workbooks(fichier1).Sheets("Comparatif").Cells(16, 2).Value = Workbooks(fichier1).Sheets("Summary").Cells(7, 2).Value
Workbooks(fichier1).Sheets("Comparatif").Cells(16, 3).Value = Workbooks(fichier1).Sheets("Summary").Cells(7, 3).Value
Workbooks(fichier1).Sheets("Comparatif").Cells(17, 2).Value = Workbooks(fichier1).Sheets("Summary").Cells(9, 2).Value
Workbooks(fichier1).Sheets("Comparatif").Cells(17, 3).Value = Workbooks(fichier1).Sheets("Summary").Cells(9, 3).Value
Workbooks(fichier1).Sheets("Comparatif").Cells(18, 2).Value = Workbooks(fichier1).Sheets("Summary").Cells(8, 2).Value
Workbooks(fichier1).Sheets("Comparatif").Cells(18, 3).Value = Workbooks(fichier1).Sheets("Summary").Cells(8, 3).Value

'3 eme composant à comparer?
If Workbooks(fichier1).Sheets("Summary").Cells(7, 4).Value <> "" Then
Cellfin = "D"
Workbooks(fichier1).Sheets("Comparatif").Cells(16, 4).Value = Workbooks(fichier1).Sheets("Summary").Cells(7, 4).Value
Workbooks(fichier1).Sheets("Comparatif").Cells(17, 4).Value = Workbooks(fichier1).Sheets("Summary").Cells(9, 4).Value
Workbooks(fichier1).Sheets("Comparatif").Cells(18, 4).Value = Workbooks(fichier1).Sheets("Summary").Cells(8, 4).Value

End If

'4 eme composant à comparer?
If Workbooks(fichier1).Sheets("Summary").Cells(7, 5).Value <> "" Then
Cellfin = "E"
Workbooks(fichier1).Sheets("Comparatif").Cells(16, 5).Value = Workbooks(fichier1).Sheets("Summary").Cells(7, 5).Value
Workbooks(fichier1).Sheets("Comparatif").Cells(17, 5).Value = Workbooks(fichier1).Sheets("Summary").Cells(9, 5).Value
Workbooks(fichier1).Sheets("Comparatif").Cells(18, 5).Value = Workbooks(fichier1).Sheets("Summary").Cells(8, 5).Value


End If

'Num DEQ
If Cellfin = "C" Then
NumDEQ = Right(Chemin1, 14)

'Num Multisourcing
Else
NumDEQ = Right(Chemin1, 5)

End If

Workbooks(fichier1).Sheets("Comparatif").Cells(1, 3).Value = NumDEQ


'Ajout des champs parametric
For i = 1 To Worksheets.Count
    If Worksheets(i).Name = "Parametric" Then
nbrepar = Workbooks(fichier1).Sheets("Parametric").Columns(1).Cells.Find("*", , , , , xlPrevious).Row - 7
Workbooks(fichier1).Sheets("Comparatif").Range("A19:" & Cellfin & 18 + nbrepar) = Workbooks(fichier1).Sheets("Parametric").Range("A8:" & Cellfin & nbrepar + 7).Value
End If
Next i
 
'Ajout des champs Package
For i = 1 To Worksheets.Count
    If Worksheets(i).Name = "Package" Then
     nbredim = Workbooks(fichier1).Sheets("Package").Columns(1).Cells.Find("*", , , , , xlPrevious).Row - 7
     Workbooks(fichier1).Sheets("Comparatif").Range("A" & 19 + nbrepar & ":" & Cellfin & 18 + nbrepar + nbredim) = Workbooks(fichier1).Sheets("Package").Range("A8:" & Cellfin & nbredim + 7).Value

    End If
Next i

'Température Min et Max

For i = 1 To Workbooks(fichier1).Sheets("Comparatif").Columns(1).Cells.Find("*", , , , , xlPrevious).Row
If Workbooks(fichier1).Sheets("Comparatif").Cells(i, 1).Value = "Minimum Operating Temperature (°C)" Then
Workbooks(fichier1).Sheets("Comparatif").Cells(i, 1).EntireRow.Cut Destination:=Workbooks(fichier1).Sheets("Comparatif").Cells(19 + nbrepar + nbredim, 1)
Workbooks(fichier1).Sheets("Comparatif").Cells(i, 1).EntireRow.Delete
Exit For
End If
Next i

For i = 1 To Workbooks(fichier1).Sheets("Comparatif").Columns(1).Cells.Find("*", , , , , xlPrevious).Row
If Workbooks(fichier1).Sheets("Comparatif").Cells(i, 1).Value = "Maximum Operating Temperature (°C)" Then
Workbooks(fichier1).Sheets("Comparatif").Cells(i, 1).EntireRow.Cut Destination:=Workbooks(fichier1).Sheets("Comparatif").Cells(19 + nbrepar + nbredim, 1)
Workbooks(fichier1).Sheets("Comparatif").Cells(i, 1).EntireRow.Delete
Exit For
End If
Next i





'Ajout ligne RoHS
Workbooks(fichier1).Sheets("Comparatif").Cells(19 + nbrepar + nbredim, 1).Value = "RoHS Compliance"

For i = 1 To Workbooks(fichier1).Sheets("Summary").Columns(1).Cells.Find("*", , , , , xlPrevious).Row
If Workbooks(fichier1).Sheets("Summary").Cells(i, 1) = "EU RoHS" Then
Workbooks(fichier1).Sheets("Comparatif").Cells(19 + nbrepar + nbredim, 1).Value = "RoHS Compliance"
Workbooks(fichier1).Sheets("Comparatif").Cells(19 + nbrepar + nbredim, 2) = Workbooks(fichier1).Sheets("Summary").Cells(i, 2)
Workbooks(fichier1).Sheets("Comparatif").Cells(19 + nbrepar + nbredim, 3) = Workbooks(fichier1).Sheets("Summary").Cells(i, 3)
If Workbooks(fichier1).Sheets("Summary").Cells(7, 4).Value <> "" Then
Workbooks(fichier1).Sheets("Comparatif").Cells(19 + nbrepar + nbredim, 4) = Workbooks(fichier1).Sheets("Summary").Cells(i, 4)
End If
If Workbooks(fichier1).Sheets("Summary").Cells(7, 5).Value <> "" Then
Workbooks(fichier1).Sheets("Comparatif").Cells(19 + nbrepar + nbredim, 5) = Workbooks(fichier1).Sheets("Summary").Cells(i, 5)
End If

Exit For
End If
Next i



'Supprimer les lignes vides ou non renseignées
i = 19
While Workbooks(fichier1).Sheets("Comparatif").Cells(i, 1).Value <> ""
If Workbooks(fichier1).Sheets("Comparatif").Cells(i, 2).Value = "" _
Or Workbooks(fichier1).Sheets("Comparatif").Cells(i, 2).Value = "N/R" Then

  If Workbooks(fichier1).Sheets("Comparatif").Cells(i, 3).Value = "" _
  Or Workbooks(fichier1).Sheets("Comparatif").Cells(i, 3).Value = "N/R" Then
  
  Workbooks(fichier1).Sheets("Comparatif").Cells(i, 1).EntireRow.Delete
  i = i - 1
  End If
End If

i = i + 1
Wend

'Power Supply
'For i = 1 To Workbooks(fichier1).Sheets("Comparatif").Columns(1).Cells.Find("*", , , , , xlPrevious).Row
'If Workbooks(fichier1).Sheets("Comparatif").Cells(i, 1).Value = "Minimum Operating Supply Voltage (V)" _
'Or Workbooks(fichier1).Sheets("Comparatif").Cells(i, 1).Value = "Minimum Single Supply Voltage (V)" Then
'Workbooks(fichier1).Sheets("Comparatif").Cells(i, 1).EntireRow.Delete
Workbooks(fichier1).Sheets("Comparatif").Cells(19, 1).EntireRow.Insert
Workbooks(fichier1).Sheets("Comparatif").Cells(19, 1).Value = "Power Supply (V):"
'Exit For
'End If
'Next i

'Supprimer les lignes inutiles du tableau

For i = Workbooks(fichier1).Sheets("Comparatif").Columns(1).Cells.Find("*", , , , , xlPrevious).Row To 19 Step -1

For j = 1 To 43
If Workbooks(fichier1).Sheets("Comparatif").Cells(i, 1).Value = Suppressed(j) Then
Workbooks(fichier1).Sheets("Comparatif").Cells(i, 1).EntireRow.Delete
End If

Next j
Next i

'Supprimer les feuilles supplémentaire
For i = Workbooks(fichier1).Sheets.Count To 2 Step -1

Application.DisplayAlerts = False
Workbooks(fichier1).Sheets(i).Delete
Application.DisplayAlerts = True

Next i

'*************************'Supprimer les dimensions si Boitier défini
'For k = 1 To Worksheets.Count
   ' If Worksheets(i).Name = "Package" Then

For i = 1 To Workbooks(fichier1).Sheets("Comparatif").Columns(1).Cells.Find("*", , , , , xlPrevious).Row
  
  If Workbooks(fichier1).Sheets("Comparatif").Cells(i, 1).Value = "Supplier Package" Then

 For j = 1 To Workbooks(fichier1).Sheets("Comparatif").Columns(1).Cells.Find("*", , , , , xlPrevious).Row
      If Workbooks(fichier1).Sheets("Comparatif").Cells(j, 1).Value = "Package Length (mm)" Then
      Workbooks(fichier1).Sheets("Comparatif").Cells(j, 1).EntireRow.Delete
      If Workbooks(fichier1).Sheets("Comparatif").Cells(j, 1).Value = "Package Width (mm)" Then
      Workbooks(fichier1).Sheets("Comparatif").Cells(j, 1).EntireRow.Delete
      If Workbooks(fichier1).Sheets("Comparatif").Cells(j, 1).Value = "Package Height (mm)" Then
      Workbooks(fichier1).Sheets("Comparatif").Cells(j, 1).EntireRow.Delete
      If Workbooks(fichier1).Sheets("Comparatif").Cells(j, 1).Value = "Package Diameter (mm)" Then
      Workbooks(fichier1).Sheets("Comparatif").Cells(j, 1).EntireRow.Delete
      If Workbooks(fichier1).Sheets("Comparatif").Cells(j, 1).Value = "Pin Pitch (mm)" Then
      Workbooks(fichier1).Sheets("Comparatif").Cells(j, 1).EntireRow.Delete
      
      End If
      End If
      End If
      End If
      End If
  Next j
  End If
Next i
'End If
 '  Next k
'********************************************

'Supprimer les feuilles supplémentaire
For i = Workbooks(fichier1).Sheets.Count To 2 Step -1

Application.DisplayAlerts = False
Workbooks(fichier1).Sheets(i).Delete
Application.DisplayAlerts = True

Next i

'Mise en forme---------------------------------------------------------------------------------------
' Mise en forme couleur pour différence
fin = Workbooks(fichier1).Sheets("Comparatif").Columns(1).Cells.Find("*", , , , , xlPrevious).Row

If Cellfin = "C" Then ' Cas d'un DEQs

For i = 18 To fin
    Range("C" & i).Select
    Selection.FormatConditions.Delete
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, _
        Formula1:="=B" & i
    Selection.FormatConditions(1).Interior.ColorIndex = 6
    
    Range("B" & i).Select
    Selection.FormatConditions.Delete
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, _
        Formula1:="=C" & i
    Selection.FormatConditions(1).Interior.ColorIndex = 6
    Next i

Else ' Cas d'un Multisourcing

If Cellfin = "D" Then
For i = 18 To fin
    Range("C" & i).Select
    Selection.FormatConditions.Delete
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, _
        Formula1:="=B" & i
    Selection.FormatConditions(1).Interior.ColorIndex = 6
    
    Range("D" & i).Select
    Selection.FormatConditions.Delete
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, _
        Formula1:="=B" & i
    Selection.FormatConditions(1).Interior.ColorIndex = 6
    Next i

Else

If Cellfin = "E" Then

For i = 18 To fin
    Range("C" & i).Select
    Selection.FormatConditions.Delete
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, _
        Formula1:="=B" & i
    Selection.FormatConditions(1).Interior.ColorIndex = 6
    
    Range("D" & i).Select
    Selection.FormatConditions.Delete
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, _
        Formula1:="=B" & i
    Selection.FormatConditions(1).Interior.ColorIndex = 6
    
    
    Range("E" & i).Select
    Selection.FormatConditions.Delete
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, _
        Formula1:="=B" & i
    Selection.FormatConditions(1).Interior.ColorIndex = 6
    
    Next i
End If

End If

End If
' Bordure
    Range("A16:" & Cellfin & fin).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    Range("A16:A" & fin).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With

' Renvoie à la ligne
    Cells.Select
    With Selection
        .WrapText = True
        .Orientation = 0
        .ShrinkToFit = False
    End With
' Mise en page
    ActiveSheet.PageSetup.PrintArea = ""
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = "Préparé par ACTIA"
        .CenterFooter = ""
        .RightFooter = "Page &P de &N"
        .LeftMargin = Application.InchesToPoints(0.393700787401575)
        .RightMargin = Application.InchesToPoints(0.393700787401575)
        .TopMargin = Application.InchesToPoints(0.590551181102362)
        .BottomMargin = Application.InchesToPoints(0.590551181102362)
        .HeaderMargin = Application.InchesToPoints(0.31496062992126)
        .FooterMargin = Application.InchesToPoints(0.31496062992126)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
'        .PrintQuality = 600
        .CenterHorizontally = True
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        '.PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 100
    End With
ActiveWorkbook.Save

MsgBox ("Done")

End Sub



