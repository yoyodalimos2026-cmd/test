Attribute VB_Name = "Module3"
Sub MAJ_Substance_Declaration()
    '****************************************'
    Dim wsPrep As Worksheet
    Dim wsMaj As Worksheet
    Dim wsSubstance As Worksheet
    Dim wsTemplate As Worksheet
    Dim Nb_Ligne_Prepa_Unife As Long
    Dim e As Long
    Dim ct As Long
    Dim b As Long
    Dim Heure_debut3 As Date
    Dim Heure_fin3 As Date
    Dim Durée3 As Double
    Dim wb3 As Workbook
    Dim lastRow As Long
    Dim excludedValues As Variant
    '****************************************'
    Set wsPrep = ActiveWorkbook.Sheets("Preparation_Unife")
    Set wsMaj = ActiveWorkbook.Sheets("MAJ")
    Set wsSubstance = ActiveWorkbook.Sheets("Substance Declaration")
    Set wsTemplate = ActiveWorkbook.Sheets("Template")
    '****************************************'
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    '****************************************'
    wsTemplate.Select
    Cells.Select
    Selection.Copy
    wsSubstance.Select
    Cells.Select
    ActiveSheet.Paste
    '****************************************'
    Nb_Ligne_Prepa_Unife = wsPrep.Range("C1048576").End(xlUp).Row
    'excludedValues = Array("", "No info")
    '****************************************'
    With wsMaj
        .Cells(4, "D").Value = Nb_Ligne_Prepa_Unife
        .Cells(4, "F").Value = Now
        Heure_debut3 = Time
        .Cells(4, "E").Value = "Exécution en cours"
        .Cells(4, "G").Value = ""
        .Cells(4, "E").Interior.ColorIndex = 4
    End With
    '****************************************'
    e = 14
    '****************************************'
    For ct = 2 To Nb_Ligne_Prepa_Unife
        e = e + 1
        With wsPrep
            wsSubstance.Cells(e, "B").Value = .Cells(ct, "G").Value
            wsSubstance.Cells(e, "L").Value = .Cells(ct, "C").Value
            wsSubstance.Cells(e, "M").Value = .Cells(ct, "M").Value
            With wsSubstance.Cells(e, "M")
                .NumberFormat = "@"
                .Value = "'" & wsPrep.Cells(ct, "M").Text
            End With
            wsSubstance.Cells(e, "U").Value = .Cells(ct, "O").Value
            wsSubstance.Cells(e, "W").Value = .Cells(ct, "P").Value
            If Not IsEmpty(.Cells(ct, "O").Value) Then
                wsSubstance.Cells(e, "X").FormulaLocal = "=RECHERCHEV(U" & e & ";" & "'Safety Use'!A:C;" & 2 & ";FAUX)"
                wsSubstance.Cells(e, "V").FormulaLocal = "=RECHERCHEV(U" & e & ";" & "'Substance List'!C:E;" & 3 & ";FAUX)"
                wsSubstance.Cells(e, "T").FormulaLocal = "=RECHERCHEV(U" & e & ";" & "'Substance List'!C:E;" & 2 & ";FAUX)"
                wsSubstance.Cells(e, "Y").Value = .Cells(ct, "N").Value
            Else
                wsSubstance.Cells(e, "Y").Value = "KO"
            End If
        End With
        On Error Resume Next
        If wsSubstance.Cells(e, "B").Value <> "" Then
            For b = 1 To 33
                With wsSubstance.Cells(e, b).Interior
                    .ColorIndex = 22
                    .ColorIndex = 15
                End With
            Next
        End If
    Next ct
    '****************************************'
    With wsSubstance
        lastRow = .Range("C1048576").End(xlUp).Row
        .Range("L:AG").Copy
        .Range("L:AG").PasteSpecial Paste:=xlPasteValues
    End With
    '****************************************'
    wsSubstance.Columns("W:W").NumberFormat = "0%"
    '****************************************'
    With wsMaj
        .Cells(4, "E").Value = "Exécution Terminé"
        .Cells(4, "E").Interior.ColorIndex = 9
        .Cells(4, "G").Value = Now
        Heure_fin3 = Time
        Durée3 = Heure_fin3 - Heure_debut3
        .Cells(4, "H").Value = Durée3
    End With
    '****************************************'
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    '****************************************'
    ActiveWorkbook.Save
End Sub



