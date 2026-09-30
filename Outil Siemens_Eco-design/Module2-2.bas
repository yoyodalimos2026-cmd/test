Attribute VB_Name = "Module2"
Sub Preparation_Unife()
    Dim wsPrep As Worksheet, wsPrepUnife As Worksheet, wsMaj As Worksheet, wsListe As Worksheet
    Dim Nb_Ligne_Prepa As Long, Nb_Ligne_Prepa_Unife As Long, Nb_Ligne_Liste As Long
    Dim d As Long, t As Long, c2 As Long, z As Long, l3 As Long
    Dim Heure_debut1 As Date, Heure_fin1 As Date, Durée1 As Double
    Dim dataPrep As Variant, dataPrepUnife As Variant
    Dim p As Long
    '****************************************'
    Set wsPrep = ActiveWorkbook.Sheets("Preparation")
    Set wsPrepUnife = ActiveWorkbook.Sheets("Preparation_Unife")
    Set wsMaj = ActiveWorkbook.Sheets("MAJ")
    'Set wsListe = ActiveWorkbook.Sheets("Liste des code articles ss dbl")
    '****************************************'
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    '****************************************'
    Nb_Ligne_Prepa = wsPrep.Cells(wsPrep.Rows.Count, "C").End(xlUp).Row
    Nb_Ligne_Prepa_Unife = wsPrepUnife.Cells(wsPrepUnife.Rows.Count, "C").End(xlUp).Row
    '****************************************'
    wsMaj.Cells(3, "D") = Nb_Ligne_Prepa
    wsMaj.Cells(3, "F") = Date & " à " & Time
    Heure_debut1 = Time
    wsMaj.Cells(3, "E") = "Exécution en cours"
    wsMaj.Cells(3, "G") = ""
    wsMaj.Cells(3, "E").Interior.ColorIndex = 4
    '****************************************'
    wsPrepUnife.Rows("2:" & wsPrepUnife.Rows.Count).ClearContents
    d = 2
    dataPrep = wsPrep.Range("A2:AZ" & Nb_Ligne_Prepa).Value
    '****************************************'
    For t = 1 To UBound(dataPrep, 1)
        '****************************************'
        If InStr(dataPrep(t, 15), "%") <> 0 Then
            '****************************************'
            For c = 1 To 13
                With wsPrepUnife.Cells(d, c)
                    .NumberFormat = "@"
                    .Value = "'" & wsPrep.Cells(t + 1, c).Text
                End With
            Next c
            '****************************************'
            For c2 = 15 To UBound(dataPrep, 2)
                '****************************************'
                If InStr(1, dataPrep(t, c2), "%") Then
                    d = d + 1
                    With wsPrepUnife.Cells(d, "N")
                        .NumberFormat = "@"
                        .Value = "'" & wsPrep.Cells(t + 1, 14).Text
                    End With
                    With wsPrepUnife.Cells(d, "O")
                        .NumberFormat = "@"
                        .Value = "'" & wsPrep.Cells(t + 1, c2).Text
                    End With

                End If
                '****************************************'
            Next c2
            '****************************************'
            d = d + 1
            '****************************************'
        End If
        '****************************************'
    Next t
    '****************************************'
    'Nb_Ligne_Liste = wsListe.Cells(wsListe.Rows.Count, "C").End(xlUp).Row
    'Nb_Ligne_Prepa = wsPrepUnife.Cells(wsPrepUnife.Rows.Count, "C").End(xlUp).Row
    'wsListe.Columns("N:N").ClearContents
    '****************************************'

    'Dim matchRow As Variant
    'Dim srcVal As Variant
    'wsListe.Columns("N:N").ClearContents
    'For z = 2 To Nb_Ligne_Prepa
        'srcVal = wsListe.Cells(z, "C").Value
        'If Len(Trim(srcVal & "")) = 0 Then
            'wsListe.Cells(z, "N").Value = ""
        'Else
            'matchRow = Application.Match(srcVal, wsPrepUnife.Columns("C"), 0)
            'If Not IsError(matchRow) Then
                ' Match found; copy displayed text from column E (3rd column of C:O)
                'With wsListe.Cells(z, "N")
                    '.NumberFormat = "@"
                    '.Value = "'" & wsPrepUnife.Cells(matchRow, "E").Text
                'End With
            'Else
                'wsListe.Cells(z, "N").Value = ""
            'End If
        'End If
    'Next z
    '****************************************'
    'For z = 2 To Nb_Ligne_Prepa
        'If Not IsEmpty(wsListe.Cells(z, "N").Value) Then
            'If wsListe.Cells(z, "N").Value = wsListe.Cells(z, "E").Value Then
                'wsListe.Cells(z, "N") = "Oui"
            'End If
        'End If
   ' Next z
    '****************************************'
    'wsListe.Cells(1, "N") = "Présence dans l'onglet Preparation nUnife"
    'p = wsPrepUnife.Cells(wsPrepUnife.Rows.Count, "C").End(xlUp).Row
    'For l3 = 2 To Nb_Ligne_Prepa
        'If wsListe.Cells(l3, "N").Value <> "Oui" Then
            'p = p + 1
            'For c3 = 1 To 13
                'With wsPrepUnife.Cells(p, c3)
                    '.NumberFormat = "@"
                    '.Value = "'" & wsListe.Cells(l3, c3).Text
                'End With
            'Next c3
            'wsPrepUnife.Cells(p, 15) = "No Info"
        'End If
    'Next l3
    '****************************************'
    Application.DisplayAlerts = False
    '****************************************'
    ' Ensure the destination columns are formatted as Text so values like 2426-08-6 are not
    ' auto-converted to dates when using TextToColumns.
    wsPrepUnife.Columns("O:S").NumberFormat = "@"
    wsPrepUnife.Columns("O:O").TextToColumns Destination:=wsPrepUnife.Range("O1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :=":", FieldInfo:=Array(Array(1, 2), Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2)), TrailingMinusNumbers:=True
    '****************************************'
    Application.DisplayAlerts = True
    '****************************************'
    wsMaj.Cells(3, "E") = "Exécution Terminé"
    wsMaj.Cells(3, "E").Interior.ColorIndex = 9
    wsMaj.Cells(3, "G") = Date & " à " & Time
    Heure_fin1 = Time
    Durée1 = Heure_fin1 - Heure_debut1
    wsMaj.Cells(3, "H") = Durée1
    '****************************************'
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    '****************************************'
    ActiveWorkbook.Save
    '****************************************'
    End Sub





















