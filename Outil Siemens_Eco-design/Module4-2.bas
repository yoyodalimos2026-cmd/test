Attribute VB_Name = "Module4"
Sub Mat_Mix_CAT()
    '****************************************'
    Dim wsBord As Worksheet
    Dim wsSubstance As Worksheet
    Dim wsMaj As Worksheet
    Dim Nb_Ligne_BORD As Long
    Dim Nb_Ligne_Substance As Long
    Dim e As Long
    Dim ct As Long
    Dim concatValue As String
    Dim Mat As String
    Dim Mix As String
    Dim excludedValues As Variant
    Dim Matitems As Variant
    Dim Mixitems As Variant
    '****************************************'
    Set wsBord = ActiveWorkbook.Sheets("Donnée_BORD")
    Set wsSubstance = ActiveWorkbook.Sheets("Substance Declaration")
    Set wsMaj = ActiveWorkbook.Sheets("MAJ")
    '****************************************'
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    '****************************************'
    Nb_Ligne_BORD = wsBord.UsedRange.Rows.Count
    Nb_Ligne_Substance = wsSubstance.UsedRange.Rows.Count
    excludedValues = Array("", "-")
    '****************************************'
    With wsMaj
        .Cells(5, "D").Value = Nb_Ligne_Prepa_Unife
        .Cells(5, "F").Value = Now
        Heure_debut3 = Time
        .Cells(5, "E").Value = "Exécution en cours"
        .Cells(5, "G").Value = ""
        .Cells(5, "E").Interior.ColorIndex = 4
    End With
    '****************************************'
    For ct = 15 To Nb_Ligne_Substance
        With wsSubstance
            If Not IsEmpty(.Cells(ct, "B").Value) Then
                concatValue = .Cells(ct, "L").Value & ";" & .Cells(ct, "M").Value
                '****************************************'
                Set foundCell = wsBord.Range("I1:I" & Nb_Ligne_BORD).Find(What:=concatValue, LookIn:=xlValues, LookAt:=xlWhole)
                '****************************************'
                If Not foundCell Is Nothing Then
                    Mat = wsBord.Cells(foundCell.Row, "T").Value
                    Mix = wsBord.Cells(foundCell.Row, "U").Value
                    '****************************************'
                    If IsError(Application.Match(Mat, excludedValues, 0)) Then
                        Matitems = Split(Mat & "//", "//")
                    Else
                        Matitems = Array()
                    End If
                    
                    If IsError(Application.Match(Mix, excludedValues, 0)) Then
                        Mixitems = Split(Mix & "//", "//")
                    Else
                        Mixitems = Array()
                    End If
                    '****************************************'
                    e = ct + 1
                    '****************************************'
                    Do While e <= Nb_Ligne_Substance And IsEmpty(.Cells(e, "B").Value)
                        If StrComp(.Cells(e, "Y").Value, "SVHC", vbTextCompare) = 0 Then
                            cas = .Cells(e, "U").Value
                            '****************************************'
                            If IsArrayNotEmpty(Matitems) Then
                                For i = LBound(Matitems) To UBound(Matitems) - 1
                                    item = Trim(Matitems(i))
                                    subItems = Split(item, ":")
                                    If StrComp(Trim(subItems(0)), cas, vbTextCompare) = 0 Then
                                        .Cells(e, "R") = Trim(subItems(1))
                                    End If
                                Next i
                            End If
                            '****************************************'
                            If IsArrayNotEmpty(Mixitems) Then
                                For j = LBound(Mixitems) To UBound(Mixitems) - 1
                                    item1 = Trim(Mixitems(j))
                                    subItems1 = Split(item1, ":")
                                    If StrComp(Trim(subItems1(0)), cas, vbTextCompare) = 0 Then
                                        .Cells(e, "S") = Trim(subItems1(1))
                                    End If
                                Next j
                            End If
                            
                        End If
                            '****************************************'
                        e = e + 1
                    Loop
                    '****************************************'
                    ct = e - 1
                End If
            End If
        End With
    '****************************************'
    Next ct
    '****************************************'
    
    p = wsSubstance.Cells(wsSubstance.Rows.Count, "Y").End(xlUp).Row
        
    For l3 = 2 To Nb_Ligne_BORD
        If wsBord.Cells(l3, "N").Value = "Unknown" Then
            p = p + 1
            wsSubstance.Cells(p, "B").Value = wsBord.Cells(l3, "G").Value
            wsSubstance.Cells(p, "L").Value = wsBord.Cells(l3, "C").Value
            wsSubstance.Cells(p, "M").Value = wsBord.Cells(l3, "M").Value
            wsSubstance.Cells(p, "U") = "No Info"
        End If
    Next l3
    '****************************************'
    With wsMaj
        .Cells(5, "E").Value = "Exécution Terminé"
        .Cells(5, "E").Interior.ColorIndex = 9
        .Cells(5, "G").Value = Now
        Heure_fin3 = Time
        Durée3 = Heure_fin3 - Heure_debut3
        .Cells(5, "H").Value = Durée3
   End With
    '****************************************'
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    ActiveWorkbook.Save
End Sub
Function IsArrayNotEmpty(arr As Variant) As Boolean
    On Error Resume Next
    IsArrayNotEmpty = (Not IsEmpty(arr)) And (UBound(arr) >= LBound(arr))
    On Error GoTo 0
End Function


