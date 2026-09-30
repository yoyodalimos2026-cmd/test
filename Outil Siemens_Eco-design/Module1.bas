Attribute VB_Name = "Module1"
Sub Preparation()
    Dim wsSource As Worksheet, wsDest As Worksheet, wsMaj As Worksheet
    Dim Nb_Ligne_BORD As Long
    Dim excludedValuesDict As Object
    Dim validListDict As Object
    Dim sourceData As Variant, destData() As Variant
    Dim i As Long, x As Long, y As Integer, j As Integer
    Dim ligne As Long, Block As Long
    Dim Heure_debut As Date, Heure_fin As Date, Durée As Double
    Dim cellVal As String
    Dim items As Variant, item As Variant, subItems As Variant
    Dim firstSubItem As String
    Dim validItems As String, invalidItems As String
    Dim key As Variant
    '****************************************'
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    '****************************************'
    Set wsSource = ActiveWorkbook.Sheets("Donnée_BORD")
    Set wsDest = ActiveWorkbook.Sheets("Preparation")
    Set wsMaj = ActiveWorkbook.Sheets("MAJ")
    '****************************************'
    Set excludedValuesDict = CreateObject("Scripting.Dictionary")
    excludedValuesDict.Add "Unknown", True
    excludedValuesDict.Add "NA", True
    excludedValuesDict.Add "non", True
    '****************************************'
    Nb_Ligne_BORD = wsSource.Cells(wsSource.Rows.Count, "R").End(xlUp).Row
    '****************************************'
    With wsMaj
        .Cells(2, "D") = Nb_Ligne_BORD
        .Cells(2, "F") = Date & " a " & Time
        Heure_debut = Time
        .Cells(2, "E") = "Exécution en cours"
        .Cells(2, "G") = ""
        .Cells(2, "E").Interior.ColorIndex = 4
    End With
    '****************************************'
    Dim r As Long, c As Long
    ReDim sourceData(1 To Nb_Ligne_BORD - 1, 1 To 18)
    For r = 2 To Nb_Ligne_BORD
        For c = 1 To 18
            sourceData(r - 1, c) = CStr(wsSource.Cells(r, c).Text)
        Next c
    Next r
    '****************************************'
    ReDim destData(1 To (Nb_Ligne_BORD * 3), 1 To 15)
    '****************************************'
    Set validListDict = CreateObject("Scripting.Dictionary")
    Dim lastValidRow As Long
    lastValidRow = wsSource.Cells(wsSource.Rows.Count, "S").End(xlUp).Row
    Dim validListRange As Variant
    validListRange = wsSource.Range(wsSource.Cells(2, "S"), wsSource.Cells(lastValidRow, "S")).Value
    '****************************************'
    For i = 1 To UBound(validListRange, 1)
        If Not IsEmpty(validListRange(i, 1)) Then
            validListDict(Trim(validListRange(i, 1))) = True
        End If
    Next i
    '****************************************'
    For i = 1 To UBound(sourceData, 1)
        
        ligne = (i - 1) * 3 + 1
        Block = ligne + 2
        y = 1
        
        For x = ligne To Block
        
            For j = 1 To 13
                destData(x, j) = CStr(sourceData(i, j))
            Next j
            
            If y = 1 Or y = 2 Then
                cellVal = sourceData(i, 17)
                If Len(cellVal) > 0 Then
                    If Not excludedValuesDict.Exists(cellVal) Then
                        destData(x, 15) = CStr(ExtractItems(CleanString(cellVal), y, validListDict))
                    End If
                End If
                y = y + 1
            Else
                cellVal = sourceData(i, 18)
                If Len(cellVal) > 0 Then
                    If Not excludedValuesDict.Exists(cellVal) Then
                        destData(x, 15) = CStr(CleanString(cellVal))
                    End If
                End If
            End If
        Next x
        
    Next i
    
    With wsDest.Range("A2").Resize(UBound(destData, 1), UBound(destData, 2))
        .NumberFormat = "@"
        .Value = destData
    End With

    For i = 1 To Nb_Ligne_BORD
        ligne = (i - 1) * 3 + 3
        With wsDest
            .Cells(ligne - 1, 14).Value = "SVHC"
            .Cells(ligne - 1, 14).Interior.ColorIndex = 10
            .Cells(ligne, 14).Value = "D(FA) not SVHC"
            .Cells(ligne, 14).Interior.ColorIndex = 14
            .Cells(ligne + 1, 14).Value = "D(FI)"
            .Cells(ligne + 1, 14).Interior.ColorIndex = 12
        End With
    Next i
    

    wsDest.Columns("O:O").TextToColumns Destination:=wsDest.Range("O1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar:="/", _
        FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1)), _
        TrailingMinusNumbers:=True
    

    With wsMaj
        .Cells(2, "E") = "Exécution Terminée"
        .Cells(2, "E").Interior.ColorIndex = 9
        .Cells(2, "G") = Date & " a " & Time
        Heure_fin = Time
        Durée = Heure_fin - Heure_debut
        .Cells(2, "H") = Durée
    End With
    
    ActiveWorkbook.Save
    
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Function CleanString(inputString As String) As String
    Dim result As String
    result = Replace(inputString, "entre", "")
    result = Replace(result, "et", ">")
    result = Replace(result, ">", ":>")
    result = Replace(result, "<", ":<")
    result = Replace(result, " ", "")
    result = Replace(result, Chr(13), "")
    result = Replace(result, Chr(10), "")
    result = Replace(result, "//", "/")
    result = Replace(result, ",", ".")
    CleanString = result
End Function

Function ExtractItems(cellVal As String, y As Integer, validListDict As Object) As String
    Dim items As Variant, item As Variant, subItems As Variant
    Dim validItems As String, invalidItems As String
    Dim firstSubItem As String
    Dim i As Long
    
    items = Split(cellVal, "/")
    validItems = ""
    invalidItems = ""
    
    For i = LBound(items) To UBound(items)
        item = Trim(items(i))
        subItems = Split(item, ":")
        If UBound(subItems) >= 0 Then
            firstSubItem = Trim(subItems(0))
            If validListDict.Exists(firstSubItem) Then
                validItems = validItems & item & "/"
            Else
                invalidItems = invalidItems & item & "/"
            End If
        End If
    Next i
    
    If Len(validItems) > 0 Then validItems = Left(validItems, Len(validItems) - 1)
    If Len(invalidItems) > 0 Then invalidItems = Left(invalidItems, Len(invalidItems) - 1)
    
    If y = 1 Then
        ExtractItems = validItems
    Else
        ExtractItems = invalidItems
    End If
End Function












