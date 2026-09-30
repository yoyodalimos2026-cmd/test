Attribute VB_Name = "Module23"
Sub AjouterTempsTraitement()
  Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Const colType As Integer = 17 ' Colonne Q
    
    ' 1. Définition de l'onglet par son index (le 1er onglet à gauche)
    Set ws = ActiveWorkbook.Sheets(1)
    
    ' 2. Insérer la nouvelle colonne après la colonne 17 (donc en colonne 18 / R)
    ws.Columns(colType + 1).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ws.Cells(1, colType + 1).Value = "Temps estimé de traitement (h)"
    
    ' 3. Trouver la dernière ligne remplie dans la colonne 17
    lastRow = ws.Cells(ws.Rows.Count, colType).End(xlUp).Row
    
    ' 4. Remplissage des données avec une boucle
    For i = 2 To lastRow
        ' On utilise Trim() pour éviter les erreurs si un espace traîne dans la cellule
        Select Case Trim(ws.Cells(i, colType).Value)
            Case "Actif":               ws.Cells(i, colType + 1).Value = 1.5
            Case "Passif":              ws.Cells(i, colType + 1).Value = 1.5
            Case "Cots":                ws.Cells(i, colType + 1).Value = 2.5
            Case "Mécanique Standard":  ws.Cells(i, colType + 1).Value = 1.5
            Case "Particulier":         ws.Cells(i, colType + 1).Value = 2
        End Select
    Next i
    
    ' Mise en forme rapide
    ws.Columns(colType + 1).AutoFit
    
    MsgBox "Colonne 18 mise à jour sur la feuille 1 !", vbInformation
End Sub
