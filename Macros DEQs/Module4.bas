Attribute VB_Name = "Module4"
Sub New_taille()
'
' taille Macro
'
Dim lin As Integer
Dim taille As Double
Dim Rep, adr As String
'Prend chaque cellulle ou il y a la référence du DEQ
For lin = 2 To Range("A65536").End(xlUp).Row
'Les repertoire DEQ devront etre dans le meme path que le fichier excel

    Rep = ActiveWorkbook.Path & "\" & Cells(lin, 1).Value
    taille = 0

    adr = Dir(Rep & "\")
 
    While (adr <> "")
       taille = taille + FileLen(Rep & "\" & adr)
       adr = Dir
    Wend
' MsgBox (taille / 1024)
    If (taille / 1024) > 5000 Then
    MsgBox ("Attention ! " & Cells(lin, 1) & " a une taille de : " & Int(taille / 1024) & " KOctets !!")
    Cells(lin, 1).Interior.Color = RGB(250, 0, 50)
    End If

Next lin
MsgBox ("vérifié")
'
  Range("F24").Select
  'MsgBox ("terminé")
End Sub
