Attribute VB_Name = "Module6"
Sub Verif()

Dim err, lin As Integer
Dim RapEQUIV, RapPDN As String
Dim comp, INIT, EQUIV, PDN As String

Dim taille As Double
Dim Rep, adr As String

RapEQUIV = "Les DEQs suivants ne sont pas conformes" & vbCr & vbCr
RapPDN = "Les DEQs suivants ont un PDN Manquant" & vbCr & vbCr
For lin = 2 To ActiveWorkbook.Sheets(1).Columns(1).Cells.Find("*", , , , , xlPrevious).Row

EQUIV = ActiveWorkbook.Path & "\" & Cells(lin, 1).Value & "\EQUIV.pdf"
INIT = ActiveWorkbook.Path & "\" & Cells(lin, 1).Value & "\INIT.pdf"
PDN = ActiveWorkbook.Path & "\" & Cells(lin, 1).Value & "\PDN.pdf"
comp = ActiveWorkbook.Path & "\" & Cells(lin, 1).Value & "\COMP.pdf"

'Tester si il y'a un fichier EQUIV alors que pas d'equivalent proposé
    If ActiveWorkbook.Sheets(1).Cells(lin, 10) = "SANS" Then
        
        If ExistenceFichier(EQUIV) Then
        RapEQUIV = RapEQUIV & ActiveWorkbook.Path & "\" & Cells(lin, 1) & " ==> Datasheet EQUIV présent à enlever" & vbCr & vbCr
        Else
            
            If ActiveWorkbook.Sheets(1).Cells(lin, 2) <> "vr" Then
            RapEQUIV = RapEQUIV & ActiveWorkbook.Path & "\" & Cells(lin, 1) & " ==> Sanction à mettre en vr! " & vbCr & vbCr
            End If
        
        End If
    
'Tester s'il y'a un Datasheet ou COMP.pdf manquant(s)
    Else
        If ActiveWorkbook.Sheets(1).Cells(lin, 2) = "vr" And ActiveWorkbook.Sheets(1).Cells(lin, 25) = "vr" Then
        RapEQUIV = RapEQUIV & ActiveWorkbook.Path & "\" & Cells(lin, 1) & " ==> Il faut ajouter un commentaire " & vbCr & vbCr
        End If
        
        If Not ExistenceFichier(INIT) Or Not ExistenceFichier(EQUIV) Or Not ExistenceFichier(comp) Then
        RapEQUIV = RapEQUIV & ActiveWorkbook.Path & "\" & Cells(lin, 1) & " ==> Il y a un fichier manquant" & vbCr & vbCr
        End If
        
    End If
'Tester si il y'a une date LBO ou un statut jaune sans avoir un document PDN
    If (ActiveWorkbook.Sheets(1).Cells(lin, 9) = "Jaune" Or ActiveWorkbook.Sheets(1).Cells(lin, 19).Value <> "") And Not ExistenceFichier(PDN) Then
    RapPDN = RapPDN & ActiveWorkbook.Path & "\" & Cells(lin, 1) & vbCr
    Cells(lin, 1).Interior.Color = RGB(0, 255, 0)
    End If


           
'Tester s'il y'a un mauvais format
    err = 0
    Rep = ActiveWorkbook.Path & "\" & Cells(lin, 1).Value & "\"
    taille = 0
    adr = Dir(Rep)
    
    While (adr <> "")
       taille = taille + FileLen(Rep & "\" & adr)
       adr = Dir
       If adr <> "COMP.xls" And adr <> "EQUIV.pdf" And adr <> "INIT.pdf" And adr <> "EQUIV.pdf" And adr <> "PDN.pdf" And adr <> "" Then
       err = 1
       End If
    Wend
    
'    If taille / 1024 > 5000 Then
'       ActiveWorkbook.Sheets(1).Cells(lin, 1).Interior.Color = RGB(250, 0, 50)
'    End If
    If err = 1 Then
        RapEQUIV = RapEQUIV & ActiveWorkbook.Path & "\" & Cells(lin, 1) & " ==> Il y a un format différent dans les fichiers" & vbCr & vbCr
    End If
    
           
Next lin

MsgBox (RapPDN)
MsgBox (RapEQUIV)

End Sub

Public Function ExisteFichier(S As String) As Boolean
Dim tatiak As Object
    Set tatiak = CreateObject("Scripting.FileSystemObject")
    ExisteFichier = tatiak.FileExists(S)
End Function

Function ExistenceFichier(ByVal sFichier As String) As Boolean
  ExistenceFichier = Dir(sFichier) <> ""
End Function

