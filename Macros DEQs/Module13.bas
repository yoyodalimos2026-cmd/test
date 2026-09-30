Attribute VB_Name = "Module13"
Sub New_Verif()
'
'
'
Dim err, lin As Integer
Dim RapEQUIV, RapPDN As String
Dim comp, INIT, EQUIV, PDN As String
Dim Date_LBO As Date

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

    If ActiveWorkbook.Sheets(1).Cells(lin, 11) = "SANS" Then
        
        If ExistenceFichier(EQUIV) Then
        RapEQUIV = RapEQUIV & ActiveWorkbook.Path & "\" & Cells(lin, 1) & " ==> Datasheet EQUIV présent à enlever" & vbCr & vbCr
                          '**********************************************
                          Cells(lin, 12).Interior.Color = RGB(255, 0, 0)   ' Rouge
                          '**********************************************
    
        End If
        
'Tester si le DEQ sans équivalent et la sanction Vs
          
        If ActiveWorkbook.Sheets(1).Cells(lin, 2) <> "vr" Then
            RapEQUIV = RapEQUIV & ActiveWorkbook.Path & "\" & Cells(lin, 1) & " ==> Sanction à mettre en vr! " & vbCr & vbCr
                          '***********************************************
                          Cells(lin, 11).Interior.Color = RGB(0, 255, 0)  'Vert
                          '***********************************************
         End If
        
    
    
'Tester s'il y'a un Datasheet ou COMP.pdf manquant(s)

    Else
        If ActiveWorkbook.Sheets(1).Cells(lin, 2) = "vr" And ActiveWorkbook.Sheets(1).Cells(lin, 27) = "" Then
        RapEQUIV = RapEQUIV & ActiveWorkbook.Path & "\" & Cells(lin, 1) & " ==> Il faut ajouter un commentaire " & vbCr & vbCr
                         '**********************************************
                          Cells(lin, 4).Interior.Color = RGB(0, 0, 255) ' Bleu
                         '**********************************************
                          
        End If
        
        If Not ExistenceFichier(INIT) Or Not ExistenceFichier(EQUIV) Or Not ExistenceFichier(comp) Then
        RapEQUIV = RapEQUIV & ActiveWorkbook.Path & "\" & Cells(lin, 1) & " ==> Il y a un fichier manquant" & vbCr & vbCr
        End If
        
    End If
    
  
'Tester si le  document PDN N'existe Pas

    If Not ExistenceFichier(PDN) Then
             RapPDN = RapPDN & ActiveWorkbook.Path & "\" & Cells(lin, 1) & vbCr
                         '*********************************************
                         Cells(lin, 1).Interior.Color = RGB(96, 0, 255)  ' Indigo
                         '*********************************************
    End If
  
'Tester si il y a un statut Jaune sans date LBO ou avec LBO erronée < date aujourd'hui


         If (ActiveWorkbook.Sheets(1).Cells(lin, 10) = "Jaune" And ActiveWorkbook.Sheets(1).Cells(lin, 21) = "") Then
                         '**********************************************
                         Cells(lin, 10).Interior.Color = RGB(0, 255, 255)  ' Cyan
                         '**********************************************
         End If
    
         If (ActiveWorkbook.Sheets(1).Cells(lin, 10) = "Jaune" And ActiveWorkbook.Sheets(1).Cells(lin, 21).Value < Now And ActiveWorkbook.Sheets(1).Cells(lin, 21).Value <> "") Then
                         '**********************************************
                         Cells(lin, 10).Interior.Color = RGB(0, 255, 255)  ' Cyan
                         '**********************************************
         End If
    
         If (ActiveWorkbook.Sheets(1).Cells(lin, 10) <> "Jaune" And ActiveWorkbook.Sheets(1).Cells(lin, 21).Value > Now And ActiveWorkbook.Sheets(1).Cells(lin, 21).Value <> "") Then
                         '**********************************************
                         Cells(lin, 10).Interior.Color = RGB(0, 255, 255)  ' Cyan
                         '**********************************************
         End If
         
         
      'Tester si le statut de l'équivalent Orange et le valideur oublie de mettre un comentaire


         If ActiveWorkbook.Sheets(1).Cells(lin, 14) = "Orange" Then
         
                If ActiveWorkbook.Sheets(1).Cells(lin, 27) <> "" Then
            
                       ActiveWorkbook.Sheets(1).Cells(lin, 27).Value = ActiveWorkbook.Sheets(1).Cells(lin, 27).Value & " / Equivalent Orange"
                Else
                        ActiveWorkbook.Sheets(1).Cells(lin, 27).Value = "Equivalent Orange"
                        
                End If
                        '**********************************************
                          Cells(lin, 27).Interior.Color = RGB(255, 96, 0)   ' Orange
                        '**********************************************
         End If
         
      'Tester si le statut de l'équivalent Rouge et le valideur oublie de mettre un comentaire

         If ActiveWorkbook.Sheets(1).Cells(lin, 14) = "Rouge" Then
         
                If ActiveWorkbook.Sheets(1).Cells(lin, 27) <> "" Then
            
                       ActiveWorkbook.Sheets(1).Cells(lin, 27).Value = ActiveWorkbook.Sheets(1).Cells(lin, 27).Value & " / Equivalent Rouge"
                Else
                        ActiveWorkbook.Sheets(1).Cells(lin, 27).Value = "Equivalent Rouge"
                        
                End If
                        '**********************************************
                          Cells(lin, 27).Interior.Color = RGB(255, 0, 0)   ' Rouge
                        '**********************************************
         End If
    
    '***********************************************************************
    'Tester si la proposition faite par le client/Fabricant
    '***********************************************************************
    
    If ActiveWorkbook.Sheets(1).Cells(lin, 17).Value Like "*Recommandation*" _
        Or ActiveWorkbook.Sheets(1).Cells(lin, 17).Value Like "*Proposition*" _
        Or ActiveWorkbook.Sheets(1).Cells(lin, 17).Value Like "*Choix*" _
        Or ActiveWorkbook.Sheets(1).Cells(lin, 17).Value Like "*Demande*" _
        Or ActiveWorkbook.Sheets(1).Cells(lin, 17).Value Like "*Suggestion*" _
        Or ActiveWorkbook.Sheets(1).Cells(lin, 17).Value Like "*proposé" _
        Or ActiveWorkbook.Sheets(1).Cells(lin, 17).Value Like "*recommandation*" _
        Or ActiveWorkbook.Sheets(1).Cells(lin, 17).Value Like "*proposition*" _
        Or ActiveWorkbook.Sheets(1).Cells(lin, 17).Value Like "*demande*" _
        Or ActiveWorkbook.Sheets(1).Cells(lin, 17).Value Like "*choix*" _
        Or ActiveWorkbook.Sheets(1).Cells(lin, 17).Value Like "*par le*" _
        Or ActiveWorkbook.Sheets(1).Cells(lin, 17).Value Like "*selon*" _
        Or ActiveWorkbook.Sheets(1).Cells(lin, 17).Value Like "*proposer*" _
        Or ActiveWorkbook.Sheets(1).Cells(lin, 17).Value Like "*suggestion*" Then
        
         
                If ActiveWorkbook.Sheets(1).Cells(lin, 17).Value Like "*client*" _
                Or ActiveWorkbook.Sheets(1).Cells(lin, 17).Value Like "*Client*" Then
        
                        If ActiveWorkbook.Sheets(1).Cells(lin, 27) <> "" Then
            
                            ActiveWorkbook.Sheets(1).Cells(lin, 27).Value = ActiveWorkbook.Sheets(1).Cells(lin, 27).Value & " / Equivalence faite selon la recommandation client"
                        Else
                              ActiveWorkbook.Sheets(1).Cells(lin, 27).Value = "Equivalence faite selon la recommandation client"
                        
                        End If
                        '**********************************************
                          Cells(lin, 27).Interior.Color = RGB(255, 0, 255)   ' Rose
                         '**********************************************
               End If
        

        
                If ActiveWorkbook.Sheets(1).Cells(lin, 17).Value Like "*fabricant*" _
                Or ActiveWorkbook.Sheets(1).Cells(lin, 17).Value Like "*Fabricant*" Then
        
                         If ActiveWorkbook.Sheets(1).Cells(lin, 27) <> "" Then
            
                                ActiveWorkbook.Sheets(1).Cells(lin, 27).Value = ActiveWorkbook.Sheets(1).Cells(lin, 27).Value & " / Equivalence faite selon la recommandation fabricant"
                         Else
                                ActiveWorkbook.Sheets(1).Cells(lin, 27).Value = "Equivalence faite selon la recommandation fabricant"
                        
                         End If
                         
                        '**********************************************
                          Cells(lin, 27).Interior.Color = RGB(128, 0, 255)   ' Violet
                        '**********************************************
               End If
    
    End If
    
    
    
    '***********************************************************************
    '***********************************************************************
           
'Tester s'il y'a un mauvais format

    err = 0
    Rep = ActiveWorkbook.Path & "\" & Cells(lin, 1).Value & "\"
    taille = 0
    adr = Dir(Rep)
    
    While (adr <> "")
       taille = taille + FileLen(Rep & "\" & adr)
       adr = Dir
       If adr <> "COMP.xlsx" And adr <> "COMP.pdf" And adr <> "INIT.pdf" And adr <> "EQUIV.pdf" And adr <> "PDN.pdf" And adr <> "" Then
       err = 1
       End If
    Wend
    

         If err = 1 Then
               RapEQUIV = RapEQUIV & ActiveWorkbook.Path & "\" & Cells(lin, 1) & " ==> Il y a un format différent dans les fichiers" & vbCr & vbCr
                           '*********************************************
                           Cells(lin, 2).Interior.Color = RGB(224, 0, 192) ' Rose
                           '*********************************************
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
