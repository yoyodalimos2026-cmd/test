Attribute VB_Name = "Module17"
Sub New_Verif_comment()
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


    
    '***********************************************************************
    'Tester si la proposition faite par le client/Fabricant
    '***********************************************************************
    
    If ActiveWorkbook.Sheets(1).Cells(lin, 15).Value Like "*Recommandation*" _
        Or ActiveWorkbook.Sheets(1).Cells(lin, 15).Value Like "*Proposition*" _
        Or ActiveWorkbook.Sheets(1).Cells(lin, 15).Value Like "*Choix*" _
        Or ActiveWorkbook.Sheets(1).Cells(lin, 15).Value Like "*Demande*" _
        Or ActiveWorkbook.Sheets(1).Cells(lin, 15).Value Like "*Suggestion*" _
        Or ActiveWorkbook.Sheets(1).Cells(lin, 15).Value Like "*proposé" _
        Or ActiveWorkbook.Sheets(1).Cells(lin, 15).Value Like "*recommandation*" _
        Or ActiveWorkbook.Sheets(1).Cells(lin, 15).Value Like "*proposition*" _
        Or ActiveWorkbook.Sheets(1).Cells(lin, 15).Value Like "*demande*" _
        Or ActiveWorkbook.Sheets(1).Cells(lin, 15).Value Like "*choix*" _
        Or ActiveWorkbook.Sheets(1).Cells(lin, 15).Value Like "*par le*" _
        Or ActiveWorkbook.Sheets(1).Cells(lin, 15).Value Like "*selon*" _
        Or ActiveWorkbook.Sheets(1).Cells(lin, 15).Value Like "*proposer*" _
        Or ActiveWorkbook.Sheets(1).Cells(lin, 15).Value Like "*suggestion*" Then
        
         
                If ActiveWorkbook.Sheets(1).Cells(lin, 15).Value Like "*client*" _
                Or ActiveWorkbook.Sheets(1).Cells(lin, 15).Value Like "*Client*" Then
        
                        If ActiveWorkbook.Sheets(1).Cells(lin, 25) <> "" Then
            
                            ActiveWorkbook.Sheets(1).Cells(lin, 25).Value = ActiveWorkbook.Sheets(1).Cells(lin, 25).Value & " / Equivalence faite selon la recommandation client"
                        Else
                              ActiveWorkbook.Sheets(1).Cells(lin, 25).Value = "Equivalence faite selon la recommandation client"
                        
                        End If
                        '**********************************************
                          Cells(lin, 25).Interior.Color = RGB(255, 0, 255)   ' Rose
                         '**********************************************
               End If
        

        
                If ActiveWorkbook.Sheets(1).Cells(lin, 15).Value Like "*fabricant*" _
                Or ActiveWorkbook.Sheets(1).Cells(lin, 15).Value Like "*Fabricant*" Then
        
                         If ActiveWorkbook.Sheets(1).Cells(lin, 25) <> "" Then
            
                                ActiveWorkbook.Sheets(1).Cells(lin, 25).Value = ActiveWorkbook.Sheets(1).Cells(lin, 25).Value & " / Equivalence faite selon la recommandation fabricant"
                         Else
                                ActiveWorkbook.Sheets(1).Cells(lin, 25).Value = "Equivalence faite selon la recommandation fabricant"
                        
                         End If
                         
                        '**********************************************
                          Cells(lin, 25).Interior.Color = RGB(128, 0, 255)   ' Violet
                        '**********************************************
               End If
    
    End If
    
    
    
    '***********************************************************************
    '***********************************************************************
           
Next lin

End Sub

Public Function ExisteFichier(S As String) As Boolean
Dim tatiak As Object
    Set tatiak = CreateObject("Scripting.FileSystemObject")
    ExisteFichier = tatiak.FileExists(S)
End Function

Function ExistenceFichier(ByVal sFichier As String) As Boolean
  ExistenceFichier = Dir(sFichier) <> ""
End Function

