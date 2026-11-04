Attribute VB_Name = "Module1"

Public Choix_CBO As String
Public Fichier_trame As String

Sub New_COMP()
'
'Macro enregistrer le 07/09/2016 par Mr Yassine HASSEN
'
Dim a, l, i, j, k, m, n, f, compteur, NbCom, Nbpar, Nbpck, existe, Nbpak, Nbsum As Integer
Dim Vmin, Vmax, Vma, Vmi As Variant
Dim File, Message, nom, Val1, Val2, Val3, Val4, Val5, Val6, comp, Chemin, Chemin1, NumDEQ, test  As String

l = Range("A65536").End(xlUp).Row
'Nom du fichier extract
nom = ActiveWorkbook.Name

'Ouverture fichier trame
  Chemin = "E:\Validation des équivalents\Maac\trame1.xls"
  Chemin1 = "E:\Validation des équivalents\Maac\comparatif\type.xls"
  
  Workbooks.Open Chemin, 0, ReadOnly:=False
  Fichier_trame = ActiveWorkbook.Name
  
'Parcourir les lignes de l'Extract
For i = 2 To l
Workbooks(nom).Activate
'Forcer la valeur du ComboBox à "" , à moins que l'utilisateur choisisse une valeur
Choix_CBO = ""
UserForm1.CommandButton2.Visible = False
'Tester si le fichier Comp existe
comp = Workbooks(nom).Path & "\" & Workbooks(nom).Sheets(1).Cells(i, 1).Value & "\COMP.xlsx"
existe = 0
    If ExistenceFichier(comp) Then

    Workbooks.Open comp, 0, ReadOnly:=False
    If Workbooks("COMP.xlsx").Sheets(1).Cells(1, 4).Value = Workbooks(nom).Sheets(1).Cells(i, 1).Value Then
    existe = 2
    Workbooks(nom).Sheets(1).Cells(i, 2).Interior.Color = RGB(40, 200, 50)
    UserForm1.TextBox2 = "Le fichier Comparatif est déja traité !!"
    Else
    If Workbooks("COMP.xlsx").Sheets(1).Cells(1, 3).Value = Workbooks(nom).Sheets(1).Cells(i, 1).Value Then
    existe = 3
    Workbooks(nom).Sheets(1).Cells(i, 2).Interior.Color = RGB(40, 200, 50)
    UserForm1.TextBox2 = "Le fichier Comparatif est déja traité (Version ancienne)!!"
    Else
    existe = 1
    UserForm1.TextBox2 = "Le fichier Comparatif est téléchargé !"
    UserForm1.CommandButton2.Visible = True
    End If
    End If
    Else
    
    UserForm1.TextBox2 = "Le fichier Comparatif n'est pas téléchargé !"
    End If

'Renseigner le Userform
UserForm1.TextBox1 = Workbooks(nom).Sheets(1).Cells(i, 1).Value
UserForm1.TextBox3 = Workbooks(nom).Sheets(1).Cells(i, 7).Value
UserForm1.TextBox4 = Workbooks(nom).Sheets(1).Cells(i, 8).Value
UserForm1.Show


If Choix_CBO = "" Then
    Workbooks(nom).Sheets(1).Cells(i, 2).Interior.Color = RGB(219, 20, 20)
Else
If Choix_CBO = "SILICON" Then
Call SILICON

Else
   
   If existe = 1 Then
        
        Workbooks(nom).Sheets(1).Cells(i, 2).Interior.Color = RGB(0, 190, 250)
        Workbooks(Fichier_trame).Sheets(Choix_CBO).Copy Before:=Workbooks("COMP.xlsx").Sheets(1)
         
    'Num DEQ
        Workbooks("COMP.xlsx").Sheets(1).Name = "Comparatif"
        NumDEQ = Workbooks(nom).Sheets(1).Cells(i, 1).Value
        Workbooks("COMP.xlsx").Sheets(1).Cells(1, 4).Value = NumDEQ
       
    
    'Import des données de l'anglet Parametric/Technical
    
For a = 1 To Worksheets.Count

If Worksheets(a).Name = "Technical" Then

        Nbpar = Workbooks("COMP.xlsx").Sheets("Technical").Columns(1).Cells.Find("*", , , , , xlPrevious).Row
        NbCom = Workbooks("COMP.xlsx").Sheets(1).Columns(5).Cells.Find("*", , , , , xlPrevious).Row
        l = 1
        For k = 7 To Nbpar
        test = 0
            For j = 18 To NbCom
                If Workbooks("COMP.xlsx").Sheets("Technical").Cells(k, 1).Value = Workbooks("COMP.xlsx").Sheets(1).Cells(j, 5).Value And j <= NbCom Then
                Val1 = Workbooks("COMP.xlsx").Sheets("Technical").Cells(k, 2).Value
                Val2 = Workbooks("COMP.xlsx").Sheets("Technical").Cells(k, 3).Value
                Workbooks("COMP.xlsx").Sheets(1).Cells(j, 3) = Val1
            Workbooks("COMP.xlsx").Sheets(1).Cells(j, 4) = Val2
        test = 1
                End If
        
                If j = NbCom And test = 0 Then
                Workbooks("COMP.xlsx").Sheets(1).Range("B" & NbCom + l & ":" & "D" & l + NbCom) = Workbooks("COMP.xlsx").Sheets("Technical").Range("A" & k & ":" & "C" & k).Value
                l = l + 1
                End If
        
            Next j
        Next k
End If

  
    'Import des données de l'anglet "summary" P/N + Mfr + Description / RoHS
    
If Worksheets(a).Name = "Overview" Then

     Nbsum = Workbooks("COMP.xlsx").Sheets("Overview").Columns(1).Cells.Find("*", , , , , xlPrevious).Row
        NbCom = Workbooks("COMP.xlsx").Sheets(1).Columns(5).Cells.Find("*", , , , , xlPrevious).Row
       
           For k = 7 To Nbsum
        
              For j = 7 To NbCom
                If Workbooks("COMP.xlsx").Sheets("Overview").Cells(k, 1).Value = Workbooks("COMP.xlsx").Sheets(1).Cells(j, 5).Value And j <= NbCom Then
                Val3 = Workbooks("COMP.xlsx").Sheets("Overview").Cells(k, 2).Value
                Val4 = Workbooks("COMP.xlsx").Sheets("Overview").Cells(k, 3).Value
                Workbooks("COMP.xlsx").Sheets(1).Cells(j, 3) = Val3
            Workbooks("COMP.xlsx").Sheets(1).Cells(j, 4) = Val4
        
                End If
        
        
              Next j
           Next k
 End If
 
    'Rempliçage des champs Package/Compliance
    
If Worksheets(a).Name = "Compliance" Then

        Nbpak = Workbooks("COMP.xlsx").Sheets("Compliance").Columns(1).Cells.Find("*", , , , , xlPrevious).Row
        NbCom = Workbooks("COMP.xlsx").Sheets(1).Columns(5).Cells.Find("*", , , , , xlPrevious).Row
    
         For k = 7 To Nbpak
        
             For j = 7 To NbCom
                If Workbooks("COMP.xlsx").Sheets("Compliance").Cells(k, 1).Value = Workbooks("COMP.xlsx").Sheets(1).Cells(j, 5).Value And j <= NbCom Then
                Val5 = Workbooks("COMP.xlsx").Sheets("Compliance").Cells(k, 2).Value
                Val6 = Workbooks("COMP.xlsx").Sheets("Compliance").Cells(k, 3).Value
                Workbooks("COMP.xlsx").Sheets(1).Cells(j, 3) = Val5
             Workbooks("COMP.xlsx").Sheets(1).Cells(j, 4) = Val6
         
                End If
        
        
             Next j
         Next k
End If

Next a


 'Supprimer les feuilles supplémentaire
    
         For m = Workbooks("COMP.xlsx").Sheets.Count To 2 Step -1

                Application.DisplayAlerts = False
                Workbooks("COMP.xlsx").Sheets(m).Delete
                Application.DisplayAlerts = True
    
         Next m

 ' Mise en forme couleur pour différence


    For f = 18 To Workbooks("COMP.xlsx").Sheets(1).Columns(5).Cells.Find("*", , , , , xlPrevious).Row
                 Range("D" & f).Select
                    Selection.FormatConditions.Delete
                    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, _
                 Formula1:="=C" & f
                    Selection.FormatConditions(1).Interior.ColorIndex = 6
    
                 Range("C" & f).Select
                    Selection.FormatConditions.Delete
                    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, _
                 Formula1:="=D" & f
                 Selection.FormatConditions(1).Interior.ColorIndex = 6
    Next f
     
   '------------------------------------------------------------------------------------------
    'Cas ou le comparatif n'est pas télécharger depuis silicon
   '------------------------------------------------------------------------------------------
    
    
    Else
    If existe = 0 Then

     Workbooks(nom).Sheets(1).Cells(i, 2).Interior.Color = RGB(220, 90, 20)
     Workbooks.Open Chemin1, 0, ReadOnly:=False
     ActiveWorkbook.SaveAs FileName:=comp
     
     ' Déplacer le comparatif dans la première anglet

        Workbooks("COMP.xlsx").Sheets(Choix_CBO).Copy Before:=Workbooks("COMP.xls").Sheets(1)
        Workbooks("COMP.xlsx").Sheets(1).Name = "Comparatif"
        
     ' Rempliçage le champs numéro_DEQ
     
        NumDEQ = Workbooks(nom).Sheets(1).Cells(i, 1).Value
        Workbooks("COMP.xlsx").Sheets(1).Cells(1, 4).Value = NumDEQ
        
     ' Supprimer les feuilles supplémentaire
     
For m = Workbooks("COMP.xlsx").Sheets.Count To 2 Step -1
'If
Application.DisplayAlerts = False
Workbooks("COMP.xls").Sheets(m).Delete
Application.DisplayAlerts = True

Next m
' Mise en forme couleur pour différence


    For f = 18 To Workbooks("COMP.xls").Sheets(1).Columns(5).Cells.Find("*", , , , , xlPrevious).Row
                 Range("D" & f).Select
                    Selection.FormatConditions.Delete
                    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, _
                 Formula1:="=C" & f
                    Selection.FormatConditions(1).Interior.ColorIndex = 6
    
                 Range("C" & f).Select
                    Selection.FormatConditions.Delete
                    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, _
                 Formula1:="=D" & f
                 Selection.FormatConditions(1).Interior.ColorIndex = 6
    Next f

     
    ' Workbooks("trame1.xls").Close

    End If
    End If
End If
End If
If ExistenceFichier(comp) Then
'If IsFileOpen(comp) Then
Workbooks("COMP.xlsx").Save
Workbooks("COMP.xlsx").Close
'End If
End If


Next i
 Workbooks(Fichier_trame).Close



End Sub



Function ExistenceFichier(ByVal sFichier As String) As Boolean
  ExistenceFichier = Dir(sFichier) <> ""
End Function



Function IsFileOpen(FileName As String)
    Dim filenum As Integer, errnum As Integer

    On Error Resume Next   ' Turn error checking off.
    filenum = FreeFile()   ' Get a free file number.
    ' Attempt to open the file and lock it.
    Open FileName For Input Lock Read As #filenum
    Close filenum          ' Close the file.
    errnum = err           ' Save the error number that occurred.
    On Error GoTo 0        ' Turn error checking back on.

    ' Check to see which error occurred.
    Select Case errnum

        ' No error occurred.
        ' File is NOT already open by another user.
        Case 0
         IsFileOpen = False

        ' Error number for "Permission Denied."
        ' File is already opened by another user.
        Case 70
            IsFileOpen = True

        ' Another error occurred.
        Case Else
            Error errnum
    End Select

End Function
