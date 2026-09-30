Attribute VB_Name = "Module3"
Sub MAJ_SOL_V3_27_10_23()


Windows(ThisWorkbook.Name).Activate
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(21, "G") = Date & " a " & Time
Heure_debut_SOL = Time
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(20, "G").Value = "Execution en cours"
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(20, "G").Interior.ColorIndex = 4
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(20, "E").Value = ""

Set fso = CreateObject("Scripting.FileSystemObject")
Current_Path = ThisWorkbook.Path
Nom_SOL = ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(20, "B").Value '<===========Extraire le nom du fichier a mettre a jour
SOL_File_path = Current_Path & "\" & Nom_SOL                           '<==========================='L'adresse du fichier SOL
SSDBL_file_path = Current_Path & "\" & ThisWorkbook.Name
Nom_fichier_SSdbl = ThisWorkbook.Name
'Debug.Print Nom_fichier_SSdbl
'=======================================Vérifier que le fichier existe sosu le dossier============================================
If fso.FileExists(SOL_File_path) = True Then

    ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(20, "E").Value = "Fichier Trouvé"
    ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(20, "E").Interior.ColorIndex = 4
    Set SOL_FILE = Workbooks.Open(SOL_File_path)
    Set ssdbl_FILE = Workbooks.Open(SSDBL_file_path)
    Windows(Nom_SOL).Activate
    Nb_Lignes_SOL = ActiveWorkbook.Sheets("LAF").Range("A65536").End(xlUp).Row
    Nom_Feuil_SOL = "LAF ss dbl_NEXTEO_KCVP_ss_db"
    String1 = """Unknown"""
    'Debug.Print "=SI(RECHERCHEV(C2" & ";'[" & Nom_fichier_SSdbl & "]" & Nom_Feuil_SOL & "'!$E:$W;19;FAUX)=0" & ";" & String1 & ";RECHERCHEV(C2;'[" & Nom_fichier_SSdbl & "]" & Nom_Feuil_SOL & "'!$E:$W;19;FAUX))"
    
    ActiveWorkbook.Sheets("LAF").Columns("X:BL").NumberFormat = "General"
    For c = 2 To Nb_Lignes_SOL
    'For c = 2 To 30
    d = 19
        For l = 23 To 54
            ActiveWorkbook.Sheets("LAF").Cells(c, l).FormulaLocal = "=RECHERCHEV(C" & c & ";'[" & Nom_fichier_SSdbl & "]" & Nom_Feuil_SOL & "'!$F:$BL;" & d & ";FAUX)"
        d = d + 1
        Next
        
    e = 57
        For l = 61 To 63
        ActiveWorkbook.Sheets("LAF").Cells(c, l).FormulaLocal = "=RECHERCHEV(C" & c & ";'[" & Nom_fichier_SSdbl & "]" & Nom_Feuil_SOL & "'!$F:$BL;" & e & ";FAUX)"
        e = e + 1
        Next
        
    Next
Else
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(20, "E").Value = "Fichier Introuvable"
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(20, "E").Interior.ColorIndex = 9
End If


'=============================================================================================================
Windows(Nom_SOL).Activate
ActiveWorkbook.Sheets("LAF").Activate
    ActiveWorkbook.Sheets("LAF").Columns("X:BL").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

'===============================================================================================================
For c = 2 To Nb_Lignes_SOL
'For c = 2 To 30

        
        For l = 23 To 54
            If ActiveWorkbook.Sheets("LAF").Cells(c, l) = "0" Then
            ActiveWorkbook.Sheets("LAF").Cells(c, l) = "Unknown"
            End If
        Next
        
        For l = 61 To 63
            If ActiveWorkbook.Sheets("LAF").Cells(c, l) = "0" Then
            ActiveWorkbook.Sheets("LAF").Cells(c, l) = "Unknown"
            End If
        Next
        
        If ActiveWorkbook.Sheets("LAF").Cells(c, "AS") = "Unknown" Then
        ActiveWorkbook.Sheets("LAF").Cells(c, "AS") = Empty
        End If
        'If ActiveWorkbook.Sheets("LAF").Cells(c, "AT") = "Unknown" Then
        'ActiveWorkbook.Sheets("LAF").Cells(c, "AT") = Empty
        'End If
        If ActiveWorkbook.Sheets("LAF").Cells(c, "BJ") = "Unknown" Then
        ActiveWorkbook.Sheets("LAF").Cells(c, "BJ") = Empty
        End If
        
    Next

'===============================================================================================================
Windows(Nom_SOL).Activate
'ActiveWorkbook.Sheets("LAF").Range("$A$1:$Q$5960").AutoFilter Field:=3
ActiveWorkbook.Sheets("Macro_MAJ_Document_SOL").Cells(2, 2) = Date
ActiveWorkbook.Sheets("Macro_MAJ_Document_SOL").Cells(2, 3) = Time
Windows(ThisWorkbook.Name).Activate
ActiveWorkbook.Sheets("Macro_MAJ_Document").Activate
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(1, 2) = "5959"
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(20, "G").Value = "Execution Terminée"
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(20, "G").Interior.ColorIndex = 9
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(21, "H") = Date & " a " & Time
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(20, "H") = Nb_Lignes_SOL
Heure_fin_SOL = Time

Durée_SOL = Heure_fin_SOL - Heure_debut_SOL
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(21, "I") = Durée_SOL


End Sub







