Attribute VB_Name = "Module8"
Sub MAJ_MOBILITY_V1_09_11_2023()


Windows(ThisWorkbook.Name).Activate
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(31, "G") = Date & " a " & Time
Heure_debut_MOBILITY = Time
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(30, "G").Value = "Execution en cours"
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(30, "E").Value = ""

Set fso = CreateObject("Scripting.FileSystemObject")
Current_Path = ThisWorkbook.Path
Nom_MOBILITY = ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(30, "B").Value '<===========Extraire le nom du fichier a mettre a jour
MOBILITY_File_path = Current_Path & "\" & Nom_MOBILITY                           '<==========================='L'adresse du fichier MOBILITY
SSDBL_file_path = Current_Path & "\" & ThisWorkbook.Name
Nom_fichier_SSdbl = ThisWorkbook.Name
'Debug.Print Nom_fichier_SSdbl
'=======================================Vérifier que le fichier existe sosu le dossier============================================
If fso.FileExists(MOBILITY_File_path) = True Then

    ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(30, "E").Value = "Fichier Trouvé"
    ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(30, "E").Interior.ColorIndex = 4


    Set MOBILITY_FILE = Workbooks.Open(MOBILITY_File_path)
    Set ssdbl_FILE = Workbooks.Open(SSDBL_file_path)
    Windows(Nom_MOBILITY).Activate
    Nb_Lignes_MOBILITY = ActiveWorkbook.Sheets("LAF").Range("A65536").End(xlUp).Row
    Nom_Feuil_MOBILITY = "LAF ss dbl_NEXTEO_KCVP_ss_db"
    String1 = """Unknown"""
    'Debug.Print "=SI(RECHERCHEV(C2" & ";'[" & Nom_fichier_SSdbl & "]" & Nom_Feuil_MOBILITY & "'!$E:$W;19;FAUX)=0" & ";" & String1 & ";RECHERCHEV(C2;'[" & Nom_fichier_SSdbl & "]" & Nom_Feuil_MOBILITY & "'!$E:$W;19;FAUX))"
    
    ActiveWorkbook.Sheets("LAF").Columns("X:BL").NumberFormat = "General"
    For c = 2 To Nb_Lignes_MOBILITY
    'For c = 2 To 100
    
    d = 19
        For l = 23 To 54
            ActiveWorkbook.Sheets("LAF").Cells(c, l).FormulaLocal = "=RECHERCHEV(C" & c & ";'[" & Nom_fichier_SSdbl & "]" & Nom_Feuil_MOBILITY & "'!$F:$BL;" & d & ";FAUX)"
        d = d + 1
        Next
        
    e = 57
        For l = 61 To 63
        ActiveWorkbook.Sheets("LAF").Cells(c, l).FormulaLocal = "=RECHERCHEV(C" & c & ";'[" & Nom_fichier_SSdbl & "]" & Nom_Feuil_MOBILITY & "'!$F:$BL;" & e & ";FAUX)"
        e = e + 1
        Next
        
    Next
Else
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(15, "E").Value = "Fichier Introuvable"
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(15, "E").Interior.ColorIndex = 9
End If


'=============================================================================================================
Windows(Nom_MOBILITY).Activate
ActiveWorkbook.Sheets("LAF").Activate
    ActiveWorkbook.Sheets("LAF").Columns("X:BL").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

'===============================================================================================================
For c = 2 To Nb_Lignes_MOBILITY
'For c = 2 To 100
        
        For l = 23 To 54
            On Error Resume Next
            If Application.WorksheetFunction.IsNA(ActiveWorkbook.Sheets("LAF").Cells(c, l)) Then
            'ActiveWorkbook.Sheets("LAF").Cells(c, l).interiror.ColorIndex = 2
            ActiveWorkbook.Sheets("LAF").Cells(c, l) = "Introuvable"
            'ActiveWorkbook.Sheets("LAF").Cells(c, l).interiror.ColorIndex = 6
            End If
            
            If ActiveWorkbook.Sheets("LAF").Cells(c, l) = "0" Then
            ActiveWorkbook.Sheets("LAF").Cells(c, l) = "Unknown"
            End If
        Next
        
        For l = 61 To 63
            On Error Resume Next
            If Application.WorksheetFunction.IsNA(ActiveWorkbook.Sheets("LAF").Cells(c, l)) Then
            'ActiveWorkbook.Sheets("LAF").Cells(c, l).interiror.ColorIndex = 2
            ActiveWorkbook.Sheets("LAF").Cells(c, l) = "Introuvable"
            'ActiveWorkbook.Sheets("LAF").Cells(c, l).interiror.ColorIndex = 6
            End If
            
            If ActiveWorkbook.Sheets("LAF").Cells(c, l) = "0" Then
            ActiveWorkbook.Sheets("LAF").Cells(c, l) = "Unknown"
            End If
        Next
        
        If ActiveWorkbook.Sheets("LAF").Cells(c, "AS") = "Unknown" Then
        ActiveWorkbook.Sheets("LAF").Cells(c, "AS") = Empty
        End If
        If ActiveWorkbook.Sheets("LAF").Cells(c, "AU") = "Unknown" Then
        ActiveWorkbook.Sheets("LAF").Cells(c, "AU") = Empty
        End If
        If ActiveWorkbook.Sheets("LAF").Cells(c, "BK") = "Unknown" Then
        ActiveWorkbook.Sheets("LAF").Cells(c, "BK") = Empty
        End If
        
    Next

'===============================================================================================================
Windows(Nom_MOBILITY).Activate
ActiveWorkbook.Sheets("Macro_MAJ_Document_MOBILITY").Cells(2, 2) = Date
ActiveWorkbook.Sheets("Macro_MAJ_Document_MOBILITY").Cells(2, 3) = Time
'ActiveWorkbook.Sheets("LAF").Range("$A$1:$Q$5960").AutoFilter Field:=3
Windows(ThisWorkbook.Name).Activate
ActiveWorkbook.Sheets("Macro_MAJ_Document").Activate
'ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(1, 2) = "5959"
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(30, "G").Value = "Execution Terminée"
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(30, "G").Interior.ColorIndex = 9
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(31, "H") = Date & " a " & Time
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(30, "H") = Nb_Lignes_MOBILITY
Heure_fin_MOBILITY = Time

Durée_MOBILITY = Heure_fin_MOBILITY - Heure_debut_MOBILITY
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(31, "I") = Durée_MOBILITY



End Sub















