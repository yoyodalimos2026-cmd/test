Attribute VB_Name = "Module4"
Sub MAJ_KCVP_V3_08_11_2023()


Windows(ThisWorkbook.Name).Activate
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(26, "G") = Date & " a " & Time
Heure_debut_KCVP = Time
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(25, "G").Value = "Execution en cours"
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(25, "G").Interior.ColorIndex = 4
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(25, "E").Value = ""

Set fso = CreateObject("Scripting.FileSystemObject")
Current_Path = ThisWorkbook.Path
Nom_KCVP = ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(25, "B").Value '<===========Extraire le nom du fichier a mettre a jour
KCVP_File_path = Current_Path & "\" & Nom_KCVP                           '<==========================='L'adresse du fichier KCVP
SSDBL_file_path = Current_Path & "\" & ThisWorkbook.Name
Nom_fichier_SSdbl = ThisWorkbook.Name
'Debug.Print Nom_fichier_SSdbl
'=======================================Vérifier que le fichier existe sosu le dossier============================================
If fso.FileExists(KCVP_File_path) = True Then

    ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(25, "E").Value = "Fichier Trouvé"
    ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(25, "E").Interior.ColorIndex = 4


    Set KCVP_FILE = Workbooks.Open(KCVP_File_path)
    Set ssdbl_FILE = Workbooks.Open(SSDBL_file_path)
    Windows(Nom_KCVP).Activate
    Nb_Lignes_KCVP = ActiveWorkbook.Sheets("LAF complet").Range("A65536").End(xlUp).Row
    Nom_Feuil_KCVP = "LAF ss dbl_NEXTEO_KCVP_ss_db"
    String1 = """Unknown"""
    'Debug.Print "=SI(RECHERCHEV(C2" & ";'[" & Nom_fichier_SSdbl & "]" & Nom_Feuil_KCVP & "'!$E:$W;19;FAUX)=0" & ";" & String1 & ";RECHERCHEV(C2;'[" & Nom_fichier_SSdbl & "]" & Nom_Feuil_KCVP & "'!$E:$W;19;FAUX))"
    
    ActiveWorkbook.Sheets("LAF complet").Columns("AC:BH").NumberFormat = "General"
    For c = 2 To Nb_Lignes_KCVP
    'For c = 2 To 30
    d = 19
        For l = 29 To 60
            ActiveWorkbook.Sheets("LAF complet").Cells(c, l).FormulaLocal = "=RECHERCHEV(F" & c & ";'[" & Nom_fichier_SSdbl & "]" & Nom_Feuil_KCVP & "'!$F:$BL;" & d & ";FAUX)"
        d = d + 1
        Next
        
    e = 57
        For l = 67 To 69
        ActiveWorkbook.Sheets("LAF complet").Cells(c, l).FormulaLocal = "=RECHERCHEV(F" & c & ";'[" & Nom_fichier_SSdbl & "]" & Nom_Feuil_KCVP & "'!$F:$BL;" & e & ";FAUX)"
        e = e + 1
        Next
        
    Next
Else
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(20, "E").Value = "Fichier Introuvable"
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(20, "E").Interior.ColorIndex = 9
End If


'=============================================================================================================
Windows(Nom_KCVP).Activate
ActiveWorkbook.Sheets("LAF complet").Activate
    ActiveWorkbook.Sheets("LAF complet").Columns("AC:BH").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

'===============================================================================================================
For c = 2 To Nb_Lignes_KCVP
'For c = 2 To 30

        
        For l = 29 To 58
            If ActiveWorkbook.Sheets("LAF complet").Cells(c, l) = "0" Then
            ActiveWorkbook.Sheets("LAF complet").Cells(c, l) = "Unknown"
            End If
        Next
        
        For l = 65 To 69
            If ActiveWorkbook.Sheets("LAF complet").Cells(c, l) = "0" Then
            ActiveWorkbook.Sheets("LAF complet").Cells(c, l) = "Unknown"
            End If
        Next
        
        If ActiveWorkbook.Sheets("LAF complet").Cells(c, "AY") = "Unknown" Then
        ActiveWorkbook.Sheets("LAF complet").Cells(c, "AY") = Empty
        End If
        If ActiveWorkbook.Sheets("LAF complet").Cells(c, "AZ") = "Unknown" Then
        ActiveWorkbook.Sheets("LAF complet").Cells(c, "AZ") = Empty
        End If
        If ActiveWorkbook.Sheets("LAF complet").Cells(c, "BP") = "Unknown" Then
        ActiveWorkbook.Sheets("LAF complet").Cells(c, "BP") = Empty
        End If
        
    Next

'===============================================================================================================
Windows(Nom_KCVP).Activate
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(2, 2) = Date
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(2, 3) = Time
'ActiveWorkbook.Sheets("LAF").Range("$A$1:$Q$5960").AutoFilter Field:=3
Windows(ThisWorkbook.Name).Activate
ActiveWorkbook.Sheets("Macro_MAJ_Document").Activate
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(1, 2) = "9643"
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(25, "G").Value = "Execution Terminée"
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(25, "G").Interior.ColorIndex = 9
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(26, "H") = Date & " a " & Time
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(25, "H") = Nb_Lignes_KCVP
Heure_fin_KCVP = Time

Durée_KCVP = Heure_fin_KCVP - Heure_debut_KCVP
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(26, "I") = Durée_KCVP


End Sub








