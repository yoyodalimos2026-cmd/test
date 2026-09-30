Attribute VB_Name = "Module2"
Sub MAJ_BORD_V3_27_10_23()

Windows(ThisWorkbook.Name).Activate
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(16, "G") = Date & " a " & Time
Heure_debut_BORD = Time
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(15, "G").Value = "Execution en cours"
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(15, "E").Value = ""

Set fso = CreateObject("Scripting.FileSystemObject")
Current_Path = ThisWorkbook.Path
Nom_BORD = ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(15, "B").Value '<===========Extraire le nom du fichier a mettre a jour
BORD_File_path = Current_Path & "\" & Nom_BORD                           '<==========================='L'adresse du fichier BORD
SSDBL_file_path = Current_Path & "\" & ThisWorkbook.Name
Nom_fichier_SSdbl = ThisWorkbook.Name
'Debug.Print Nom_fichier_SSdbl
'=======================================Vérifier que le fichier existe sosu le dossier============================================
If fso.FileExists(BORD_File_path) = True Then

    ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(15, "E").Value = "Fichier Trouvé"
    ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(15, "E").Interior.ColorIndex = 4


    Set BORD_FILE = Workbooks.Open(BORD_File_path)
    Set ssdbl_FILE = Workbooks.Open(SSDBL_file_path)
    Windows(Nom_BORD).Activate
    Nb_Lignes_BORD = ActiveWorkbook.Sheets("LAF").Range("A65536").End(xlUp).Row
    Nom_Feuil_BORD = "LAF ss dbl_NEXTEO_KCVP_ss_db"
    String1 = """Unknown"""
    'Debug.Print "=SI(RECHERCHEV(C2" & ";'[" & Nom_fichier_SSdbl & "]" & Nom_Feuil_BORD & "'!$E:$W;19;FAUX)=0" & ";" & String1 & ";RECHERCHEV(C2;'[" & Nom_fichier_SSdbl & "]" & Nom_Feuil_BORD & "'!$E:$W;19;FAUX))"
    
    ActiveWorkbook.Sheets("LAF").Columns("X:BL").NumberFormat = "General"
    For c = 2 To Nb_Lignes_BORD
    'For c = 2 To 1000
    d = 19
        For l = 23 To 54
            ActiveWorkbook.Sheets("LAF").Cells(c, l).FormulaLocal = "=RECHERCHEV(C" & c & ";'[" & Nom_fichier_SSdbl & "]" & Nom_Feuil_BORD & "'!$F:$BL;" & d & ";FAUX)"
        d = d + 1
        Next
        
    e = 57
        For l = 61 To 63
        ActiveWorkbook.Sheets("LAF").Cells(c, l).FormulaLocal = "=RECHERCHEV(C" & c & ";'[" & Nom_fichier_SSdbl & "]" & Nom_Feuil_BORD & "'!$F:$BL;" & e & ";FAUX)"
        e = e + 1
        Next
        
    Next
Else
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(15, "E").Value = "Fichier Introuvable"
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(15, "E").Interior.ColorIndex = 9
End If


'=============================================================================================================
Windows(Nom_BORD).Activate
ActiveWorkbook.Sheets("LAF").Activate
    ActiveWorkbook.Sheets("LAF").Columns("X:BL").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

'===============================================================================================================
For c = 2 To Nb_Lignes_BORD
'For c = 2 To 1000
        
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
        If ActiveWorkbook.Sheets("LAF").Cells(c, "AU") = "Unknown" Then
        ActiveWorkbook.Sheets("LAF").Cells(c, "AU") = Empty
        End If
        If ActiveWorkbook.Sheets("LAF").Cells(c, "BK") = "Unknown" Then
        ActiveWorkbook.Sheets("LAF").Cells(c, "BK") = Empty
        End If
        
    Next

'===============================================================================================================
Windows(Nom_BORD).Activate
ActiveWorkbook.Sheets("Macro_MAJ_Document_BORD").Cells(2, 2) = Date
ActiveWorkbook.Sheets("Macro_MAJ_Document_BORD").Cells(2, 3) = Time
'ActiveWorkbook.Sheets("LAF").Range("$A$1:$Q$5960").AutoFilter Field:=3
Windows(ThisWorkbook.Name).Activate
ActiveWorkbook.Sheets("Macro_MAJ_Document").Activate
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(1, 2) = "5959"
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(15, "G").Value = "Execution Terminée"
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(15, "G").Interior.ColorIndex = 9
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(16, "H") = Date & " a " & Time
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(15, "H") = Nb_Lignes_BORD
Heure_fin_BORD = Time

Durée_BORD = Heure_fin_BORD - Heure_debut_BORD
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(16, "I") = Durée_BORD



End Sub






