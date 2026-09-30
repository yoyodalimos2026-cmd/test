Attribute VB_Name = "Module7"
Sub MAJ_Lien_Document_V3_26_10_23()

'ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(2, 2) = Date
'ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(2, 3) = Time
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(7, "G") = ""
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(7, "G").Interior.ColorIndex = 2
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(7, "G") = "Execution En cours"
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(7, "G").Interior.ColorIndex = 4
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(8, "G") = "Durée d'éxecution max= 15 min"
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(8, "G").Interior.ColorIndex = 2
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(8, "G").Interior.ColorIndex = 4

'====================================Remplacement des cases vide par "-"========================================================================================
Nb_Ligne_ss_dbl = ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Range("A65536").End(xlUp).Row

'For t = 2 To Nb_Ligne_ss_dbl
    'For r = 24 To 64
    
        'If ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(t, r) = Empty Then
        'ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(t, r) = "'-"
        'End If

    'Next
'Next

'===============================================================================================================================


For Init = 3 To 8
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(Init, "l").Interior.ColorIndex = 2
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(Init, "l") = "xx"
Next

NB_Doc_ROHS = 0
NB_Doc_REACH = 0
NB_Doc_FMD_Template = 0
NB_Doc_PDN = 0
NB_Doc_dmd = 0
NB_Doc_rep = 0
Temps_debut = Time

Ligne_Min = ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(5, 5).Value
Ligne_Max = ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(5, 6).Value
'================================================================================================
Dossier_REACH = ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(5, 2).Value
Dossier_ROHS = ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(6, 2).Value
Dossier_FMD = ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(7, 2).Value
Dossier_Template = ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(8, 2).Value
Dossier_Demandes = ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(9, 2).Value
Dossier_Reponses = ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(10, 2).Value
Dossier_PDN = ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(11, 2).Value
'=============================================================================================
For f = Ligne_Min To Ligne_Max
ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(f, 1).Interior.ColorIndex = 2
Next
'=============================================================================================
Set fso = CreateObject("Scripting.FileSystemObject")
For i = Ligne_Min To Ligne_Max
'============================================Identification des lignes mis a jours
ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, 1).Interior.ColorIndex = 3
'=============================================Initialisation des couleurs
ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "AP").Interior.ColorIndex = 2 'Couleur blanche
ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "AQ").Interior.ColorIndex = 2
ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "AR").Interior.ColorIndex = 2
ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "BJ").Interior.ColorIndex = 2
ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "BK").Interior.ColorIndex = 2
ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "BL").Interior.ColorIndex = 2

'============================================'Recupére les lien hypertext Documents justificatifs ROHS ===========================================
    Nom_lien_ROHS = ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "AP").Value
    Nom_lien = Chr(34) & Nom_lien_ROHS & Chr(34)
    Current_Path = ThisWorkbook.Path
    
'================================================Test si le nom contient les mots clés ROHS=======================================================
    If InStr(1, Nom_lien_ROHS, "ROHS") <> 0 Or InStr(1, Nom_lien_ROHS, "RoHS") <> 0 Then
    code_article_path = "." & "\" & Dossier_ROHS & "\" & Nom_lien_ROHS
    ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "AP").FormulaLocal = "=LIEN_HYPERTEXTE(" & Chr(34) & code_article_path & Chr(34) & ";" & Nom_lien & ")"
    End If
    
'================================================'Vérifier si le document existe'================================================'================================================
    ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "AP").Interior.ColorIndex = 2
    ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "AP").Borders.Color = RGB(0, 0, 0)
    
    If ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "AP") <> "-" And ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "AP").Value <> "" Then
    Code_art_path_verif_ROHS = Current_Path & "\" & Dossier_ROHS & "\" & Nom_lien_ROHS

        If fso.FileExists(Code_art_path_verif_ROHS) = False Then
        ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "AP").Interior.ColorIndex = 3
        NB_Doc_ROHS = NB_Doc_ROHS + 1
        ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(3, "l").Interior.ColorIndex = 3
        Else
        ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "AP").Interior.ColorIndex = 2
        End If
    End If


'============================================'Recupére les lien hypertext Documents justificatifs FMD / template===========================================
    Nom_lien_FMD = ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "AR").Value
    Nom_lien = Chr(34) & Nom_lien_FMD & Chr(34)
    Current_Path = ThisWorkbook.Path
    
'================================================Test si le nom contient les mots clés FMD/FDS=======================================================
    If InStr(1, Nom_lien_FMD, "FMD") <> 0 Or InStr(1, Nom_lien_FMD, "FDS") Then
    code_article_path = "." & "\" & Dossier_FMD & "\" & Nom_lien_FMD
    ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "AR").FormulaLocal = "=LIEN_HYPERTEXTE(" & Chr(34) & code_article_path & Chr(34) & ";" & Nom_lien & ")"
    ElseIf InStr(1, Nom_lien_FMD, "Template") <> 0 Then
    code_article_path_temp = "." & "\" & Dossier_Template & "\" & Nom_lien_FMD
    ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "AR").FormulaLocal = "=LIEN_HYPERTEXTE(" & Chr(34) & code_article_path_temp & Chr(34) & ";" & Nom_lien & ")"
    End If
    
    
'================================================'Vérifier si le document existe'================================================'================================================
    ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "AR").Interior.ColorIndex = 2
    ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "AR").Borders.Color = RGB(0, 0, 0)
    
    If ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "AR") <> "-" And ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "AR").Value <> "" Then
    Code_art_path_verif_FMD = Current_Path & "\" & Dossier_FMD & "\" & Nom_lien_FMD
    Code_art_path_verif_Template = Current_Path & "\" & Dossier_Template & "\" & Nom_lien_FMD
        
        If fso.FileExists(Code_art_path_verif_FMD) = False And fso.FileExists(Code_art_path_verif_Template) = False Then
        ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "AR").Interior.ColorIndex = 3
        NB_Doc_FMD_Template = NB_Doc_FMD_Template + 1
        ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(5, "l").Interior.ColorIndex = 3
        Else
        ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "AR").Interior.ColorIndex = 2
        End If
    End If
    
'=============================================================================================================================================================
'===============================================Recupére les lien hypertext des Documents justificatifs REACH COFC=============================================
    Nom_lien_REACH = ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "AQ").Value
    Nom_lien = Chr(34) & Nom_lien_REACH & Chr(34)
    Current_Path = ThisWorkbook.Path
    
'=======================================Test si le nom contient les mots clés FMD/FDS/COFC/ROHS===============================================================

    If InStr(1, Nom_lien_REACH, "COFC") <> 0 Then
    code_article_path = "." & "\" & Dossier_REACH & "\" & Nom_lien_REACH
    ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "AQ").FormulaLocal = "=LIEN_HYPERTEXTE(" & Chr(34) & code_article_path & Chr(34) & ";" & Nom_lien & ")"
    End If
    
'======================================= Vérifier si le document existe'======================================='=======================================
    ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "AQ").Interior.ColorIndex = 2
    ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "AQ").Borders.Color = RGB(0, 0, 0)
    If ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "AQ") <> "-" And ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "AQ").Value <> "" Then
    Code_art_path_verif_REACH = Current_Path & "\" & Dossier_REACH & "\" & Nom_lien_REACH

        If fso.FileExists(Code_art_path_verif_REACH) = False Then
        ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "AQ").Interior.ColorIndex = 3
        NB_Doc_REACH = NB_Doc_REACH + 1
        ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(4, "l").Interior.ColorIndex = 3
        Else
        ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "AQ").Interior.ColorIndex = 2
        End If

    End If
'======================================================================================================================================================================
'==========================================='Recupére les lien hypertext des demandes envoyées'===========================================
Colonne_Etat = ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "AZ").Value
Colonne_Cmnt = ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "BB").Value

'If Colonne_Etat <> "A ne pas traiter" And Colonne_Etat <> "Aucun traitement a faire" Then

If Colonne_Etat <> "A ne pas traiter" Then
    Nom_lien_DMD = ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "BJ").Value
    Nom_lien = Chr(34) & Nom_lien_DMD & Chr(34)
    Current_Path = ThisWorkbook.Path
    DMD_path = "." & "\" & Dossier_Demandes & "\" & Nom_lien_DMD
    ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "BJ").FormulaLocal = "=LIEN_HYPERTEXTE(" & Chr(34) & DMD_path & Chr(34) & ";" & Nom_lien & ")"
    
'===========================================''Vérifier si le document existe'==========================================='
    ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "BJ").Interior.ColorIndex = 2
    ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "BJ").Borders.Color = RGB(0, 0, 0)
    If ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "BJ") <> "-" And ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "BJ").Value <> "" And ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "BJ").Value <> "Envoi_KO" Then
        DMD_path_verif = Current_Path & "\" & Dossier_Demandes & "\" & Nom_lien_DMD
        
        If fso.FileExists(DMD_path_verif) = False Then
        ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "BJ").Interior.ColorIndex = 3
        NB_Doc_dmd = NB_Doc_dmd + 1
        ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(6, "l").Interior.ColorIndex = 3
        End If
    End If
    
Else

'ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "BJ").Value = "Envoi_KO"
'ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "BK").Value = "Reponse_KO"

End If

'==========================================='''Recupére les lien hypertext des réponses recues'===========================================''
Colonne_Etat = ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "BA").Value
Colonne_Cmnt = ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "BC").Value

If InStr(1, Colonne_Cmnt, "Réponse recue(") Or InStr(1, Colonne_Cmnt, "Réponse reçue(") Or InStr(1, Colonne_Cmnt, "Réponse reçue") Then

    Nom_lien_REP = ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "BK").Value
    Nom_lien = Chr(34) & Nom_lien_REP & Chr(34)
    Current_Path = ThisWorkbook.Path
    REP_path = "." & "\" & Dossier_Reponses & "\" & Nom_lien_REP
    ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "BK").FormulaLocal = "=LIEN_HYPERTEXTE(" & Chr(34) & REP_path & Chr(34) & ";" & Nom_lien & ")"
    
'===========================================''Vérifier si le document existe'==========================================='
    If ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "BK") <> "-" And ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "BK").Value <> "" And ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "BK").Value <> "Reponse_KO" Then
        REP_path_verif = Current_Path & "\" & Dossier_Reponses & "\" & Nom_lien_REP
        'Debug.Print "REP_path_verif", REP_path_verif
        If fso.FileExists(REP_path_verif) = False Then
        ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "BK").Interior.ColorIndex = 3
        ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(7, "l").Interior.ColorIndex = 3
        NB_Doc_rep = NB_Doc_rep + 1
        End If
    End If
    
Else

    ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "BK").Interior.ColorIndex = 2
    ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "BK").Borders.Color = RGB(0, 0, 0)
    
    If ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "BK").Value = "Reponse_KO" Then
        ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "BK").Interior.ColorIndex = 18
    End If

End If
'===========================================''===========================================''===========================================''==========================================='

'===========================================''Recupére les lien hypertext Documents PDN'==========================================='

    Nom_lien_PDN = ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "BL").Value
    Nom_lien = Chr(34) & Nom_lien_PDN & Chr(34)
    Current_Path = ThisWorkbook.Path
    
'===========================================''Test si le nom contient les mots clés FMD/FDS/COFC/ROHS'==========================================='
    If InStr(1, Nom_lien_PDN, "PDN") <> 0 Then
    code_article_path = "." & "\" & Dossier_PDN & "\" & Nom_lien_PDN
    ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "BL").FormulaLocal = "=LIEN_HYPERTEXTE(" & Chr(34) & code_article_path & Chr(34) & ";" & Nom_lien & ")"
    End If
    
'===========================================''Vérifier si le document existe'==========================================='
    ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "BL").Interior.ColorIndex = 2
    ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "BL").Borders.Color = RGB(0, 0, 0)
    If ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "BL") <> "-" And ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "BL").Value <> "" Then
    Code_art_path_verif = Current_Path & "\" & Dossier_PDN & "\" & Nom_lien_PDN
    
        If fso.FileExists(Code_art_path_verif) = False Then
        ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(i, "BL").Interior.ColorIndex = 3
        NB_Doc_PDN = NB_Doc_PDN + 1
        ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(8, "l").Interior.ColorIndex = 3
        End If
    End If

Next i

Temps_fin = Time
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(7, "G") = "Execution Terminée"
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(7, "G").Interior.ColorIndex = 9
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(8, "G") = Date & " a " & Time
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(8, "G").Interior.ColorIndex = 2
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(8, "H") = Date & " a " & Time
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(8, "I") = Temps_fin - Temps_debut

ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(3, "L") = NB_Doc_ROHS
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(4, "L") = NB_Doc_REACH
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(5, "L") = NB_Doc_FMD_Template
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(6, "L") = NB_Doc_dmd
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(7, "L") = NB_Doc_rep
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(8, "L") = NB_Doc_PDN

End Sub




