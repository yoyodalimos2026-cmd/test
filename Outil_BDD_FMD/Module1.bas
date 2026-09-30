Attribute VB_Name = "Module1"
Sub Lien_Document_26_01_2024()
Attribute Lien_Document_26_01_2024.VB_ProcData.VB_Invoke_Func = " \n14"
ActiveWorkbook.Save
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(3, "A").Interior.ColorIndex = 2
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(3, "A") = "Execution En cours"
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(3, "A").Interior.ColorIndex = 4
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(4, "A") = Date & " a " & Time


'===============================Supprimer les doublons=========================================================================
    ActiveWorkbook.Sheets("BDD_FMD").Activate
    Columns("I:I").Select
    Range("I20000").Activate
    ActiveSheet.Range("$A$1:$V$20000").RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5, _
        6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22), Header:=xlYes
ActiveWorkbook.Sheets("Macro_MAJ_Document").Activate
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(1, 1).Select

Call MAJ_Etat_FMD_BDD
'====================================Remplacement des cases vide par "-"========================================================================================
Nb_Ligne_BDD_FMD = ActiveWorkbook.Sheets("BDD_FMD").Range("F200000").End(xlUp).Row
'Nb_Ligne_BDD_FMD = ActiveWorkbook.Sheets("BDD_FMD").UsedRange.Rows.Count
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(13, "B") = Nb_Ligne_BDD_FMD
'===============================================================================================================================
'For Init = 3 To 8
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(10, "B").Interior.ColorIndex = 2
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(10, "B") = "xx"
'Next
NB_Doc_FMD_Template = 0
Temps_debut = Time
Ligne_Min = ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(15, "B").Value
Ligne_Max = ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(16, "B").Value
Dossier_FMD = ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(7, 2).Value
Nb_Lignes_Data = ActiveWorkbook.Sheets("LU_FMD").Range("A65536").End(xlUp).Row
lgn = Nb_Lignes_Data + 1
'=============================================================================================
Set fso = CreateObject("Scripting.FileSystemObject")
For i = Ligne_Min To Ligne_Max
'For i = 2 To Nb_Ligne_BDD_FMD

'============================================Identification des lignes mis a jours
ActiveWorkbook.Sheets("BDD_FMD").Cells(i, 1).Interior.ColorIndex = 3
'=============================================Initialisation des couleurs
ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "I").Interior.ColorIndex = 2
ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "I").Interior.ColorIndex = 2
ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "J") = ""
ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "J").Interior.ColorIndex = 2
ActiveWorkbook.Sheets("LU_FMD").Range("A" & Ligne_Min & ":K" & Ligne_Max).Borders.LineStyle = xlContinuous
ActiveWorkbook.Sheets("BDD_FMD").Range("A" & Ligne_Min & ":V" & Ligne_Max).Borders.LineStyle = xlContinuous
'============================================'Recupére les lien hypertext Documents justificatifs FMD / template===========================================
    Nom_lien_FMD = ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "I").Value
    Nom_lien = Chr(34) & Nom_lien_FMD & Chr(34)
    Nom_lien_FMD_ = Replace(ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "I").Value, " ", "_")
    Nom_lien_FMD_ = Replace(Nom_lien_FMD, " ", "_")
    Nom_lien_ = Chr(34) & Nom_lien_FMD_ & Chr(34)
    'Debug.Print ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "I").Value
    code_article_path_old = "." & "\" & Dossier_FMD & "\" & Nom_lien_FMD
    Current_path = ThisWorkbook.Path
'================================================Test si le nom contient les mots clés FMD/FDS=======================================================
    If InStr(1, Nom_lien_FMD, ".pdf") <> 0 Or InStr(1, Nom_lien_FMD, ".PDF") <> 0 Then
    
        If InStr(1, Nom_lien_FMD, " ") <> 0 Then
        
        ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "I").Value = Nom_lien_FMD_
        code_article_path_new = "." & "\" & Dossier_FMD & "\" & Nom_lien_FMD_
        ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "I").FormulaLocal = "=LIEN_HYPERTEXTE(" & Chr(34) & code_article_path_new & Chr(34) & ";" & Nom_lien_ & ")"
        Else
        code_article_path_new = "." & "\" & Dossier_FMD & "\" & Nom_lien_FMD_
        ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "I").FormulaLocal = "=LIEN_HYPERTEXTE(" & Chr(34) & code_article_path_new & Chr(34) & ";" & Nom_lien_ & ")"
        End If
        
    ElseIf InStr(1, Nom_lien_FMD, ".pdf") = 0 And ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "I") <> Empty Then
    ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "I") = ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "I") & ".pdf"
    Nom_lien_FMD_ = ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "I")
        If InStr(1, Nom_lien_FMD, " ") <> 0 Then
        ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "I").Value = Nom_lien_FMD_
        Nom_lien_ = Chr(34) & Nom_lien_FMD_ & Chr(34)
        code_article_path_new = "." & "\" & Dossier_FMD & "\" & Nom_lien_FMD_
        ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "I").FormulaLocal = "=LIEN_HYPERTEXTE(" & Chr(34) & code_article_path_new & Chr(34) & ";" & Nom_lien_ & ")"
        Else
        ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "I").Value = Nom_lien_FMD_
        Nom_lien_ = Chr(34) & Nom_lien_FMD_ & Chr(34)
        code_article_path_new = "." & "\" & Dossier_FMD & "\" & Nom_lien_FMD_
        ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "I").FormulaLocal = "=LIEN_HYPERTEXTE(" & Chr(34) & code_article_path_new & Chr(34) & ";" & Nom_lien_ & ")"
        End If
        
    ElseIf ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "I") = Empty Then
    ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "J") = "Pas de FMD associé"
    End If
    
    
'================================================'Vérifier si le document existe'================================================'================================================
    ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "I").Interior.ColorIndex = 2
    ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "I").Borders.Color = RGB(0, 0, 0)
    
    If ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "I") <> "-" And ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "I").Value <> "" Then
    Code_art_path_old = Current_path & "\" & Dossier_FMD & "\" & Nom_lien_FMD
    code_article_path_new_name = Current_path & "\" & Dossier_FMD & "\" & Nom_lien_FMD_
    'Debug.Print Code_art_path_old
    'Debug.Print code_article_path_new_name
        
        If fso.FileExists(Code_art_path_old) = False Then
            If fso.FileExists(code_article_path_new_name) = False Then
            ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "I").Interior.ColorIndex = 3
            ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "J").Interior.ColorIndex = 3
            NB_Doc_FMD_Template = NB_Doc_FMD_Template + 1
            ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(10, "B").Interior.ColorIndex = 3
            ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "J") = "FMD introuvable sous le dossier FMD"
            Else
            ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "I").Interior.ColorIndex = 2
            ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "J") = ""
            End If
        Else
            'Debug.Print Code_art_path_old
            'Debug.Print code_article_path_new_name
            
            If fso.FileExists(code_article_path_new_name) = False Then
            Name Code_art_path_old As code_article_path_new_name
            Nom_lien_FMD_ = Replace(ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "I").Value, " ", "_")
            End If
            
            If fso.FileExists(code_article_path_new_name) = True Then
            ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "I").Interior.ColorIndex = 2
            'ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "J") = ""
            Else
            ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "I").Interior.ColorIndex = 3
            ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "J").Interior.ColorIndex = 3
            NB_Doc_FMD_Template = NB_Doc_FMD_Template + 1
            ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(10, "B").Interior.ColorIndex = 3
            ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "J") = "FMD introuvable sous le dossier FMD"
            End If
        End If
        
    End If
    correspondance_ref = 1
    correspondance_fab = 1
    'Debug.Print Mid(ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "F"), 1, 4)
    'Debug.Print UCase(ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "I"))
'===========================================''Remplissage Colonne J (Etat FMD)''==========================================='
'Debug.Print UCase(ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "I")), Replace(UCase(ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "E")), " ", "_")

    If InStr(UCase(ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "I")), Replace(UCase(ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "E")), " ", "_")) = 0 _
    And ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "I") <> Empty Then
    ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "J") = ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "J") & "/Vérif Nom Fab"
    ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "J").Interior.ColorIndex = 46
    Else
    correspondance_fab = 0
    'ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "J") = ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "J") & "/Nom Fab OK"
    End If
    
    If InStr(UCase(ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "I")), Mid(ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "F"), 1, 3)) = 0 _
    And ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "I") <> Empty Then
    ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "J") = ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "J") & "/Vérif Nom Ref"
    ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "J").Interior.ColorIndex = 46
    Else
    correspondance_ref = 0
    'ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "J") = ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "J") & "/Nom Ref OK"
    End If

    'If InStr(UCase(ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "I")), ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "C")) = 0 _
    'And ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "I") <> Empty And correspondance_ref = 1 And correspondance_fab = 1 Then
    If InStr(UCase(ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "I")), ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "C")) = 0 _
    And ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "I") <> Empty Then
    'ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "J") = ""
    ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "J") = ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "J") & "/Vérif Code Article"
    ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "J").Interior.ColorIndex = 46
    End If
'===========================================''===========================================''==========================================='
    If InStr(ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "I"), "MDS") <> 0 Then
    ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "K") = ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "C")
    ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "L") = ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "E")
    ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "A") = ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "C") & ";" & ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "E") & ";" & ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "F")
    ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "G") = ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "E") & ";" & ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "F")
    Else
    ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "K") = ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "F")
    ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "L") = ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "E")
    ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "A") = ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "C") & ";" & ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "E") & ";" & ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "F")
    ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "G") = ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "E") & ";" & ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "F")
    End If

'===========================================''===========================================''==========================================='


ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "M") = ""

If ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "I") <> Empty Then
    If ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "M") <> "Pas encore analysé" And InStr(ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "J"), "FMD introuvable sous le dossier FMD") = 0 Then
        ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "M").FormulaLocal = "=RECHERCHEV(I" & i & ";'" & "LU_FMD" & "'!A:K;7;FAUX)"
        ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "M").NumberFormat = "@"
        ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "M").NumberFormat = "mm/dd/yyyy"
            
            If Application.IsNA(ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "M")) Then
            
            ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "M") = "Pas encore analysé"
            
            ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "I").Copy
            'ActiveWorkbook.Sheets("LU_FMD").Cells(i, "I").FormulaLocal = "=LIEN_HYPERTEXTE(" & Chr(34) & code_article_path_new & Chr(34) & ";" & Nom_lien_ & ")"
            ActiveWorkbook.Sheets("LU_FMD").Cells(lgn, "A").PasteSpecial Paste:=xlPasteFormulas
            ActiveWorkbook.Sheets("LU_FMD").Cells(lgn, "A").PasteSpecial Paste:=xlPasteFormats
            ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "J").Copy
            ActiveWorkbook.Sheets("LU_FMD").Cells(lgn, "B").PasteSpecial Paste:=xlPasteFormulas
            ActiveWorkbook.Sheets("LU_FMD").Cells(lgn, "B").PasteSpecial Paste:=xlPasteFormats
    
                If InStr(ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "I"), "MDS") = 0 Then
                ActiveWorkbook.Sheets("LU_FMD").Cells(lgn, "C").Value = "'" & Mid(ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "K").Value, 1, 4)
                ActiveWorkbook.Sheets("LU_FMD").Cells(lgn, "C").PasteSpecial Paste:=xlPasteFormats
                ActiveWorkbook.Sheets("LU_FMD").Range("A" & lgn & ":K" & lgn).Borders.LineStyle = xlContinuous
                
                ElseIf InStr(ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "I"), "MDS") <> 0 Then
                ActiveWorkbook.Sheets("LU_FMD").Cells(lgn, "C") = ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "K")
                ActiveWorkbook.Sheets("LU_FMD").Range("A" & lgn & ":K" & lgn).Borders.LineStyle = xlContinuous
                'ActiveWorkbook.Sheets("LU_FMD").Cells(lgn, "C").PasteSpecial Paste:=xlPasteFormats
                End If
                
            lgn = lgn + 1

            ElseIf ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "M") = 0 Then
            ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "M").FormulaLocal = "=RECHERCHEV(I" & i & ";'" & "LU_FMD" & "'!A:K;2;FAUX)"
            ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "M") = "Pas encore analysé"
            ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "M").NumberFormat = "@"
            ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "M").NumberFormat = "mm/dd/yyyy"
            ActiveWorkbook.Sheets("LU_FMD").Range("A" & lgn & ":K" & lgn).Borders.LineStyle = xlContinuous
            'lgn = lgn + 1

            End If
    End If
   
ActiveWorkbook.Sheets("LU_FMD").Range("A" & lgn & ":K" & lgn).Borders.LineStyle = xlContinuous
Else
ActiveWorkbook.Sheets("BDD_FMD").Cells(i, "M") = "Pas de FMD"
End If


Next i

ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(10, "B") = NB_Doc_FMD_Template
'Call Liste_Unique_FMD_29_01_2024
'Call Ajout_FMD

Temps_fin = Time
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(3, "A") = "Execution Terminée"
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(3, "A").Interior.ColorIndex = 9
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(4, "A") = Date & " a " & Time
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(4, "B") = Temps_fin - Temps_debut
Nb_Lignes_Data = ActiveWorkbook.Sheets("LU_FMD").Range("A65536").End(xlUp).Row
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(28, "B") = Nb_Lignes_Data
Nb_Ligne_BDD_FMD = ActiveWorkbook.Sheets("BDD_FMD").Range("F200000").End(xlUp).Row
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(13, "B") = Nb_Ligne_BDD_FMD
ActiveWorkbook.Save



End Sub
