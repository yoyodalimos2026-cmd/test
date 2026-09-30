Attribute VB_Name = "Module6"
Sub Macro_NM_Synthese_consultation_20_11_2023_V1()

ActiveWorkbook.Sheets("NM Synthese_consultation").Cells(1, "AE") = ""
ActiveWorkbook.Sheets("NM Synthese_consultation").Cells(1, "AE").Interior.ColorIndex = 2
ActiveWorkbook.Sheets("NM Synthese_consultation").Cells(1, "AE") = "Execution En cours" & vbCrLf & " temps d'execution 4h"
ActiveWorkbook.Sheets("NM Synthese_consultation").Cells(1, "AE").Interior.ColorIndex = 4
'1518
Temps_debut = Time

For y = 2 To 2982
'For y = 2 To 500
Nb_artcile = 0
ActiveWorkbook.Sheets("NM Synthese_consultation").Activate
Article_N = ActiveWorkbook.Sheets("NM Synthese_consultation").Cells(y, "E").Value
    Range("E2").Select
    Selection.Copy
    ActiveWorkbook.Sheets("Etat de consultation").Activate
    Sheets("Etat de consultation").Select
    ActiveSheet.Range("$A$1:$U$8138").AutoFilter Field:=8, Criteria1:=Article_N
    'Debug.Print Article_N
    
    Dim cel As Range
    'Plage du filtre la colonne D
    Plage = 0
    last_row = ActiveWorkbook.Sheets("Etat de consultation").Cells(Rows.Count, "H").End(xlUp).Row 'l'indice de la derniére ligne suite au filtre
    'Set Plage = ActiveWorkbook.Sheets("Etat de consultation").Range("G2:G" & last).SpecialCells(xlCellTypeVisible)
    'Debug.Print last_row
    'Debug.Print Range("E2:E" & last_row).Count
    
    Application.Volatile True
        Nb_Envoi = 0
        Nb_Rep = 0
        Nb_REACH_Oui = 0
        Nb_REACH_Non = 0
        Nb_REACH_Unknown = 0
        Nb_AnnexeXIV_Oui = 0
        Nb_AnnexeXIV_Non = 0
        Nb_AnnexeXIV_Unknown = 0
        Nb_ligne_REACH_Anx = 0
    
    For Each rngcell In Range("H2:H" & last_row)
        If rngcell.EntireRow.Hidden = False Then
            If Cells(rngcell.Row, "S") <> "A ne pas traiter" Then
                nextVisibleCell = rngcell.Value
                Debug.Print rngcell
                Debug.Print rngcell.Row
'==============================================='Nombre de références liée a cet Article================================================
                Nb_artcile = Nb_artcile + 1
'===============================================''nombre de composants sources consultés=======''==============================================='
                If Cells(rngcell.Row, "T") = "Envoi_OK" Then
                Nb_Envoi = Nb_Envoi + 1
                End If
'===============================================''nombre de composants sources avec reponse aux consultations =======''========================'
                If Cells(rngcell.Row, "U") = "Reponse_OK" Then
                Nb_Rep = Nb_Rep + 1
                End If
'================================================REACH presence SVHC================================================================================
                If Cells(rngcell.Row, "O") = "Oui" Then
                Nb_REACH_Oui = Nb_REACH_Oui + 1
                ElseIf Cells(rngcell.Row, "O") = "Non" Then
                Nb_REACH_Non = Nb_REACH_Non + 1
                ElseIf Cells(rngcell.Row, "O") = "Unknown" Then
                Nb_REACH_Unknown = Nb_REACH_Unknown + 1
                End If
'==================================================Présence de substances de l’annexeXIV==================================================================
                If Cells(rngcell.Row, "P") = "Oui" Then
                Nb_AnnexeXIV_Oui = Nb_AnnexeXIV_Oui + 1
                ElseIf Cells(rngcell.Row, "P") = "Non" Then
                Nb_AnnexeXIV_Non = Nb_AnnexeXIV_Non + 1
                ElseIf Cells(rngcell.Row, "P") = "Unknown" Then
                Nb_AnnexeXIV_Unknown = Nb_AnnexeXIV_Unknown + 1
                End If
'=======================================================================================================================================================
                If (Cells(rngcell.Row, "O") = "Oui" Or Cells(rngcell.Row, "O") = "Non") And (Cells(rngcell.Row, "P") = "Oui" Or Cells(rngcell.Row, "P") = "Non") Then
                Nb_ligne_REACH_Anx = Nb_ligne_REACH_Anx + 1
                End If
            End If
        End If
    Next rngcell
'======================================'======================================'======================================'======================================
'======================================Remplissage des cases Synthese consultation================================================================
ActiveWorkbook.Sheets("NM Synthese_consultation").Cells(y, "H") = Nb_artcile
ActiveWorkbook.Sheets("NM Synthese_consultation").Cells(y, "I") = Nb_Envoi
ActiveWorkbook.Sheets("NM Synthese_consultation").Cells(y, "J") = Nb_Rep
ActiveWorkbook.Sheets("NM Synthese_consultation").Cells(y, "O") = Nb_ligne_REACH_Anx
'=================================================nombre de composants avec presence dans SVHC (>0,1%)=====================================================================
        ActiveWorkbook.Sheets("NM Synthese_consultation").Cells(y, "L") = Nb_REACH_Oui
'=================================================Remplissage de la colonne Nb de références avec info REACH (SVHC)===============================================================
ActiveWorkbook.Sheets("NM Synthese_consultation").Cells(y, "N") = Nb_REACH_Oui + Nb_REACH_Non
'=================================================Remplissage de la colonne REACH presence SVHC=================================================================================
        If Nb_REACH_Oui >= 1 Then
        ActiveWorkbook.Sheets("NM Synthese_consultation").Cells(y, "K") = "Oui"
        ElseIf Nb_REACH_Oui = 0 Then
            If Nb_REACH_Unknown >= 1 Then
            ActiveWorkbook.Sheets("NM Synthese_consultation").Cells(y, "K") = "Unknown"
            Else
            ActiveWorkbook.Sheets("NM Synthese_consultation").Cells(y, "K") = "Non"
            End If
        End If
'==============================='Remplissage de la colonne Présence de substances de l’annexe XIV=======================================================
        If Nb_AnnexeXIV_Oui >= 1 Then
        ActiveWorkbook.Sheets("NM Synthese_consultation").Cells(y, "M") = "Oui"
        ElseIf Nb_AnnexeXIV_Oui = 0 Then
            If Nb_AnnexeXIV_Unknown >= 1 Then
            ActiveWorkbook.Sheets("NM Synthese_consultation").Cells(y, "M") = "Unknown"
            Else
            ActiveWorkbook.Sheets("NM Synthese_consultation").Cells(y, "M") = "Non"
            End If
        End If
'============================'============================'============================'============================'============================
 '======================================Remplissage des cases Synthese consultation================================================================
ActiveWorkbook.Sheets("NM Synthese_consultation").Cells(y, "H") = Nb_artcile
ActiveWorkbook.Sheets("NM Synthese_consultation").Cells(y, "I") = Nb_Envoi
ActiveWorkbook.Sheets("NM Synthese_consultation").Cells(y, "J") = Nb_Rep
ActiveWorkbook.Sheets("NM Synthese_consultation").Cells(y, "O") = Nb_ligne_REACH_Anx
'============================'============================'============================'============================'============================
Next
Temps_fin = Time
ActiveWorkbook.Sheets("Etat de consultation").Range("$A$1:$Q$5960").AutoFilter Field:=4
ActiveWorkbook.Sheets("NM Synthese_consultation").Activate
ActiveWorkbook.Sheets("NM Synthese_consultation").Cells(1, "AF") = "Execution Terminée" & vbCrLf & Date & " a " & Time
ActiveWorkbook.Sheets("NM Synthese_consultation").Cells(1, "AF").Interior.ColorIndex = 9
ActiveWorkbook.Sheets("NM Synthese_consultation").Cells(1, "AG") = Temps_fin - Temps_debut
End Sub




