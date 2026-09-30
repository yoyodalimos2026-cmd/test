Attribute VB_Name = "Module5"
Sub MAJ_Etat_Consultation_20_11_2023()

ActiveWorkbook.Sheets("Etat de consultation").Cells(1, "AB") = ""
ActiveWorkbook.Sheets("Etat de consultation").Cells(1, "AB").Interior.ColorIndex = 2
ActiveWorkbook.Sheets("Etat de consultation").Cells(1, "AB") = "Execution En cours" & vbCrLf & "Durée estimatif <3 min"
ActiveWorkbook.Sheets("Etat de consultation").Cells(1, "AB").Interior.ColorIndex = 4

Temps_debut = Time
For x = 2 To 10262

'Affectation de variable
Article_SSdbl = ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(x, "F").Value
Article_Etat_Consul = ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "F").Value

ROHS_Consul = ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "N").Value
REACH_Consul = ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "O").Value
AnnexXIV_Consul = ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "P").Value
RISL_Consul = ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "Q").Value

Action_Envoi_Consul = ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "R").Value
Cmnt_Consul = ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "S").Value

ROHS_ssdbl = ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(x, "AA").Value
REACH_ssdbl = ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(x, "AD").Value
AnnexXIV_ssdbl = ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(x, "AG").Value
RISL_ssdbl = ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(x, "AJ").Value

Action_Envoi_ssdbl = ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(x, "BB").Value
Cmnt_ssdbl = ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(x, "BC").Value

DMD_Envoi_ssdbl = ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(x, "BJ").Value
Rep_Recu_ssdbl = ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(x, "BK").Value
Etat_Envoi_Consul = ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "T").Value
Etat_Rep_Consul = ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "U").Value
'=================================================================================================
'Initialisation des couleurs

ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "N").Interior.ColorIndex = 2
ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "N").Borders.Color = RGB(0, 0, 0)

ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "O").Interior.ColorIndex = 2
ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "O").Borders.Color = RGB(0, 0, 0)

ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "P").Interior.ColorIndex = 2
ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "P").Borders.Color = RGB(0, 0, 0)

ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "Q").Interior.ColorIndex = 2
ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "Q").Borders.Color = RGB(0, 0, 0)

ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "R").Interior.ColorIndex = 2
ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "R").Borders.Color = RGB(0, 0, 0)

ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "S").Interior.ColorIndex = 2
ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "S").Borders.Color = RGB(0, 0, 0)

ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "T").Interior.ColorIndex = 2
ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "T").Borders.Color = RGB(0, 0, 0)

ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "U").Interior.ColorIndex = 2
ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "U").Borders.Color = RGB(0, 0, 0)
'==============================================================================================
'================================================================================================
If Article_SSdbl = Article_Etat_Consul Then

'Remplissage des colonnes ROHS/REACH/AnnexXIV/Action/Commentaire
If ROHS_Consul = ROHS_ssdbl Then
ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "N") = ROHS_ssdbl
ElseIf ROHS_Consul <> ROHS_ssdbl Then
ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "N").Interior.ColorIndex = 31
ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "N") = ROHS_ssdbl
End If

If REACH_Consul = REACH_ssdbl Then
ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "O") = REACH_ssdbl
ElseIf REACH_Consul <> REACH_ssdbl Then
ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "O").Interior.ColorIndex = 31
ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "O") = REACH_ssdbl
End If

If AnnexXIV_Consul = AnnexXIV_ssdbl Then
ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "P") = AnnexXIV_ssdbl
ElseIf AnnexXIV_Consul <> AnnexXIV_ssdbl Then
ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "P").Interior.ColorIndex = 31
ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "P") = AnnexXIV_ssdbl
End If

If RISL_Consul = RISL_ssdbl Then
ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "Q") = RISL_ssdbl
ElseIf RISL_Consul <> RISL_ssdbl Then
ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "Q").Interior.ColorIndex = 31
ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "Q") = RISL_ssdbl
End If

If Action_Envoi_Consul = Action_Envoi_ssdbl Then
ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "R") = Action_Envoi_ssdbl
ElseIf Action_Envoi_Consul <> Action_Envoi_ssdbl Then
ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "R").Interior.ColorIndex = 31
ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "R") = Action_Envoi_ssdbl
End If

If Cmnt_Consul = Cmnt_ssdbl Then
ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "S") = Cmnt_ssdbl
ElseIf Cmnt_Consul <> Cmnt_ssdbl Then
ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "S").Interior.ColorIndex = 31
ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "S") = Cmnt_ssdbl
End If
'=====================================================================================================
'Remplissage des colonnes etat d'envoi
DMD_Envoi_ssdbl = ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(x, "BJ").Value
Rep_Recu_ssdbl = ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(x, "BK").Value
Etat_Envoi_Consul = ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "T").Value
Etat_Rep_Consul = ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "U").Value

If DMD_Envoi_ssdbl <> Etat_Envoi_Consul Then
ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "T").Interior.ColorIndex = 31
End If
    
If InStr(1, DMD_Envoi_ssdbl, "Dem") <> 0 Then
ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "T") = "Envoi_OK"
ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "T").Interior.ColorIndex = 10
Else
ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "T") = DMD_Envoi_ssdbl
End If
'========================================================================================================
'Remplissage des colonnes etat de reponse
If Rep_Recu_ssdbl <> Etat_Rep_Consul Then
ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "U").Interior.ColorIndex = 31
End If
    
If InStr(1, Rep_Recu_ssdbl, "Rép") <> 0 Then
ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "U") = "Reponse_OK"
ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "U").Interior.ColorIndex = 10
Else
ActiveWorkbook.Sheets("Etat de consultation").Cells(x, "U") = Rep_Recu_ssdbl
End If
'========================================================================================================
End If
Next
Temps_fin = Time
ActiveWorkbook.Sheets("Etat de consultation").Cells(1, "AB") = "Execution Terminée" & vbCrLf & Date & " a " & Time
ActiveWorkbook.Sheets("Etat de consultation").Cells(1, "AB").Interior.ColorIndex = 9
'ActiveSheet.Range("$A$1:$Q$5960").AutoFilter Field:=4
ActiveWorkbook.Sheets("Etat de consultation").Cells(1, "AC") = Temps_fin - Temps_debut

End Sub




