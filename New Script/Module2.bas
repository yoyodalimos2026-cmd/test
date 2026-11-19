Attribute VB_Name = "Module2"
Sub Conversion_xlsx_29_01_2024()

Set fso = CreateObject("Scripting.FileSystemObject")
Current_path = ThisWorkbook.Path
Nb_Lignes_Data = ActiveWorkbook.Sheets("LU_FMD").Range("A65536").End(xlUp).Row
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(28, "B") = Nb_Lignes_Data
Dossier_FMD = ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(7, "B").Value
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(6, "B") = Nb_Lignes_Data - 1
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(7, "B").Interior.ColorIndex = 2
Chemin_Dossier_FMD = ThisWorkbook.Path & "\" & Dossier_FMD
NB_FMD_trouvé = 0
FMD_trait = 0
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(21, "B") = Nb_Lignes_Data
Ligne_debut = ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(31, "B")
Ligne_Fin = ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(32, "B")

If ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(7, "B") = Empty Then
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(7, "B").Borders.Color = RGB(0, 0, 0)
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(7, "B").Value = Empty
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(7, "B").Value = "Veuilez Précisez le nom du dossier contenant les FMD"
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(7, "B").Interior.ColorIndex = 3

Else
'===='============================'Vérifie si le dossier des FMD exist============================'============================'
    If Dir(Chemin_Dossier_FMD, vbDirectory) = "" Then
    ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(7, "B").Borders.Color = RGB(0, 0, 0)
    ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(7, "B").Interior.ColorIndex = 3
    ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(7, "B").Value = Empty
    ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(7, "B").Value = "Le dossier spécifié est introuvable"
    
    Else
    
    'For x = 2 To Nb_Lignes_Data
    For x = Ligne_debut To Ligne_Fin
'============================'''============================'''============================'''============================''

        Nom_lien_FMD = ActiveWorkbook.Sheets("LU_FMD").Cells(x, "A").Value
        Nom_lien = Chr(34) & Nom_lien_FMD & Chr(34)

'============================''Test si le nom contient les mots clés FMD/FDS/COFC/ROHS'============================''============================'
        
'============================''Vérifier si le document existe'============================''============================''============================'
        ActiveWorkbook.Sheets("LU_FMD").Cells(x, "A").Interior.ColorIndex = 2
        ActiveWorkbook.Sheets("LU_FMD").Cells(x, "A").Borders.Color = RGB(0, 0, 0)

        If ActiveWorkbook.Sheets("LU_FMD").Cells(x, "A").Value <> "" Then
        
        Code_art_path_verif_FMD = Current_path & "\" & Dossier_FMD & "\" & Nom_lien_FMD
        'Debug.Print Code_art_path_verif_FMD
            If fso.FileExists(Code_art_path_verif_FMD) = False Then
            ActiveWorkbook.Sheets("LU_FMD").Cells(x, "A").Interior.ColorIndex = 3
            Else
            ActiveWorkbook.Sheets("LU_FMD").Cells(x, "A").Interior.ColorIndex = 2
            NB_FMD_trouvé = NB_FMD_trouvé + 1
'============================''Renommer les FMD contenant des espace dans leurs nom'============================''============================'
                If InStr(Nom_lien_FMD, " ") <> 0 Then
                Nom_FMD_ = Replace(ActiveWorkbook.Sheets("LU_FMD").Cells(x, "A").Value, " ", "_")
                Nom_FMD_1 = Chr(34) & Nom_FMD_ & Chr(34)
                ActiveWorkbook.Sheets("LU_FMD").Cells(x, "A").FormulaLocal = "=LIEN_HYPERTEXTE(" & Chr(34) & code_article_path & Chr(34) & ";" & Nom_FMD_1 & ")"
            
            If fso.FileExists(Current_path & "\" & Dossier_FMD & "\" & Nom_FMD_) = False Then
            Name Code_art_path_verif_FMD As Current_path & "\" & Dossier_FMD & "\" & Nom_FMD_
            Else
            Debug.Print "document exist"
            End If
            
            End If
''==========================================Convertir le fichier PDF en CSV =============================================================

Dossier_FMD = ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(7, "B").Value
Chemin_Dossier_FMD = ThisWorkbook.Path & "\" & Dossier_FMD & "\"
Chemin_xpdf_file = ThisWorkbook.Path & "\" & ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(12, "B")
Chemin_FMD = Chemin_Dossier_FMD & Nom_lien_FMD
Nom_lien_FMD_csv = Chemin_Dossier_FMD & Replace(Nom_lien_FMD, ".pdf", ".csv")
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(12, "C") = ""
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(12, "C").Interior.ColorIndex = 2
Fichier_Illisible = 0

Set fso = CreateObject("Scripting.FileSystemObject")

If ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(12, "B") <> Empty Then

    If fso.FileExists(Chemin_xpdf_file) = False Then   '<===================================== Vérifier si le fichier XPDF existe
    ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(12, "C") = "Le fichier n'existe pas"
    ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(12, "C").Interior.ColorIndex = 3
    Else
        'If fso.FileExists(Nom_lien_FMD_csv) = False Then
    
        CreateObject("Wscript.Shell").Run Chemin_xpdf_file & " " & Chemin_FMD & " " & "-raw" & " " & Nom_lien_FMD_csv, 0, True

'==========================='============='renomer la premiere feuille en Feuil1/Detect si le fichier csv est lisible'==========================
        
        'Debug.Print Nom_lien_FMD_csv
        'ActiveWorkbook.Sheets("LU_FMD").Cells(x, "B") = ""
            If fso.FileExists(Nom_lien_FMD_csv) = True Then
                Set WkBCible = Workbooks.Open(Nom_lien_FMD_csv)
                WkBCible.Worksheets(1).Name = "Feuil1"
                Fichier_Illisible = 0
                Nb_Lignes_worksht = WkBCible.Worksheets(1).Range("A65536").End(xlUp).Row
                Nbr_Colonne_worksht = WkBCible.Worksheets(1).UsedRange.Columns.Count
                
                    If InStr(WkBCible.Worksheets(1).Cells(1, "A"), "IMDS") = 0 Then
                        For j = 1 To Nb_Lignes_worksht
                            For e = 1 To Nbr_Colonne_worksht
                                If IsError(WkBCible.Worksheets(1).Cells(j, e)) Then
                                WkBCible.Worksheets(1).Cells(j, e) = "Error"
                            End If
                            Next
                        WkBCible.Worksheets(1).Cells(j, "A") = Replace(WkBCible.Worksheets(1).Cells(j, "A"), "=", "")
                        Next
                    End If

                If InStr(WkBCible.Worksheets(1).Cells(1, 1), "Please") <> 0 Or Nb_Lignes_worksht < 5 Then
                Fichier_Illisible = 1
                'ActiveWorkbook.Sheets("LU_FMD").Cells(x, "AA") = Nb_Lignes_worksht
                Else
                FMD_trait = FMD_trait + 1
                End If
                
            Application.DisplayAlerts = False
            WkBCible.SaveAs Filename:=Chemin_Dossier_FMD & Replace(Nom_lien_FMD, ".pdf", ".xlsx"), FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
            WkBCible.Save
            WkBCible.Close
            ActiveWorkbook.Sheets("LU_FMD").Cells(x, "AA") = Nb_Lignes_worksht
            Else
            ActiveWorkbook.Sheets("LU_FMD").Cells(x, "B") = "FMD Format IPC"
            'FMD_trait = FMD_trait + 1
            End If
            
            If Fichier_Illisible = 1 Then
            'Current_path = ThisWorkbook.Path
            ActiveWorkbook.Sheets("LU_FMD").Cells(x, "B") = "FMD Format IPC"
            'filePath = Current_path & "\" & Dossier_FMD & "\" & ActiveWorkbook.Sheets("LU_FMD").Cells(x, "A")
            'acrobatPath = "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe" ' <== Specify the path to Adobe Acrobat Reader or Acrobat Pro
            'acrobatPrintCommand = "/t """ & filePath & """" ' <==Specify the print command for Adobe Acrobat
            '?Application.ActivePrinter <== dans la fenetre d'excution
            'Application.ActivePrinter = "PDFCreator sur Ne00:" '<== Choisir PDF Creator comme imprimante par default
            'Shell acrobatPath & " " & acrobatPrintCommand, vbNormalFocus ' <==Execute the print command with Adobe Acrobat
            'Application.Wait Now + TimeValue("0:00:05")
            
            Else
            ActiveWorkbook.Sheets("LU_FMD").Cells(x, "B") = ""
            'ActiveWorkbook.Sheets("LU_FMD").Cells(x, "B").Interior.ColorIndex = 2
            End If
        'End If
    End If

Else
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(12, "C") = "Pas de nom de fichier XPDF"
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(12, "C").Interior.ColorIndex = 3

End If
      
'================================================================================================================================
        
        End If
            End If

        Next
    
    End If
    
End If
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(29, "B") = FMD_trait
End Sub
