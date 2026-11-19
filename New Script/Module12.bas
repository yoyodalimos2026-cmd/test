Attribute VB_Name = "Module12"
Sub Verif_Recuperation_FMD_05_03_2024()

'===================================Elimination des doublons==========================================================================
  
    'ActiveWorkbook.Sheets("BOM_a_analyser").Activate
    'Columns("D:D").Select
    'ActiveWorkbook.Sheets("BOM_a_analyser").Range("$A$1:$P$1038475").RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16), Header:=xlYes
         
'================================================================================================================================================

Nb_Lgn_BOM_a_analyser = ActiveWorkbook.Sheets("BOM_a_analyser").Range("D65536").End(xlUp).Row
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(28, "I") = Nb_Lgn_BOM_a_analyser
Dossier_FMD = ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(7, "B").Value
Chemin_Dossier_FMD = ThisWorkbook.Path & "\" & Dossier_FMD & "\"
Nb_Lgn_Contact = ActiveWorkbook.Sheets("Contact").Range("A65536").End(xlUp).Row

b = Nb_Lgn_Contact + 1
FMD_exp = 0
ActiveWorkbook.Sheets("BOM_a_analyser").Cells(1, "A") = "Code Article"
ActiveWorkbook.Sheets("BOM_a_analyser").Cells(1, "A").Interior.ColorIndex = 43
ActiveWorkbook.Sheets("BOM_a_analyser").Cells(1, "B") = "Référence"
ActiveWorkbook.Sheets("BOM_a_analyser").Cells(1, "B").Interior.ColorIndex = 43
ActiveWorkbook.Sheets("BOM_a_analyser").Cells(1, "C") = "Fabricant"
ActiveWorkbook.Sheets("BOM_a_analyser").Cells(1, "C").Interior.ColorIndex = 43
ActiveWorkbook.Sheets("BOM_a_analyser").Cells(1, "D") = "FMD"
ActiveWorkbook.Sheets("BOM_a_analyser").Cells(1, "D").Interior.ColorIndex = 43
ActiveWorkbook.Sheets("BOM_a_analyser").Cells(1, "I") = "Sustance REACH"
ActiveWorkbook.Sheets("BOM_a_analyser").Cells(1, "A").Interior.ColorIndex = 43
ActiveWorkbook.Sheets("BOM_a_analyser").Cells(1, "J") = "Annex14"
ActiveWorkbook.Sheets("BOM_a_analyser").Cells(1, "K") = "Annex XVII"
ActiveWorkbook.Sheets("BOM_a_analyser").Cells(1, "L") = "RISL PAR"
ActiveWorkbook.Sheets("BOM_a_analyser").Cells(1, "M") = "RISL DFA"
ActiveWorkbook.Sheets("BOM_a_analyser").Cells(1, "N") = "DFA>0,1%"
ActiveWorkbook.Sheets("BOM_a_analyser").Cells(1, "O") = "POP"
ActiveWorkbook.Sheets("BOM_a_analyser").Cells(1, "P") = "PFAS"
ActiveWorkbook.Sheets("BOM_a_analyser").Cells(1, "Q") = "AD-DSL"
'====================================Dtection de l'onglet de la liste de susbstances=============================================================
For u = 2 To Sheets.Count
Onglet_Matrice_Liste_subs = ActiveWorkbook.Sheets(u - 1).Name
Next

Ligne_debut = ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(31, "I")
Ligne_Fin = ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(32, "I")
'=================================================================================================================================================
'For i = 2 To Nb_Lgn_BOM_a_analyser
For i = Ligne_debut To Ligne_Fin

    For K = 8 To 16
    ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, K).Interior.ColorIndex = 2
    Next K
    
    For h = 5 To 7
    ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, h).Interior.ColorIndex = 2
    ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, h) = ""
    Next h
    
    ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "D") = Replace(ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "D"), " ", "_")
    ActiveWorkbook.Sheets("BOM_a_analyser").Range("A" & Ligne_debut & ":O" & Ligne_Fin).Borders.LineStyle = xlContinuous
    ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "E").FormulaLocal = "=RECHERCHEV(D" & i & ";" & "LU_FMD" & "!A:AA;2;FAUX)"
    
    If Application.IsNA(ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "E")) Then
    ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "E") = "FMD(Pas d'analyse préliminaire faite)"
    ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "AA") = 0
    End If
    
    If ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "E") = 0 Then
    ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "E") = "FMD(Pas d'analyse préliminaire faite)"
    ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "AA") = 0
    End If
    
'================='==============================================================================================================================
If InStr(ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "E"), "FMD introuvable sous le dossier FMD") <> 0 _
Or InStr(ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "E"), "FMD Format IPC") <> 0 _
Or InStr(ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "E"), "Pas de FMD associé") <> 0 _
Or InStr(ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "E"), "FMD(Pas d'analyse préliminaire faite)") _
Or InStr(ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "E"), "Pas de susbtance interdite") <> 0 Then
ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "E").Interior.ColorIndex = 3
'Debug.Print "FMD Illisible"
ActiveWorkbook.Sheets("FMD_Récupéré").Rows(1 & ":" & 20000).Delete
ActiveWorkbook.Sheets("FMD_CSV").Rows(1 & ":" & 20000).Delete

Else
    ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "D") = Replace(ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "D"), " ", "_")
    ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "AA").FormulaLocal = "=RECHERCHEV(D" & i & ";" & "LU_FMD" & "!A:AA;27;FAUX)"
    Last_Row = ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "AA") + 1
    Nb_ligne_CSV = ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "AA") + 1
    ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "E") = "FMD exploitable"
    ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "F") = Date
    ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "G") = Onglet_Matrice_Liste_subs
    FMD_exp = FMD_exp + 1
    ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(29, "I") = FMD_exp + 1

'========================================='Initialisation de l'onglet FMD_CSV et FMD_Récupéré'========================================='=========================================
    'Debug.Print "FMD Format lisible"
    Nb_lignes_FMD_recup = ActiveWorkbook.Sheets("FMD_Récupéré").Range("A65536").End(xlUp).Row
    If Nb_lignes_FMD_recup > 1 Then
    ActiveWorkbook.Sheets("FMD_Récupéré").Rows(1 & ":" & ActiveWorkbook.Sheets("FMD_Récupéré").Rows.Count).Delete
    End If
    
    Nb_lignes_FMD_csv = ActiveWorkbook.Sheets("FMD_CSV").Range("A65536").End(xlUp).Row
    
    If Nb_lignes_FMD_csv > 1 Then
    ActiveWorkbook.Sheets("FMD_CSV").Rows(1 & ":" & ActiveWorkbook.Sheets("FMD_CSV").Rows.Count).Delete
    'ActiveWorkbook.Sheets("FMD_Récupéré").Columns("A:Z").NumberFormat = "@"
    End If
'=========================================''=========================================''=========================================''========================================='

    ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "D") = Replace(ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "D"), " ", "_")
    '=========================================Récupération des données de ma FMD dans la feuille "FMD CSV"==================
        FMD_Name = ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "D").Value
        FMD_Name_xlsx = Replace(FMD_Name, ".pdf", ".xlsx")
        FMD_csv_Sheet_name = Replace(FMD_Name_csv, ".csv", "")
                        
        ThisWorkbook.Names.Add "plage", RefersTo:="='" & Chemin_Dossier_FMD & "[" & FMD_Name_xlsx & "]Feuil1'!$A$1:$N$" & Nb_ligne_CSV
            'De la feuille 1 du fichier source
            With Worksheets("FMD_CSV")
            'Colle les données du fichier ouvert sur la feuille 6, sur la plage A1>F10 à partir des données du fichier source copiées dans "=laplage"
            Worksheets("FMD_CSV").Range("A1:N" & Nb_ligne_CSV).Value = "=plage"
            End With

    '=================================================================================================================================
    Nb_ligne_CSV = ActiveWorkbook.Sheets("FMD_CSV").Range("A65536").End(xlUp).Row
    Nbr_Colonne_CSV = ActiveWorkbook.Sheets("FMD_CSV").UsedRange.Columns.Count
    '=================================Mettre toutes les informations utiles dans la premiére colonne======================================
    
    'For y = 1 To nb_ligne_csv
    For y = 1 To Nb_ligne_CSV
    
        For zz = 2 To Nbr_Colonne_CSV
        '======================================Ignorer #NOM?======================================================
        If IsError(ActiveWorkbook.Sheets("FMD_CSV").Cells(y, zz).Value) Then
        ActiveWorkbook.Sheets("FMD_CSV").Cells(y, zz) = "ERROR"
        End If
         
        ActiveWorkbook.Sheets("FMD_CSV").Cells(y, "A") = Replace(ActiveWorkbook.Sheets("FMD_CSV").Cells(y, "A"), Chr(13) & Chr(10), "")
        
            If ActiveWorkbook.Sheets("FMD_CSV").Cells(y, zz).Value <> 0 Then
            ActiveWorkbook.Sheets("FMD_CSV").Cells(Last_Row, 1).Value = ActiveWorkbook.Sheets("FMD_CSV").Cells(y, zz).Value
            Last_Row = Last_Row + 1
            End If
            
            If InStr(ActiveWorkbook.Sheets("FMD_CSV").Cells(y, zz), "=") <> 0 Then
            ActiveWorkbook.Sheets("FMD_CSV").Cells(y, zz) = Replace(ActiveWorkbook.Sheets("FMD_CSV").Cells(y, zz), "- ", "/")
            End If
            
        Next
    Next
    
    '=================================================================================================================================
    '=============================================Copiage au format Text====================================================================================
    ActiveWorkbook.Sheets("FMD_CSV").Activate
    Application.DisplayAlerts = False
        Columns("A:A").Select
        Range("A126").Activate
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False

    '=============================================Application du Séparateur espace===================================
        Application.CutCopyMode = False
        Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
            :=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), _
            Array(7, 1)), TrailingMinusNumbers:=True
            'Call Macro3
     '=============================================Remplacement des caratéres +,- et = ===================================
    'For f = 1 To nb_ligne_csv * 1.5
    
    Nbr_Colonne_CSV = ActiveWorkbook.Sheets("FMD_CSV").UsedRange.Columns.Count
    'Debug.Print Nbr_Colonne_CSV
    For f = 1 To Nb_ligne_CSV * 1.5
        For g = 1 To Nbr_Colonne_CSV
            If IsError(ActiveWorkbook.Sheets("FMD_CSV").Cells(f, g)) Then
            ActiveWorkbook.Sheets("FMD_CSV").Cells(f, g) = "Error"
            End If
            
            If InStr(ActiveWorkbook.Sheets("FMD_CSV").Cells(f, g), "- ") <> 0 Then
            ActiveWorkbook.Sheets("FMD_CSV").Cells(f, g) = "/"
            End If
'===========================================================================================================
                      
        Next
    Next
'============================================='Detection des CAS Number'=============================================
    v = 2
    
    Set regexObject_CAS = CreateObject("vbscript.regexp") '<=== Utulisation de la fonction expression réguliére
    'regexObject_CAS.Pattern = "[0-9]{1,5}(-)[0-9]{1,4}(-)[0-9]{1,4}"
    regexObject_CAS.Pattern = "[0-9]{1,9}(-)[0-9]{1,9}(-)[0-9]{1,9}"
    Set regexObject_CAS_1 = CreateObject("vbscript.regexp") '<=== Utulisation de la fonction expression réguliére
    'regexObject_CAS_1.Pattern = "[0-9](/)[0-9]{1,3}(/)"
    regexObject_CAS_1.Pattern = "[0-9]{1,5}(/)[0-9]{1,4}(/)[0-9]{1,4}"
    Set regexObject_date = CreateObject("vbscript.regexp")
    regexObject_date.Pattern = "[0-9]{1,5}(/)[0-9]{1,4}(/)[0-9]{1,4}"
    
    Set regexObject_mail = CreateObject("vbscript.regexp") '<=== Utulisation de la fonction expression réguliére
    'regexObject_mail.Pattern = "[a-z]" & "@" & "[a-z]" & "\." & "[a-z]"
    regexObject_mail.Pattern = "^[a-z0-9._-]+@[a-z0-9._-]{2,}\.[a-z]{2,4}$"
    
    Nb_Lgn_FMD_csv = ActiveWorkbook.Sheets("FMD_CSV").Range("A65536").End(xlUp).Row
    ActiveWorkbook.Sheets("FMD_Récupéré").Cells(1, "E") = FMD_Name_xlsx
    code_art_ref = 0
    Type_Rapport = "Own"
    Accept_Date = "NA"
    
    'For t = 2 To Nb_Lgn_FMD_csv + 10
    For t = 2 To Nb_ligne_CSV * 1.5
        For r = 1 To Nbr_Colonne_CSV
        
        '===========================================================================================================
        
        '=========================================='Detection des CAS Number'==========================================
            If regexObject_CAS.Test(ActiveWorkbook.Sheets("FMD_CSV").Cells(t, r)) = True Or regexObject_CAS_1.Test(ActiveWorkbook.Sheets("FMD_CSV").Cells(t, r)) = True Then
            ActiveWorkbook.Sheets("FMD_Récupéré").Cells(v, 1).NumberFormat = "@"
            ActiveWorkbook.Sheets("FMD_Récupéré").Cells(v, 1) = Replace(ActiveWorkbook.Sheets("FMD_CSV").Cells(t, r), "/", "-")
            'ActiveWorkbook.Sheets("FMD_Récupéré").Cells(v, 1) = ActiveWorkbook.Sheets("FMD_CSV").Cells(t, r)
            ActiveWorkbook.Sheets("FMD_Récupéré").Cells(v, "B").FormulaLocal = "=RECHERCHEV(A" & v & ";'" & CStr(Onglet_Matrice_Liste_subs) & "'!A:D;1;FAUX)"
            'Debug.Print "=RECHERCHEV(A" & v & ";'" & CStr(Onglet_Matrice_Liste_subs) & "'!A:D;1;FAUX)"
            ActiveWorkbook.Sheets("FMD_Récupéré").Cells(v, "C").FormulaLocal = "=RECHERCHEV(A" & v & ";'" & CStr(Onglet_Matrice_Liste_subs) & "'!A:J;3;FAUX)"
            ActiveWorkbook.Sheets("FMD_Récupéré").Cells(v, "D").FormulaLocal = "=RECHERCHEV(A" & v & ";'" & CStr(Onglet_Matrice_Liste_subs) & "'!A:D;2;FAUX)"
                
                If Application.IsNA(ActiveWorkbook.Sheets("FMD_Récupéré").Cells(v, "B")) Then
                ActiveWorkbook.Sheets("FMD_Récupéré").Cells(v, "B") = "0"
                ActiveWorkbook.Sheets("FMD_Récupéré").Cells(v, "C") = "0"
                ActiveWorkbook.Sheets("FMD_Récupéré").Cells(v, "D") = "0"
                End If
                
            v = v + 1
            End If
            '==========================================Vérif si le code article ou ref existant dans la FMD==========================
            If InStr(ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "D"), "MDS") <> 0 Then
                If ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "B") = ActiveWorkbook.Sheets("FMD_CSV").Cells(t, r) Then
                code_art_ref = 1
                ActiveWorkbook.Sheets("FMD_Récupéré").Cells(1, "I") = ActiveWorkbook.Sheets("FMD_CSV").Cells(t, r)
                'ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "B") = ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "B")
                ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "B").Interior.ColorIndex = 4
                ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "E") = "FMD correspond a l'article/Ref"
                End If
                
                '=======================================Detection du IMDS/version============================='======================================='============================
                If ActiveWorkbook.Sheets("FMD_CSV").Cells(t, r) = "Version:" Then
                ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "I") = ActiveWorkbook.Sheets("FMD_CSV").Cells(t, r + 1) & "/" & ActiveWorkbook.Sheets("FMD_CSV").Cells(t, r + 3)
                End If
                '======================================='=====Detection own ou received'======================================='=======================================
                If ActiveWorkbook.Sheets("FMD_CSV").Cells(t, r) = "acceptée" Or ActiveWorkbook.Sheets("FMD_CSV").Cells(t, r) = "accepted" Then '
                Type_Rapport = "Received"
                Accept_Date = ActiveWorkbook.Sheets("FMD_CSV").Cells(t, r + 1)
                End If
                
                If InStr(ActiveWorkbook.Sheets("FMD_CSV").Cells(t, r), "[106266]") <> 0 Then ' <==Code ACTIA_Nordic
                Compte_IMDS = "ACTIA Nordic [106266]"
                ElseIf InStr(ActiveWorkbook.Sheets("FMD_CSV").Cells(t, r), "[30862]") <> 0 Then ' <==Code ACTIA_AAUT
                Compte_IMDS = "/ACTIA SA[30862]"
                End If
                '======================================='======================================='======================================='=======================================
                
            ElseIf InStr(ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "D"), "MDS") = 0 Then
            
                If InStr(ActiveWorkbook.Sheets("FMD_CSV").Cells(t, r), Mid(ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "B"), 1, 4)) <> 0 Then
                code_art_ref = 1
                ActiveWorkbook.Sheets("FMD_Récupéré").Cells(1, "I") = ActiveWorkbook.Sheets("FMD_CSV").Cells(t, r)
                ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "B").Interior.ColorIndex = 4
                ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "E") = "FMD correspond a l'article/Ref"
                'ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "B") = ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "B")
                End If
                
                
            End If
            '=========================================='========Detection des contacts================='==========================================
            If regexObject_mail.Test(ActiveWorkbook.Sheets("FMD_CSV").Cells(t, r)) = True Then
            'ActiveWorkbook.Sheets("FMD_Récupéré").Cells(1, 10) = ActiveWorkbook.Sheets("FMD_CSV").Cells(t, r)
            ActiveWorkbook.Sheets("Contact").Cells(b, 1) = ActiveWorkbook.Sheets("FMD_CSV").Cells(t, r)
            ActiveWorkbook.Sheets("Contact").Cells(b, 3) = FMD_Name
            'Debug.Print "=RECHERCHEV(C" & b & ";" & "BDD_FMD" & "!I:L;4;FAUX)"
            ActiveWorkbook.Sheets("Contact").Cells(b, 2).FormulaLocal = "=RECHERCHEV(C" & b & ";'" & "BDD_FMD" & "'!I:L;4;FAUX)"
            b = b + 1
            End If
            
            '=================================================================================================================================
        Next
        
        '=====================================Remplissage colone J ===========================================================================
        If InStr(ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "D"), "MDS") <> 0 Then
        ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "J") = Type_Rapport & "/" & Compte_IMDS
        ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "K") = Accept_Date
        End If
        '====================================='====================================='====================================='====================================
        
    Next
'=============================================''Lister les CAS number sur l'onglet FMD Récupéré''============================================='
    'ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "B") = ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "B") & "/"
    Nb_lignes_FMD_recup = ActiveWorkbook.Sheets("FMD_Récupéré").Range("A65536").End(xlUp).Row
    Nb_Substance = 0
    
    For f = 2 To Nb_lignes_FMD_recup
        If ActiveWorkbook.Sheets("FMD_Récupéré").Cells(f, "C") <> 0 Then
        
        CAS_Nbr = ActiveWorkbook.Sheets("FMD_Récupéré").Cells(f, "B")
        
            If InStr(ActiveWorkbook.Sheets("FMD_Récupéré").Cells(f, "C"), "RoHS") <> 0 Then
                If InStr(ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "H"), CAS_Nbr) = 0 Then
                ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "H") = ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "H") & "//" & CAS_Nbr
                ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "H").Interior.ColorIndex = 20
                Else
                ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "H").Interior.ColorIndex = 43
                End If
            End If
        
            If InStr(ActiveWorkbook.Sheets("FMD_Récupéré").Cells(f, "C"), "SVHC") <> 0 Then
                If InStr(ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "I"), CAS_Nbr) = 0 Then
                ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "I") = ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "I") & "//" & CAS_Nbr
                ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "I").Interior.ColorIndex = 20
                Else
                ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "I").Interior.ColorIndex = 43
                End If
            
            End If
             
            If InStr(ActiveWorkbook.Sheets("FMD_Récupéré").Cells(f, "C"), "Annex XIV") <> 0 Then
                If InStr(ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "J"), CAS_Nbr) = 0 Then
                Debug.Print ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "J")
                ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "J") = ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "J") & "//" & CAS_Nbr
                ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "J").Interior.ColorIndex = 20
                Else
                ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "J").Interior.ColorIndex = 43
                End If
            End If
             
            If InStr(ActiveWorkbook.Sheets("FMD_Récupéré").Cells(f, "C"), "Annex XVII") <> 0 Then
                If InStr(ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "K"), CAS_Nbr) = 0 Then
                ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "K") = ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "K") & "//" & CAS_Nbr
                ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "K").Interior.ColorIndex = 20
                Else
                ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "K").Interior.ColorIndex = 43
                End If
            End If
              
            If InStr(ActiveWorkbook.Sheets("FMD_Récupéré").Cells(f, "C"), "POP") <> 0 Then
                If InStr(ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "O"), CAS_Nbr) = 0 Then
                ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "O") = ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "O") & "//" & CAS_Nbr
                ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "O").Interior.ColorIndex = 20
                Else
                ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "O").Interior.ColorIndex = 43
                End If
            End If
            
            If InStr(ActiveWorkbook.Sheets("FMD_Récupéré").Cells(f, "C"), "PAR") <> 0 Then
                If InStr(ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "L"), CAS_Nbr) = 0 Then
                ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "L") = ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "L") & "//" & CAS_Nbr
                ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "L").Interior.ColorIndex = 20
                Else
                ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "L").Interior.ColorIndex = 43
                End If
            End If
            
            If InStr(ActiveWorkbook.Sheets("FMD_Récupéré").Cells(f, "C"), "DFA") <> 0 Then
                If InStr(ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "M"), CAS_Nbr) = 0 Then
                ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "M") = ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "M") & "//" & CAS_Nbr
                ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "N") = ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "N") & "//" & CAS_Nbr
                ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "M").Interior.ColorIndex = 20
                Else
                ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "M").Interior.ColorIndex = 43
                ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "N").Interior.ColorIndex = 43
                End If
            End If
            
            If InStr(ActiveWorkbook.Sheets("FMD_Récupéré").Cells(f, "C"), "PFAS") <> 0 Then
                If InStr(ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "P"), CAS_Nbr) = 0 Then
                ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "P") = ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "P") & "//" & CAS_Nbr
                ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "P").Interior.ColorIndex = 20
                Else
                ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "P").Interior.ColorIndex = 43
                End If
            End If
            
            If InStr(ActiveWorkbook.Sheets("FMD_Récupéré").Cells(f, "C"), "POP") <> 0 Then
                If InStr(ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "O"), CAS_Nbr) = 0 Then
                ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "O") = ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "O") & "//" & CAS_Nbr
                ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "O").Interior.ColorIndex = 20
                Else
                ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "O").Interior.ColorIndex = 43
                End If
            End If
            
            
            If InStr(ActiveWorkbook.Sheets("FMD_Récupéré").Cells(f, "C"), "AD-DSL") <> 0 Then
                If InStr(ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "Q"), CAS_Nbr) = 0 Then
                ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "Q") = ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "Q") & "//" & CAS_Nbr
                ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "Q").Interior.ColorIndex = 20
                Else
                ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "Q").Interior.ColorIndex = 43
                End If
            End If

        End If
    

        
        'If ActiveWorkbook.Sheets("FMD_Récupéré").Cells(f, "B") <> 0 Then
        'CAS_nbr = "[" & ActiveWorkbook.Sheets("FMD_Récupéré").Cells(f, "B") & "]"
        'Liste = "[" & ActiveWorkbook.Sheets("FMD_Récupéré").Cells(f, "C") & "]"
        'Substance = ActiveWorkbook.Sheets("FMD_Récupéré").Cells(f, "D")
        'Nb_Substance = Nb_Substance + 1
            'If InStr(ActiveWorkbook.Sheets("FMD_Récupéré").Cells(1, "F"), CAS_nbr) = 0 Then
            'ActiveWorkbook.Sheets("FMD_Récupéré").Cells(1, "F") = CAS_nbr & "//" & ActiveWorkbook.Sheets("FMD_Récupéré").Cells(1, "F")
            'ActiveWorkbook.Sheets("FMD_Récupéré").Cells(1, "G") = Liste & "//" & ActiveWorkbook.Sheets("FMD_Récupéré").Cells(1, "G")
            'ActiveWorkbook.Sheets("FMD_Récupéré").Cells(1, "G") = Replace(ActiveWorkbook.Sheets("FMD_Récupéré").Cells(1, "G"), "]]", "]")
            'ActiveWorkbook.Sheets("FMD_Récupéré").Cells(1, "G") = Replace(ActiveWorkbook.Sheets("FMD_Récupéré").Cells(1, "G"), "[[", "[")
            'ActiveWorkbook.Sheets("FMD_Récupéré").Cells(1, "H") = Substance & "//" & ActiveWorkbook.Sheets("FMD_Récupéré").Cells(1, "H")
            'End If
        'End If
   
    Next
    
'=============================================''Remplissage Onglet BOM_a_analyser''============================================='
    
    'If Nb_Substance > 0 Then
    'ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "D") = ActiveWorkbook.Sheets("FMD_Récupéré").Cells(1, "F")
    'ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "E") = ActiveWorkbook.Sheets("FMD_Récupéré").Cells(1, "G")
    'ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "F") = ActiveWorkbook.Sheets("FMD_Récupéré").Cells(1, "H")
    'ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "G") = Date
    'ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "H") = Onglet_Matrice_Liste_subs
    'Else
    'ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "D") = "Pas de susbtance interdite"
    'ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "D").Interior.ColorIndex = 3
    'ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "E") = "Pas de susbtance interdite"
    'ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "E").Interior.ColorIndex = 3
    'ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "F") = "Pas de susbtance interdite"
    'ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "F").Interior.ColorIndex = 3
    'ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "G") = Date
    'ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "G").Interior.ColorIndex = 3
    'ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "H") = Onglet_Matrice_Liste_subs
    'ActiveWorkbook.Sheets("BOM_a_analyser").Cells(i, "H").Interior.ColorIndex = 3
    'End If
'============================================='============================================='=============================================
End If
Next


ActiveWorkbook.Sheets("BOM_a_analyser").Activate
Columns("D:D").Select
ActiveWorkbook.Sheets("BOM_a_analyser").Range("$A$1:$P$1038475").RemoveDuplicates Columns:=Array(1, 2, 3, 4, _
        5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16), Header:=xlYes


End Sub




