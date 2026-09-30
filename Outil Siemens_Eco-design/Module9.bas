Attribute VB_Name = "Module9"
Sub macro_transfert_data_ref_supplementaire()

Nb_Lignes_Total = ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Range("A65536").End(xlUp).Row

For t = 9797 To Nb_Lignes_Total
If ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(t, 24) <> "0 A ne pas traité" Then
    If ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(t, 27) = 0 Then
    'For t = 9029 To 9030
    d = 10
        For r = 24 To 64
        'Debug.Print "RECHERCHEV(O" & t & ";O:BL;r;FAUX)"
        ActiveWorkbook.Sheets("LAF ss dbl_NEXTEO_KCVP_ss_db").Cells(t, r).FormulaLocal = "=RECHERCHEV(O" & t & ";O:BL;" & d & ";FAUX)"
        d = d + 1
        Next
    End If
End If
Next

End Sub

    
