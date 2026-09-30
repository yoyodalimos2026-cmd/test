Attribute VB_Name = "Module3"
Sub Reset()
Attribute Reset.VB_ProcData.VB_Invoke_Func = " \n14"
    Application.ScreenUpdating = False
    ActiveWorkbook.Sheets("Range").Select
    Rows("1:1").Select
    Selection.AutoFilter
    Range("A2").Select
    Rows("2:2").Select
    Range(Selection, Cells(Rows.Count, Selection.Column).End(xlDown)).Select
    Selection.ClearContents
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Application.ScreenUpdating = True
    ActiveWorkbook.Save
End Sub
