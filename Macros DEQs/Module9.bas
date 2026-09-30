Attribute VB_Name = "Module9"
Sub New_Diff_COLOR()
'
'
'
Dim i, finLin As Integer
Dim numericCheck As Boolean

comp = ActiveWorkbook.Name
finLin = Workbooks(comp).Sheets(1).Columns(5).Cells.Find("*", , , , , xlPrevious).Row
    For i = 18 To finLin
 If IsNumeric(Workbooks(comp).Sheets(1).Cells(i, 3)) = True And IsNumeric(Workbooks(comp).Sheets(1).Cells(i, 4)) = True Then
      C = Workbooks(comp).Sheets(1).Cells(i, 3).Value
      D = Workbooks(comp).Sheets(1).Cells(i, 4).Value
      If C <> "" And D <> "" Then
      
      If Workbooks(comp).Sheets(1).Cells(i, 6).Value = "<=" And C < D Or Workbooks(comp).Sheets(1).Cells(i, 6).Value = ">=" And C > D Then
                    Range("D" & i).Select
                    Selection.FormatConditions.Delete
                    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, _
                    Formula1:="=C" & i
                    Selection.FormatConditions(1).Interior.ColorIndex = 43

                    Range("C" & i).Select
                    Selection.FormatConditions.Delete
                    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, _
                    Formula1:="=D" & i
                    Selection.FormatConditions(1).Interior.ColorIndex = 43
      End If
      If Workbooks(comp).Sheets(1).Cells(i, 6).Value = ">=" And C < D Or Workbooks(comp).Sheets(1).Cells(i, 6).Value = "<=" And C > D Then
                    Range("D" & i).Select
                    Selection.FormatConditions.Delete
                    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, _
                    Formula1:="=C" & i
                    Selection.FormatConditions(1).Interior.ColorIndex = 48

                    Range("C" & i).Select
                    Selection.FormatConditions.Delete
                    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, _
                    Formula1:="=D" & i
                    Selection.FormatConditions(1).Interior.ColorIndex = 48
      End If
      End If
      
'       If Workbooks(COMP).Sheets(1).Cells(i, 6).Value = "<=" Then
'
'         If C < D Then
'
'                    Range("D" & i).Select
'                    Selection.FormatConditions.Delete
'                    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, _
'                    Formula1:="=C" & i
'                    Selection.FormatConditions(1).Interior.ColorIndex = 43
'
'                    Range("C" & i).Select
'                    Selection.FormatConditions.Delete
'                    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, _
'                    Formula1:="=D" & i
'                    Selection.FormatConditions(1).Interior.ColorIndex = 43
'       Else
'                    Range("D" & i).Select
'                    Selection.FormatConditions.Delete
'                    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, _
'                    Formula1:="=C" & i
'                    Selection.FormatConditions(1).Interior.ColorIndex = 48
'
'                    Range("C" & i).Select
'                    Selection.FormatConditions.Delete
'                    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, _
'                    Formula1:="=D" & i
'                    Selection.FormatConditions(1).Interior.ColorIndex = 48
'         End If
'       End If
'
'If Workbooks(COMP).Sheets(1).Cells(i, 6).Value = ">=" Then
'
'         If C > D Then
'
'                    Range("D" & i).Select
'                    Selection.FormatConditions.Delete
'                    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, _
'                    Formula1:="=C" & i
'                    Selection.FormatConditions(1).Interior.ColorIndex = 43
'
'                    Range("C" & i).Select
'                    Selection.FormatConditions.Delete
'                    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, _
'                    Formula1:="=D" & i
'                    Selection.FormatConditions(1).Interior.ColorIndex = 43
'       Else
'                    Range("D" & i).Select
'                    Selection.FormatConditions.Delete
'                    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, _
'                    Formula1:="=C" & i
'                    Selection.FormatConditions(1).Interior.ColorIndex = 48
'
'                    Range("C" & i).Select
'                    Selection.FormatConditions.Delete
'                    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, _
'                    Formula1:="=D" & i
'                    Selection.FormatConditions(1).Interior.ColorIndex = 48
'         End If
'       End If
  End If
  
      Next i
      Workbooks("COMP.xlsx").Save
      MsgBox ("Différence colorée")

End Sub


