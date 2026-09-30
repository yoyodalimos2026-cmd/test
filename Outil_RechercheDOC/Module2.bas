Attribute VB_Name = "Module2"
Sub Reset()

Dim ws1 As Worksheet

Set ws1 = ActiveWorkbook.Sheets("Home")
ws1.Cells(2, 8) = ""
ws1.Cells(3, 8) = ""
ws1.Cells(5, 8) = ""
ws1.Cells(6, 8) = ""
ws1.Cells(8, 8) = ""

Set fileNames = Range("A2:A" & Cells(Rows.Count, "A").End(xlUp).Row)

For Each file In fileNames
       
       file.Value = ""
       file.Interior.ColorIndex = xlColorIndexNone

Next file
   
End Sub
