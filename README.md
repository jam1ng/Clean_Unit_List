# Clean_Unit_List
Used to clean and edit Unit List from SAP into easy to use excel document
Sub clean_unit_list()
'for creating unit list. tcode ZHEUL. Extras > Download to File. Open that file in excel and run following macro
Columns("D:D").Select
Selection.Cut
Columns("A:A").Select
Selection.Insert Shift:=xlToRight
Columns("C:E").AutoFit
Columns("P:P").AutoFit

For x = 1758 To 2 Step -1
    If Cells(x, 1) = Cells(x - 1, 1) Then
    With Cells(x, 1).Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
       End With
                End If
        Next x
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
For y = 2 To lastrow
    If Cells(y, 1).Font.ColorIndex = 1 Then
    
    With Cells(y, 1)
'    .Borders(xlEdgeTop).LineStyle = xlContinuous
'    .Borders(xlEdgeTop).Weight = xlThin
    .Interior.Color = RGB(192, 192, 192)
    End With
    End If
Next y

If ActiveSheet.AutoFilterMode = False Then

Range("A1").AutoFilter

End If
    
End Sub
