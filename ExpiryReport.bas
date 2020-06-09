Sub ExpiryReport()
Dim n As Integer
Dim LastRow As Integer
    Worksheets(1).Select
    ActiveSheet.Name = "Expiring Next Month"
    n = 3
    Cells(Rows.Count, "A").End(xlUp).Offset(1 - n). _
    Resize(n).EntireRow.Delete
    Rows(1).Delete
    Columns("A").Delete
    Columns("A:H").ColumnWidth = 18
    LastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    Range("E2:H" & Range("H" & Rows.Count).End(3)(1).Row).Select
    Range("E2:H9").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=TODAY()"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
        Formula1:="=TODAY()", Formula2:="=EOMONTH(TODAY(), 1)"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16754788
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 10284031
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
End Sub
