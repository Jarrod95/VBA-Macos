Sub IncorrectLineItem()
Dim LastRow As Long
Dim cell As Range
Dim visit As Range
Dim sSheetName As String
Dim num As Integer
Application.ScreenUpdating = False
LastRow = Range("A" & Rows.Count).End(xlUp).Row

Rows("1:1").Select
Selection.Font.Bold = True
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True

Range("E:F,H:H,M:M,O:P").EntireColumn.Hidden = True
Range("Q1") = "Incorrect"
With Columns("N:N")
    .Style = "Currency"
    .ColumnWidth = 8.5
End With
For Each cell In Range("I2:I" & LastRow)
    If (cell.Value Like "*Cancel*" _
          Or cell.Value Like "*Hold*" _
          Or cell.Value Like "*Follow Up*") _
          And cell.Offset(0, 5).Value > 0 Then
        cell.EntireRow.Interior.Color = RGB(255, 113, 113)
        cell.Offset(0, 8) = "TRUE"
    ElseIf Not (cell.Value Like "*Cancel*" _
          Or cell.Value Like "*Hold*" _
          Or cell.Value Like "*Follow Up*") _
          And cell.Offset(0, 5).Value = 0 Then
          cell.EntireRow.Interior.Color = RGB(255, 153, 83)
          cell.Offset(0, 8) = "TRUE"
    Else
        cell.EntireRow.Interior.ColorIndex = xlNone
        cell.Offset(0, 8) = "FALSE"
    End If
Next cell
num = Application.WorksheetFunction.CountIf(Range("Q2:Q" & LastRow), "TRUE")
Range("B:C,F:H").EntireColumn.AutoFit
Columns("A").ColumnWidth = 10
Columns("D").ColumnWidth = 30
Columns("I").ColumnWidth = 35
Columns("K").ColumnWidth = 32

With ActiveSheet.Sort
    .SortFields.Clear
    .SortFields.Add Key:=Range("Q2"), Order:=xlDescending
    .SortFields.Add Key:=Range("G2"), Order:=xlDescending
    .SortFields.Add Key:=Range("D2"), Order:=xlAscending
    .SetRange Range("A1:Q" & LastRow)
    .Header = xlYes
    .Apply
End With
Application.ScreenUpdating = True
MsgBox num & " jobs with incorrect line items"

End Sub
