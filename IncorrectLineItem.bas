Sub IncorrectLineItem()
Dim LastRow As Long
Dim cell As Range
Dim visit As Range
Dim sSheetName As String
Dim num As Integer

LastRow = Range("A" & Rows.Count).End(xlUp).Row
Range("Q1") = "Incorrect"
If Range("M1").Value = "One-off job $" Then
    Columns("M").EntireColumn.Hidden = True
Else
    Columns("M").EntireColumn.Hidden = False
End If

For Each cell In Range("I2:I" & LastRow)
    If (cell.Value Like "*Cancel*" _
          Or cell.Value Like "*Hold*" _
          Or cell.Value Like "*Follow Up*") _
          And cell.Offset(0, 5).Value > 0 Then
        cell.EntireRow.Interior.ColorIndex = 38
        cell.Offset(0, 8) = "TRUE"
    Else
        cell.EntireRow.Interior.ColorIndex = xlNone
        cell.Offset(0, 8) = "FALSE"
    End If
Next cell
num = Application.WorksheetFunction.CountIf(Range("Q2:Q" & LastRow), "TRUE")
With ActiveSheet.Sort
    .SortFields.Add Key:=Range("Q2"), Order:=xlDescending
    .SetRange Range("A1:Q" & LastRow)
    .Header = xlYes
    .Apply
End With
MsgBox num & " jobs with incorrect line items"

End Sub



