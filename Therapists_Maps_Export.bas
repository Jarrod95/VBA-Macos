Sub Therapist_Maps_Export()
Dim n As Long
Dim LastRow As Integer
Dim i As Integer
    Sheets.Add(After:=Sheets(1)).Name = "Expired"
    Worksheets(1).Select
    ActiveSheet.Name = "Active"
    n = 3
    Cells(Rows.Count, "A").End(xlUp).Offset(1 - n). _
    Resize(n).EntireRow.Delete
    Rows(1).Delete
    Columns("A").Delete
    Worksheets(1).Range("A1:N1").Copy _
    Destination:=Worksheets(2).Range("A1")
    Application.ScreenUpdating = False
    Range("H2:H" & Range("H" & Rows.Count).End(3)(1).Row).AutoFilter 1, "*Expired*"
    Range("A2:O" & Range("A" & Rows.Count).End(3)(1).Row).SpecialCells(xlCellTypeVisible).Copy _
    Sheets(2).Cells(Rows.Count, "A").End(xlUp).Offset(1)
    Range("A2:O" & Range("A" & Rows.Count).End(3)(1).Row).SpecialCells(xlCellTypeVisible).Delete Shift:=xlUp
    Worksheets(1).Select
    Range("H2:H" & Range("H" & Rows.Count).End(3)(1).Row).AutoFilter 1, "*Unavailable*"
    Range("A2:O" & Range("A" & Rows.Count).End(3)(1).Row).SpecialCells(xlCellTypeVisible).Copy _
    Sheets(2).Cells(Rows.Count, "A").End(xlUp).Offset(0)
    Range("A2:O" & Range("A" & Rows.Count).End(3)(1).Row).SpecialCells(xlCellTypeVisible).Delete Shift:=xlUp
    Worksheets(2).Select
    Worksheets("Expired").Select
    Range("O1") = "Name"
    LastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To LastRow
        ActiveSheet.Range("O" & i).Value = ActiveSheet.Range("A" & i).Value & " (" & ActiveSheet.Range("H" & i).Value & ")"
    Next i
    Application.Wait (Now + TimeValue("0:00:03"))
    Sheets(1).Copy
    ActiveWorkbook.SaveAs "C:\Users\info\OneDrive\1. M2M Administration\EXPORTED FROM SOFTWARE\Maps Data\Active-Therapists_" & _
    Format(Date, "dd-mm-yyyy") & ".xlsx"
    ActiveWorkbook.Close
    Sheets(2).Copy
    ActiveWorkbook.SaveAs "C:\Users\info\OneDrive\1. M2M Administration\EXPORTED FROM SOFTWARE\Maps Data\Expired-Therapists_" & _
    Format(Date, "dd-mm-yyyy") & ".xlsx"
    ActiveWorkbook.Close
    Application.ScreenUpdating = True
    MsgBox "New Spreadsheets created in" & vbCrLf & vbCrLf & "OneDrive\1. M2M Administration\EXPORTED FROM SOFTWARE\Maps Data"
End Sub