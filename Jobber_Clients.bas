Sub Jobber_Clients()
Dim retVal As Long
    Sheets.Add(After:=Sheets(1)).Name = "New Clients"
    Worksheets(1).Select
    ActiveSheet.Name = "Existing Clients"
    Range("A:D,F:F,I:J,L:Y,AA:AA,AD:AJ,AM:AN,AU:AZ").EntireColumn.Delete
    ActiveSheet.Range("$A$1:$O$1017").RemoveDuplicates Columns:=Array(2, 3), _
        Header:=xlYes
    Range("E1") = "Age"
    Range("F1") = "Next of Kin"
    Range("G1") = "Next of Kin Contact"
    On Error Resume Next
    With Columns("D")
        .Replace "*PassedAway*", "=1", xlWhole
        .SpecialCells(xlFormulas).EntireRow.Delete
    End With
    With Columns("D")
        .Replace "*Cancel*", "=1", xlWhole
        .SpecialCells(xlFormulas).EntireRow.Delete
    End With
    With Columns("D")
        .Replace "*Hold*", "=1", xlWhole
        .SpecialCells(xlFormulas).EntireRow.Delete
    End With
    Worksheets(1).Range("A1:O1").Copy _
    Destination:=Worksheets(2).Range("A1")
    Application.ScreenUpdating = False
    Range("D2:D" & Range("D" & Rows.Count).End(3)(1).Row).AutoFilter 1, "*New*"
    Range("A2:O" & Range("A" & Rows.Count).End(3)(1).Row).SpecialCells(xlCellTypeVisible).Copy _
    Sheets(2).Cells(Rows.Count, "A").End(xlUp).Offset(1)
    Range("A2:O" & Range("A" & Rows.Count).End(3)(1).Row).SpecialCells(xlCellTypeVisible).Delete Shift:=xlUp
    Application.Wait (Now + TimeValue("0:00:03"))
    ActiveSheet.AutoFilterMode = False
    Sheets(1).Copy
    ActiveWorkbook.SaveAs "C:\Users\%USERPROFILE%\OneDrive\1. M2M Administration\EXPORTED FROM SOFTWARE\Maps Data\Existing_Clients_" & _
    Format(Date, "dd-mm-yyyy") & ".xlsx"
    ActiveWorkbook.Close
    Sheets(2).Copy
    ActiveWorkbook.SaveAs "C:\Users\%USERPROFILE%\OneDrive\1. M2M Administration\EXPORTED FROM SOFTWARE\Maps Data\New_Clients_" & _
    Format(Date, "dd-mm-yyyy") & ".xlsx"
    ActiveWorkbook.Close
    Application.ScreenUpdating = True
    MsgBox "New Spreadsheets created in" & vbCrLf & vbCrLf & "OneDrive\1. M2M Administration\EXPORTED FROM SOFTWARE\Maps Data"
    retVal = Shell("explorer.exe C:\Users\%USERPROFILE%\OneDrive\1. M2M Administration\EXPORTED FROM SOFTWARE\Maps Data", vbNormalFocus)
End Sub
