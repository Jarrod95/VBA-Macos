Sub Openfile()
''VBA to open ACFI file, paste and clean data, and to calculate ACFI scores
Application.ScreenUpdating = False
Dim FileOpen As String
Dim parentWorkBook As Excel.Workbook
Dim otherWorkBook As Excel.Workbook
Dim workBookName As Variant

    Application.ScreenUpdating = False
    
    Set parentWorkBook = ActiveWorkbook
    'Opens a file selection window set to .csv files
    workBookName = Application.GetOpenFilename _
    (Title:="Please select ACFI file", _
    FileFilter:="csv files *.csv* (*.csv*),")
    
    If Not workBookName = False Then
    'Copies ACFI data to 'Payment Summary' spreadsheet
        Set otherWorkBook = Workbooks.Open(workBookName)
        parentWorkBook.Sheets("Payment Statement").Range("A1:T500").Value = otherWorkBook.Sheets(1).Range("A1:T500").Value
        parentWorkBook.Sheets("Cover").Range("E9").Value = Format(Date, "Medium Date") 'Updates cell below button to today's date
        otherWorkBook.Close False
        Set otherWorkBook = Nothing
    End If
    
    'Residents = Sheets("Residents").Range(StartRow:EndRow).Cut
    Dim Residents As Range
    Set Residents = Range(StartRow, EndRow)
    Residents.Copy (Sheets("Sheet1").Range("A1")) 'Work out how to paste back into cleaned spredsheet
    Sheets("Residents").Cells.Clear
    
    'Clean up spreadsheet
    Dim columnLabels As Variant
    columnLabels = Array("First Name", "Last Name", "Client Reference", "Building", "Room/Unit", _
    "Account Balance", "Trust Balance", "Bond Payout", "ACFI Score", "ACFI Funding $")
    
    Range(Cells(1, 1), Cells(1, 10)).Value = columnLabels
    
    With Range("A1:J1") 'Fix columns
        .Font.Bold = True
        .Columns.AutoFit
    End With
    
    If Range("A1") = "First Name" Then
        Columns("B").Cut
        Columns("A").Insert Shift:=xlToRight
    End If
    
    'Find & Calculate ACFI scores
    For Each c In Sheets("Payment Statement").Range("A1:A500")
            If c.Value = "CDP" Then
                For Each d In Sheets("Sheet1").Range("A1:A500")
                    If c.Offset(0, 2).Value = d.Value Then
                        d.Offset(0, 8) = Right(c.Offset(1, 6).Value, 3)
                    End If
                Next
                Debug.Print c.Offset(0, 2).Value & " " & c.Offset(0, 3).Value & " ACFI: " & Right(c.Offset(1, 6).Value, 3)
            End If
        Next
        lastRow = Sheets("Sheet1").Cells(Rows.Count, 1).End(xlUp).Row
        'Paste ACFI formula
        Range("J2:J" & lastRow).Formula = "=ifs(left($I2, 1)=""H"", Cover!$I$5, left($I2, 1)=""M"", Cover!$I$6, left($I2, 1)=""L"", Cover!$I$7, LEFT($I2, 1)=""Nil"", 0)+ ifs(mid($I2, 2, 1)=""H"", Cover!$I$8, mid($I2, 2,1)=""M"", Cover!$I$9, mid($I2, 2,1)=""L"", Cover!$I$10, mid($I2, 2,1)=""Nil"", 0)+ ifs(left($I2, 1)=""H"", Cover!$I$11, left($I2, 1)=""M"", Cover!$I$12, left($I2, 1)=""L"", Cover!$I$13, left($I2, 1)=""Nil"", 0)"
        
    Application.ScreenUpdating = True

End Sub
