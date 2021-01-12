Sub CleanSheet()
'Macro to clean and organise St Andrews Client List from Inerva
Dim FileOpen As String
Dim Residents As Range
Dim parentWorkBook As Excel.Workbook
Dim otherWorkBook As Excel.Workbook
Dim workBookName As Variant
Dim StartRow As Range
Dim EndRow As Range

    'Open Workbook
    workBookName = Application.GetOpenFilename _
    (Title:="Please select Residents file", _
    FileFilter:="xls files *.xls* (*.xls*),")

    Set parentWorkBook = ActiveWorkbook

    If Not workBookName = False Then
    'Copies ACFI data to 'Payment Summary' spreadsheet
        Set otherWorkBook = Workbooks.Open(workBookName)
        parentWorkBook.Sheets("Residents").Range("A1:T500").Value = otherWorkBook.Sheets(1).Range("A1:T500").Value
        parentWorkBook.Sheets("Cover").Range("E16").Value = Format(Date, "Medium Date") 'Updates cell below button to today's date
        otherWorkBook.Close False
        Set otherWorkBook = Nothing
    End If

    Application.Wait (Now + TimeValue("0:00:05"))
    'Determine where residents are in sheet
    For Each c In Sheets("Residents").Range("B1:B500") '[B1:B500]
        If c.Value Like "Client Type : Resident" Then
            StartRow = "A" & c.Row
            Debug.Print StartRow
        End If

        If c.Value Like "Client Type : Staff" Then
            EndRow = "T" & c.Row - 1
            Debug.Print EndRow
        End If
    Next

    Set Residents = Sheets("Residents").Range(StartRow, EndRow)
    'Residents.Copy (Sheets("Sheet1").Range("A1")) 'Work out how to paste back into cleaned spredsheet
    Sheets("Residents").Cells.Clear
                
End Sub
