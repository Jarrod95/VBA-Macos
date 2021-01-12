Sub FindAcfi()
''A short test script to find ACFI scores in one spreadsheet to be pasted into the residents' list
''and to paste a formula in J column to calculate the amount in dollars
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
End Sub
