Public Function MacrosEnabled()
    MacrosEnabled = True
    'Paste formula into cell to determine if Macros are enabled
    '=IF(ISERROR(@MacrosEnabled()&NOW()),"In order to use this workbook, please enable Macros","Macros are enabled")
End Function
