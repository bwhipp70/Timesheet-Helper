Private Sub Worksheet_Deactivate()

    Config_WP_lastrow = Worksheets("WP #'s").Range("WP_LastRow").Value
    
' Readjust the full work package list length
    
    ActiveWorkbook.Names.Add _
        Name:="WP_List", _
        RefersTo:="='WP #''s'!$A$2:$A$" & Config_WP_lastrow
    
    If (Worksheets("Configuration").Range("WP_Dropdown").Value <> "") Then
        ActiveWorkbook.Names.Add _
            Name:="WP_LIST_Unique_Alpha", _
            RefersTo:="='WP #''s'!$J$2:INDEX('WP #''s'!$J$2:$J$" & Config_WP_lastrow & ",SUMPRODUCT(--('WP #''s'!$J$2:$J$" & Config_WP_lastrow & "<>"""")))"
            ' MsgBox "Something in the cell, changing to remove spaces"
    Else
        
            ' Set to default (empty cell) value
        ActiveWorkbook.Names.Add _
            Name:="WP_LIST_Unique_Alpha", _
            RefersTo:="='WP #''s'!$J$2:$J$" & Config_WP_lastrow

    End If
        

End Sub