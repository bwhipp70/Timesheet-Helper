Private Sub Worksheet_Change(ByVal Target As Range)
    Dim KeyCells As Range

    ' The variable KeyCells contains the cells that will
    ' cause an alert when they are changed.
    Set KeyCells = Range("WP_Dropdown")
    
    Config_WP_lastrow = Worksheets("WP #'s").Range("WP_LastRow").Value
    
    If Not Application.Intersect(KeyCells, Range(Target.Address)) _
           Is Nothing Then

        If (Range("WP_Dropdown").Value <> "") Then
             ActiveWorkbook.Names.Add _
                Name:="WP_LIST_Unique_Alpha", _
                RefersTo:="='WP #''s'!$I$2:INDEX('WP #''s'!$I$2:$I$" & Config_WP_lastrow & ",SUMPRODUCT(--('WP #''s'!$I$2:$I$" & Config_WP_lastrow & "<>"""")))"
            ' MsgBox "Something in the cell, changing to remove spaces"
        Else
        
           ' Set to default (empty cell) value
            ActiveWorkbook.Names.Add _
              Name:="WP_LIST_Unique_Alpha", _
              RefersTo:="='WP #''s'!$I$2:$I$" & Config_WP_lastrow

        End If
        
        ' MsgBox "Cell " & Target.Address & " has changed."
       
    End If
End Sub
