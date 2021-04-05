Private Sub Worksheet_Change(ByVal Target As Range)
    ' Updated Ver 3.21 to also check for Work Schedule
    
    Dim KeyCells As Range
    Dim KeyCells2 As Range

    ' The variable KeyCells contains the cells that will
    ' cause an alert when they are changed.
    Set KeyCells = Range("WP_Dropdown")
    Set KeyCells2 = Range("WorkSchedule")
    
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

    If Not Application.Intersect(KeyCells2, Range(Target.Address)) _
           Is Nothing Then

        If (Range("WorkSchedule").Value = "Flex 4/10") Then
            Sheets("Timesheet").Unprotect
            Worksheets("Timesheet").Columns("F").Hidden = True
            Sheets("Timesheet").Protect
        Else
            Sheets("Timesheet").Unprotect
            Worksheets("Timesheet").Columns("F").Hidden = False
            Sheets("Timesheet").Protect
        End If
        
        Call Update_Work_Schedule_Selection     ' This will hide either the Labor_Flex980 or Labor_Flex980_2weeks sheet
        
        'MsgBox "WorkSchedule Cell " & Target.Address & " has changed."
       
    End If

End Sub
