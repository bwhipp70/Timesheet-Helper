

Private Sub Cancel_Button_Click()
    RACFPassword = ""
    Unload Me
End Sub

Private Sub OK_Button_Click()
    RACFPassword = RACF_PW_TextBox.Value
    Unload Me
End Sub