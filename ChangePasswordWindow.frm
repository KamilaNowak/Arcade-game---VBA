Private Sub OK_Click()
' Check old password
If PreviousPasswordField.Text <> Range("A1").Value Then
MsgBox "Incorrect previous password."
PreviousPasswordField.Text = ""
Exit Sub
End If

' Check new passwords
If NewPasswordField.Text <> RepeatPasswordField.Text Then
MsgBox "Passwords are different. "
NewPasswordField.Text = ""
RepeatPasswordField.Text = ""
Exit Sub
End If

' if all is good do these instructions
Range("A1").Value = NewPasswordField.Text
ActiveWorkbook.Save
MsgBox " Operation finished successfully."
Me.Hide
Login.Show
End Sub
'Show or not password fields
Private Sub ShowPassword_Click()
PreviousPasswordField.PasswordChar = IIf(ShowPassword, "", "*")
NewPasswordField.PasswordChar = IIf(ShowPassword, "", "*")
RepeatPasswordField.PasswordChar = IIf(ShowPassword, "", "*")
End Sub
Private Sub UserForm_Terminate()
Login.Show
End Sub
