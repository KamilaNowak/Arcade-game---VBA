' Button Change Password
Private Sub ChangePasswordButton_Click()
Me.Hide
ChangePasswordWindow.Show
End Sub

'Button START
Private Sub StartButton_Click()
' A1 = reserved for password 
If Range("A1").Value = "" And PasswordField.Text = "" Then
MsgBox "Password is not entered. Change it on next startup, please."
Me.Hide
Window.Show
Else
If PasswordField.Text = Range("A1").Value Then
Me.Hide
Window.Show
Else: MsgBox "Incorrect Password! "
PasswordField.Text = ""
End If
End If
End Sub

Private Sub UserForm_Terminate()
MsgBox "Thanks for your visit! "
Application.Quit
End Sub

Private Sub EndGame_Click()
MsgBox "Thanks for your visit!"
Application.Quit
End Sub
