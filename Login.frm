VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Login 
   Caption         =   "Log in"
   ClientHeight    =   6855
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6390
   OleObjectBlob   =   "Login.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Button Change Password
Private Sub ChangePasswordButton_Click()
Me.Hide
ChangePasswordWindow.Show
End Sub

Private Sub Label1_Click()

End Sub


'Button START
Private Sub StartButton_Click()
' A1 = reserve for password
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
