VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Window 
   Caption         =   "Window"
   ClientHeight    =   7350
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12300
   OleObjectBlob   =   "Window.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Window"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Character_Click()
If CurrentInGame = False Then Exit Sub
ClicksNow = ClicksNow + 1
ClicksField.Text = ClicksNow
MsgBox "Congrats. Hit in " & ClicksNow & " click "
SeconsNow = Val(Mid(ClickDuration.Text, 1, 2))
ClickTime.Text = Val(Mid(ClickDuration.Text, 1, 2))

CurrentInGame = False
InGame = False
ClicksCounter = Val(ClicksField.Text)
ClicksNow = 0
ClicksField.Text = "0"
End Sub
Private Sub ClickDuration_Change()
ClickTime.Text = Val(Mid(ClickDuration, 1, 2))
SeconsNow = Val(ClickTime.Text)
End Sub

Private Sub MaxClick_Change()
MsgBox " Value of maximum clicks has changed on " & MaxClick.Text

End Sub

Private Sub StartReset_Click()
CurrentInGame = True
InGame = True
ClicksCounter = Val(MaxClick.Text)
ClicksNow = 0
ClicksField.Text = ClicksNow
SecondsNow = Val(Mid(ClickDuration.Text, 1, 2))
ClickTime = Val(Mid(ClickDuration.Text, 1, 2))
CountTime
Random
End Sub

Private Sub StopButton_Click()
If InGame = False Then Exit Sub
If CurrentInGame = True Then
CurrentInGame = False
StopButton.Caption = "RESUME"
Else
CurrentInGame = True
StopButton.Caption = "STOP"
CountTime
Random
End If
End Sub

Private Sub UserForm_Terminate()
MsgBox "Thanks for your game! Hope you had fun. See ya! "
Application.Quit
End Sub

Private Sub UserForm_Click()
Beep
If CurrentInGame = False Then Exit Sub
ClicksNow = ClicksNow + 1
ClicksField.Text = ClicksNow
If ClicksNow >= ClicksCounter Then
CurrentInGame = False
InGame = False
MsgBox " GAME OVER. Clicks Limit Exceed."
ClicksCounter = Val(MaxClick.Text)
ClicksNow = 0
ClicksField.Text = ClicksNow
SecondsNow = Val(Mid(ClickDuration.Text, 1, 2))
ClickTime.Text = Val(Mid(ClickDuration, 1, 2))
End If
End Sub

