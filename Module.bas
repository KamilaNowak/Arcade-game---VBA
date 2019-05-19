
Option Explicit
Public InGame As Boolean  ' on game
Public CurrentInGame As Boolean      ' currently on game
Public ClicksCounter, ClicksNow, SecondsNow As Integer
Public Leftt, Up, LeftMax, UpMax, UpMin As Integer
Public Interval As String

Sub auto_open()
LeftMax = Window.Width - Window.Character.Width
UpMax = Window.Height - Window.Character.Height
UpMin = Window.StartReset.Height
Window.ClickDuration.AddItem "5 seconds"
Window.ClickDuration.AddItem "7 seconds"
Window.ClickDuration.AddItem "9 seconds"
Window.ClickDuration.AddItem "12 seconds"
Window.ClickDuration.AddItem "15 seconds"
Window.ClickDuration.AddItem "20 seconds"
Window.MaxClick.AddItem "1"
Window.MaxClick.AddItem "2"
Window.MaxClick.AddItem "3"
Window.MaxClick.AddItem "4"
Window.MaxClick.AddItem "5"
Window.MaxClick.AddItem "6"
Window.MaxClick.AddItem "7"
Window.MaxClick.AddItem "8"
Window.MaxClick.AddItem "9"
Window.MaxClick.AddItem "10"
Window.MaxClick.AddItem "11"
Application.Visible = False
ClicksCounter = Val(Window.MaxClick.Text)
ClicksNow = 0
Window.ClicksField.Text = ClicksNow
InGame = False
Login.Show
End Sub

Sub Random()
Interval = "00:00:" & Str(Window.Speed.Value)
ClicksCounter = Val(Window.MaxClick.Text)
If CurrentInGame = False Then Exit Sub
NewPosition
Application.OnTime Now + TimeValue(Interval), "Random"
End Sub
Sub NewPosition()
Leftt = Int(Rnd * LeftMax)
Up = Int(Rnd * (UpMax - UpMin) + UpMin)
Window.Character.Left = Leftt
Window.Character.Top = Up
End Sub
Sub TimeOver()
MsgBox "The time is over! "
CurrentInGame = False
InGame = False
ClicksCounter = Val(Window.MaxClick.Text)
ClicksNow = 0
Window.ClicksField.Text = ClicksNow
Window.ClickTime.Text = Val(Mid(Window.ClickDuration.Text, 1, 2))
SecondsNow = Val(Window.ClickTime.Text)
End Sub
Sub CountTime()
If CurrentInGame = False Then Exit Sub
SecondsNow = SecondsNow - 1
Window.ClickTime.Text = SecondsNow
If SecondsNow <= 0 Then
TimeOver
Exit Sub
End If
Application.OnTime Now + TimeValue("00:00:01"), "CountTime"
End Sub
