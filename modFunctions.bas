Attribute VB_Name = "modFunctions"
'For timing & pause procedure
Public Declare Function GetTickCount Lib "kernel32" () As Long

'to see if the mouse has moved
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

'for showing/hiding the cursor
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

'for moving & clicking the mouse
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cbuttons As Long, ByVal dwExtraInfo As Long)
Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4

Public Type POINTAPI
   X As Long
   Y As Long
End Type

Public Start As POINTAPI

Public Sub MouseClick(ByVal X As Long, ByVal Y As Long)
'Moves the mouse to X & Y coords then clicks the left mouse button
Dim cbuttons As Long
Dim dwExtraInfo As Long
Dim mevent As Long

SetCursorPos X, Y 'set the mouse pos

mevent = MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_LEFTUP 'set event

mouse_event mevent, 0&, 0&, cbuttons, dwExtraInfo 'click button
End Sub

Public Sub ScrollPage()
'Scrolls the current maximized window

ww = Screen.Width / Screen.TwipsPerPixelX  'get screen width
hh = Screen.Height / Screen.TwipsPerPixelY 'get screen height

'set the start pos here so the program won't end
Start.X = ww - 10
Start.Y = hh - 10

Randomize
For i = 0 To Int(Rnd * 8) + 1 'click mouse a random amount
    MouseClick ww - 10, hh - 10 'coords at bottom arrow
    Pause 0.1 'make it pause a sec
Next i
End Sub

Public Sub Pause(Interval As Integer)
'pauses for specified seconds
Interval = Interval * 1000 'convert to milliseconds
StartTime = GetTickCount 'get current tick
Do Until GetTickCount >= StartTime + Interval 'loop until tick = start + interval
    DoEvents 'do nothing
Loop
End Sub

Sub Main()
If App.PrevInstance Then End
If Left(Command$, 1) = "c" Or Right(Left(Command$, 2), 1) = "c" Then
    frmConfig.Show 1
ElseIf Left(Command$, 1) = "p" Or Right(Left(Command$, 2), 1) = "p" Then
    End
Else
    frmMain.Show
    Exit Sub
End If
End
End Sub
