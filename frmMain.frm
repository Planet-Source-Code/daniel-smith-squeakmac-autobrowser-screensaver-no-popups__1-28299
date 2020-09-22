VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5775
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6615
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   6615
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrMouse 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4800
      Top             =   120
   End
   Begin VB.Timer tmrTxt 
      Interval        =   500
      Left            =   5760
      Top             =   120
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   200
      Left            =   30
      TabIndex        =   1
      Text            =   "Searching..."
      Top             =   30
      Width           =   855
   End
   Begin VB.Timer tmrStart 
      Interval        =   1000
      Left            =   5280
      Top             =   120
   End
   Begin SHDocVwCtl.WebBrowser w 
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      ExtentX         =   11668
      ExtentY         =   10186
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is the AutoBrowser.  You give it a topic to search off of
'and it starts a yahoo search on it.  Note that the topic is
'only something for it to start with, and sometimes it goes
'off wherever, because it is random.
'However, several modifications have been made to keep it more
'on subject.  Especially on Yahoo!, because it would tend to get
'stuck in the Help area and Terms of Service.


Dim Waiting As Boolean 'for waiting for page to load
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Private Sub Form_Load()
ShowCursor 0 'hide mouse
Show 'show form
DoEvents
Randomize Time 'randomize
sndPlaySound "C:\comp11.wav", &H8 + &H1  'play sound looping
Me.WindowState = 2 'maximize window
StartBrowse 'start the auto browse
End Sub

Private Sub Form_Resize()
'resize browser to fit window
w.Width = ScaleWidth
w.Height = ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
sndPlaySound App.Path & "\fmt.wav", &H1  'play a simple sound that has no sound (to stop sound)
ShowCursor 1 'show mouse
End 'end program
End Sub

Private Sub tmrMouse_Timer()
Dim NewP As POINTAPI
GetCursorPos NewP 'get the cursors current position

If NewP.X <> Start.X Then 'if the new X is different then the starting X
    Unload Me 'end program (mouse moved)
ElseIf NewP.Y <> Start.Y Then 'if the new Y is different then the starting Y
    Unload Me 'end program (mouse moved)
End If
End Sub

Sub StartBrowse()
Dim TempInt As Integer
Randomize Time
TempInt = Int(Rnd * 2)
If TempInt = 0 Then SearchYahoo
If TempInt = 1 Then SearchGoogle
End Sub

Sub SearchGoogle()
On Error Resume Next
'here's the browser!
Dim TempInt As Integer
Dim TempString As String
Dim PageCount As Integer

Begin:
PageCount = 0
'Google
w.Navigate "www.google.com" 'navigate to google
Wait 'wait for download to finsih
Pause 3 'um, letsee....

Randomize Time
TempInt = Int(Rnd * 4) 'get a random num to select topic
Select Case TempInt 'you can change the topics to whatever
    Case 0
        TempString = GetSetting("AutoBrowse", "Topics", "Topic1", "Visual Basic")
    Case 1
        TempString = GetSetting("AutoBrowse", "Topics", "Topic2", "Bin Laden")
    Case 2
        TempString = GetSetting("AutoBrowse", "Topics", "Topic3", "FBI")
    Case 3
        TempString = GetSetting("AutoBrowse", "Topics", "Topic4", "Games")
End Select
TempString = TempString 'replace the second TempString with a topic for static topic
w.Document.Forms(0)(0).Value = TempString 'fill in the yahoo form with the topic
Pause 1
w.Document.Forms(0).submit 'start the search (clicks search button)
Wait 'wait for download


'Search Page on Google
SearchMore:
PageCount = PageCount + 1

Me.SetFocus 'set focus on form

Pause 3
ScrollPage 'scroll down a bit (see modFunctions)
Pause 3

Randomize
If InStr(1, w.LocationURL, "search") <> 0 _
    And InStr(1, w.LocationURL, "google") <> 0 Then
    TempInt = 10 + Int(Rnd * (w.Document.links.length - 25)) 'pic a random link
Else
    TempInt = Int(Rnd * (w.Document.links.length)) + 1
End If
If w.Document.links.length = 0 Then StartBrowse: Exit Sub
w.Document.links(TempInt).Click 'click the link

Wait 'wait for download

'if we're still at a google site
If InStr(1, LCase(w.LocationURL), "google") <> 0 Then
    If PageCount >= 6 Then 'we're probably off track...
        GoTo Begin 'start over
    Else
        'make sure we're still searching
        If InStr(1, w.LocationURL, "google") <> 0 _
        And InStr(1, w.LocationURL, "search") = 0 Then
            Pause 3
            GoTo Begin 'restart search
        End If
        GoTo SearchMore 'go back and keep looking
    End If
Else
    PageCount = 0 'restart pagecount
End If


'other page on net
Other:

'if we've visited 7 pages outside of google, restart browse
If PageCount >= 7 Then StartBrowse: Exit Sub
PageCount = PageCount + 1

Me.SetFocus 'set focus on form

Pause 3
ScrollPage 'scroll down a bit (see modFunctions)
Pause 3

Randomize
pic:
TempInt = Int(Rnd * w.Document.links.length) 'pic a random link
w.Document.links(TempInt).Click 'click the link

Wait 'wait for download

GoTo Other 'do it again
End Sub

Sub SearchYahoo()
On Error Resume Next
'here's the browser!
Dim TempInt As Integer
Dim TempString As String
Dim PageCount As Integer

Begin:
PageCount = 0
'Yahoo main page
w.Navigate "www.yahoo.com" 'navigate to yahoo
Wait 'wait for download to finsih
Pause 3 'um, letsee....

Randomize Time
TempInt = Int(Rnd * 4) 'get a random num to select topic
Select Case TempInt 'you can change the topics to whatever
    Case 0
        TempString = GetSetting("AutoBrowse", "Topics", "Topic1", "Visual Basic")
    Case 1
        TempString = GetSetting("AutoBrowse", "Topics", "Topic2", "Bin Laden")
    Case 2
        TempString = GetSetting("AutoBrowse", "Topics", "Topic3", "FBI")
    Case 3
        TempString = GetSetting("AutoBrowse", "Topics", "Topic4", "Games")
End Select
TempString = TempString 'replace the second TempString with a topic for static topic
w.Document.Forms(0)(0).Value = TempString 'fill in the yahoo form with the topic
Pause 1
w.Document.Forms(0).submit 'start the search (clicks search button)
Wait 'wait for download


'Search Page on yahoo
SearchMore:
PageCount = PageCount + 1

Me.SetFocus 'set focus on form

Pause 3
ScrollPage 'scroll down a bit (see modFunctions)
Pause 3

Randomize
If InStr(1, w.LocationURL, "search") <> 0 _
    And InStr(1, w.LocationURL, "yahoo") <> 0 Then 'if on a yahoo search page
    TempInt = 16 + Int(Rnd * (w.Document.links.length - 21)) 'pic a random link
Else
    TempInt = Int(Rnd * (w.Document.links.length)) + 1 'if on other yahoo page
End If
If w.Document.links.length = 0 Then StartBrowse: Exit Sub
w.Document.links(TempInt).Click 'click the link

Wait 'wait for download

'if we're still at a yahoo site
If InStr(1, LCase(w.LocationURL), "yahoo") <> 0 Then
    If PageCount >= 6 Then 'we're probably off track...
        GoTo Begin 'start over
    Else
        If InStr(1, w.LocationURL, "yahoo") <> 0 _
        And InStr(1, w.LocationURL, "search") = 0 _
        And InStr(1, w.LocationURL, "google") = 0 Then 'make sure we're still searching
            Pause 3
            StartBrowse
            Exit Sub
        End If
        GoTo SearchMore 'go back and keep looking
    End If
Else
    PageCount = 0 'restart pagecount
End If


'other page on net
Other:

'if we've visited 7 pages outside of yahoo, restart browse
If PageCount >= 7 Then StartBrowse: Exit Sub
PageCount = PageCount + 1

Me.SetFocus 'set focus on form

Pause 3
ScrollPage 'scroll down a bit (see modFunctions)
Pause 3

Randomize
pic:
TempInt = Int(Rnd * w.Document.links.length) 'pic a random link
w.Document.links(TempInt).Click 'click the link

Wait 'wait for download

GoTo Other 'do it again
End Sub

Sub Wait()
Dim StartTime As Integer
StartTime = GetTickCount
Do Until Waiting = False 'if page isn't loaded
    DoEvents 'do nothing
    If GetTickCount - StartTime > 10000 Then 'if waiting too long
        w.GoBack 'go back to previous page
        Exit Sub
    End If
Loop
End Sub

Private Sub tmrStart_Timer()
GetCursorPos Start 'set the starting cursor pos
tmrMouse.Enabled = True 'start tracking mouse
tmrStart.Enabled = False 'disable this timer
End Sub

Private Sub tmrTxt_Timer()
Static Pos As Integer 'number of periods
Text1.Text = "Searching" & String(Pos, ".") 'print periods
Pos = Pos + 1 'increase number
If Pos = 4 Then Pos = 0 'reset number if needed
End Sub

Private Sub w_DownloadBegin()
Waiting = True 'downloading...
End Sub

Private Sub w_DownloadComplete()
Waiting = False 'download complete
End Sub

Private Sub w_NewWindow2(ppDisp As Object, Cancel As Boolean)
Cancel = True 'cancel popup
End Sub

