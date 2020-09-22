Attribute VB_Name = "modSound"
'SOUND MOD
'I use this for all my programs that use sound
'Sorry it's not commented, but it's pretty self explanatory

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long

Private Const SND_ASYNC = &H1

Public Function SoundSupported() As Boolean
If (waveOutGetNumDevs > 0) Then
    SoundSupported = True
Else
    SoundSupported = False
End If
End Function

Public Sub PlaySound(sPath As String)
If SoundSupported = True Then
    If FileExist(sPath) = True Then
        Call sndPlaySound(sPath, SND_ASYNC)
    Else
        MsgBox "The sound file specified was not found", vbCritical, "File Not Found"
    End If
Else
    MsgBox "Your system does not support sound.  The sound cannot be played.", vbCritical, "No Sound"
End If
End Sub

Private Function FileExist(ByVal FileName As String) As Boolean
On Error Resume Next
Dim fileFile As Integer
fileFile = FreeFile
Open FileName For Input As fileFile
If Err Then
    FileExist = False
Else

    Close fileFile
    FileExist = True
End If
End Function
