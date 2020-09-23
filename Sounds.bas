Attribute VB_Name = "modSnds"
Option Explicit

Private Declare Function sndPlay Lib "Winmm.dll" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long

Private Const SND_ASYNC = &H1     ' Play asynchronously
Private Const SND_NODEFAULT = &H2 ' Don't use default sound
Private Const SND_MEMORY = &H4    ' lpszSoundName points to a memory file
Private Const SND_FLAGS = SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY

Public Enum eSnd
    complete = 101
    moreinfo = 102
    errormsg = 103
End Enum

Private aBuf1() As Byte
Private aBuf2() As Byte
Private aBuf3() As Byte

Private fLoaded As Boolean

Public Sub LoadSounds()
    If fLoaded Then Exit Sub
    aBuf1 = LoadResData(complete, "SOUNDS")
    aBuf2 = LoadResData(moreinfo, "SOUNDS")
    aBuf3 = LoadResData(errormsg, "SOUNDS")
    fLoaded = True
End Sub

Public Sub PlaySound(ByVal ResId As eSnd)
    If fLoaded Then Else LoadSounds
    Select Case ResId
        Case complete: sndPlay aBuf1(0), SND_FLAGS
        Case moreinfo: sndPlay aBuf2(0), SND_FLAGS
        Case errormsg: sndPlay aBuf3(0), SND_FLAGS
    End Select
End Sub
