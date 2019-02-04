VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsZZTSound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' **************************************************************
'  clsZZTSound
'
'  A sound module for playing back ZZT sounds and music
'  SaxxonPike 2oo7-2oo8
' **************************************************************

' This system works by first pre-loading all the sounds into memory. Then
' it uses the PlaySound API to play the sounds. If it is the speaker, then the
' frequency is modified in the wave file itself and set on loop. Everything is
' played ASYNC so that the program's normal flow isn't interrupted. This is not
' an efficient way of doing sound, however it is the most compatible next to
' using DirectX.
'
' Notes are added to the play buffer and they are played back using a counter
' that decreases. When this counter is zero and there are more notes in the
' playback queue, it advances. You call the function to process one tick every
' time. This allows for a variance in how fast the notes are played back as
' determined by the program.

' PlaySound API and constants
Private Declare Function PlaySoundAny Lib "winmm.dll" Alias "PlaySoundA" (lpSound As Any, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Const SND_ASYNC = &H1
Private Const SND_FILENAME = &H20000
Private Const SND_LOOP = &H8
Private Const SND_MEMORY = &H4
Private Const SND_NODEFAULT = &H2

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

' these are all divisible 3 so you can do triplets
Private Const DurationT = 3
Private Const DurationS = 6
Private Const DurationI = 12
Private Const DurationQ = 24
Private Const DurationH = 48
Private Const DurationW = 96

Private Type SoundData
    xData() As Byte
End Type

Private SoundFiles(0 To 9) As SoundData
Private SoundBuffer As String
Private SoundTicksLeft As Long
Private TickStart As Currency
Private ToneOctave As Long
Private SoundTimer As New clsHPTimer
Private BaseSpeed As Long
Private ActualSpeed As Long
Private TripletsLeft As Long

Private Const TNoteDuration = 50

Public Sub AddToBuffer(inNotes As String)
    SoundBuffer = SoundBuffer + inNotes + Chr$(13)
End Sub

Public Sub StopPlaying()
    SoundBuffer = ""
    SoundTicksLeft = 0
    PlaySoundAny ByVal 0&, ByVal 0&, ByVal 0&
    ResetLine
End Sub

Private Sub WaitForSoundTick()
    Do While SoundTimer.TimeElapsed < TickStart
        Sleep 0
        Sleep 1
    Loop
End Sub

Private Sub SetTone(xSound As Long, ByVal xTone As Long)
    xTone = xTone - 9 'internal sound file constant
    Dim xFreq(0 To 1) As Long
    xFreq(0) = 88200 * (2 ^ ((xTone) / 12))
    xFreq(1) = xFreq(0) * 2
    CopyMemory SoundFiles(3).xData(24), xFreq(0), 8
End Sub

Private Sub Sound(soundnum As Long, Optional xloop As Boolean = False, Optional SetFreq As Long = 0)
    If soundnum = -1 Then
        PlaySoundAny ByVal 0&, ByVal 0&, ByVal 0&
        Exit Sub
    End If
    If xloop Then
        SetTone soundnum, SetFreq
        PlaySoundAny SoundFiles(soundnum).xData(0), ByVal 0&, SND_MEMORY Or SND_LOOP Or SND_ASYNC Or SND_NODEFAULT
    Else
        PlaySoundAny SoundFiles(soundnum).xData(0), ByVal 0&, SND_MEMORY Or SND_ASYNC Or SND_NODEFAULT
    End If
End Sub

Public Sub DebugPlaySound(soundnum As Long)
    Sound soundnum
End Sub

Public Sub LoadSoundResource(soundnum As Long, resnum As Long)
    SoundFiles(soundnum).xData() = LoadResData(resnum, "WAVE")
End Sub

Public Sub ProcessCycle()
    Dim x As String
    Dim donote As Boolean
    Dim toneset As Long
    SoundTicksLeft = SoundTicksLeft - 1
    WaitForSoundTick
    If SoundTicksLeft <= 0 Then
        If SoundBuffer = "" Then
            SoundTicksLeft = 0
        Else
            Do While SoundBuffer <> "" And SoundTicksLeft <= 0
                x = UCase$(Left$(SoundBuffer, 1))
                SoundBuffer = Mid$(SoundBuffer, 2)
                Select Case x
                    'duration
                    Case "W": SetDuration DurationW
                    Case "H": SetDuration DurationH
                    Case "Q": SetDuration DurationQ
                    Case "I": SetDuration DurationI
                    Case "S": SetDuration DurationS
                    Case "T": SetDuration DurationT
                    Case ".": ApplyDot
                    Case "3": TripletsLeft = 3
                    'octave
                    Case "+": OctaveUp
                    Case "-": OctaveDown
                    'notes
                    Case "X": donote = True: toneset = -256
                    Case "C": donote = True: toneset = (ToneOctave * 12) + 0
                    Case "D": donote = True: toneset = (ToneOctave * 12) + 2
                    Case "E": donote = True: toneset = (ToneOctave * 12) + 4
                    Case "F": donote = True: toneset = (ToneOctave * 12) + 5
                    Case "G": donote = True: toneset = (ToneOctave * 12) + 7
                    Case "A": donote = True: toneset = (ToneOctave * 12) + 9
                    Case "B": donote = True: toneset = (ToneOctave * 12) + 11
                    Case "#": ' do nothing for these two. if they show up here
                    Case "!": ' they are junk. they must be after a note.
                    'drums
                    Case "0": donote = True: toneset = 256
                    Case "1": donote = True: toneset = 257
                    Case "2": donote = True: toneset = 258
                    Case "4": donote = True: toneset = 260
                    Case "5": donote = True: toneset = 261
                    Case "6": donote = True: toneset = 262
                    Case "7": donote = True: toneset = 263
                    Case "8": donote = True: toneset = 264
                    Case "9": donote = True: toneset = 265
                    'extra
                    Case Chr$(13): ResetLine
                End Select
                If donote Then
                    If toneset > -256 Then
                        If toneset < 256 Then
                            If Left$(SoundBuffer, 1) = "#" Then
                                toneset = toneset + 1
                                SoundBuffer = Mid$(SoundBuffer, 2)
                            ElseIf Left$(SoundBuffer, 1) = "!" Then
                                toneset = toneset - 1
                                SoundBuffer = Mid$(SoundBuffer, 2)
                            End If
                            Sound 3, True, toneset
                        Else
                            Sound toneset - 256
                        End If
                    Else
                        Sound -1
                    End If
                    SoundTicksLeft = ActualSpeed
                    If TripletsLeft > 0 Then
                        TripletsLeft = TripletsLeft - 1
                        SoundTicksLeft = SoundTicksLeft / 3
                    End If
                End If
            Loop
        End If
    End If
    TickStart = SoundTimer.TimeElapsed + (TNoteDuration / 3) '1/96th of a second or 1 unit
End Sub

Private Sub ResetLine()
    ToneOctave = 0
    TripletsLeft = 0
    SetDuration DurationT
    PlaySoundAny ByVal 0&, ByVal 0&, ByVal 0&
End Sub

Private Sub SetDuration(newdur As Long)
    BaseSpeed = newdur
    ActualSpeed = newdur
End Sub

Private Sub Class_Initialize()
    ResetLine
    StopPlaying
    TickStart = SoundTimer.TimeElapsed
End Sub

Private Sub OctaveUp()
    If ToneOctave < 2 Then
        ToneOctave = ToneOctave + 1
    End If
End Sub

Private Sub OctaveDown()
    If ToneOctave > -2 Then
        ToneOctave = ToneOctave - 1
    End If
End Sub

Private Sub ApplyDot()
    ActualSpeed = ActualSpeed + (((BaseSpeed * 2) - ActualSpeed) / 2)
End Sub
