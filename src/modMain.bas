Attribute VB_Name = "modMain"
' we will be using Sub Main to ensure that we close fully and
' safely.

Option Explicit

Private FormsEnableCount As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SetWindowPos Lib "USER32" (ByVal hWnd As Long, _
ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, _
ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Type xtiBufferStats
    xID As Byte
    xColor As Byte
    xUseStats As Boolean
    xIsFilled As Boolean
    xXStep As Integer
    xYStep As Integer
    xCycle As Integer
    xP1 As Byte
    xP2 As Byte
    xP3 As Byte
    xFollow As Integer
    xLeader As Integer
    xPointer As Long
    xInstruction As Integer
    xLength As Integer
    xOOP As String
    xUnderColor As Byte
    xUnderID As Byte
End Type
Private Type xtiBoardInformation
    xshots As Byte
    xdark As Byte
    xNorth As Byte
    xSouth As Byte
    xWest As Byte
    xEast As Byte
    xRestart As Byte
    xTimeLimit As Integer
End Type

Public ExHeight As Long
Public ExWidth As Long

Public MainStatBuffer As xtiBufferStats
Public BoardInfo As xtiBoardInformation
Public ForegroundWindow As Long

Public Const BOARD_HEIGHT = 25
Public Const BOARD_WIDTH = 60

Public bRunning As Boolean
Public iUseScale As Single
Public CurrentBoard As Long
Public World As New clsZZTWorld
Public ObjectLibrary As New clsZZTLibrary
Public Sound As New clsZZTSound
Public Help As New clsZZTHelp

Public FloodFillPattern(0 To 7, 0 To 7) As Boolean
Public bFillPattern As Boolean
Public lFillSizeX As Long
Public lFillSizeY As Long
Public bSkipStats As Boolean

Public MainTickCounter As Long

Sub Main()

    'resolution check
    If (Screen.Width \ Screen.TwipsPerPixelX) < 800 Or (Screen.Height \ Screen.TwipsPerPixelY) < 600 Then
        If MsgBox("Your screen resolution may be too low to use this editor effectively." + vbCrLf + "Run it anyway?", vbYesNo, "Low Resolution") = vbNo Then
            End
        End If
    End If

    'zzt engine init
    World.NewWorld False
    Dim s As String
    
    'init
    SetupSound
    GetRecentFiles
    frmEdit.Show
    frmEdit.Visible = True
    DoEvents
    SetFontDCResource frmEdit.picMain.hdc, 101, 8, 16
    SetEditScale 1
    SetTargetDC frmEdit.picMain.hdc, 60, 25
    DoEvents
    SetupPalette
    frmEdit.ShowToolColors
    bRunning = True
    DrawBoard
    frmEdit.RedrawBuffer
    frmEdit.RedrawBuffer2
    
    'load anything on the command line input
    If Command$ <> "" Then
        If Left$(Command$, 1) <> Chr$(34) Then
            frmEdit.LoadFile Command$, False
        Else
            s = Mid$(Command$, 2)
            If InStr(s, Chr$(34)) > 0 Then
                s = Left$(s, InStr(s, Chr$(34)) - 1)
                frmEdit.LoadFile Command$, False
            End If
        End If
    End If
    
    frmEdit.LoadConfig
    
    'main loop
    Do While bRunning
        Sound.ProcessCycle
        DoEvents
    Loop
    
    'unload forms
    Unload frmEdit
    
    'unload other modules
    UnloadFontSystem
    Set World = Nothing
    
    On Error Resume Next
    If Dir(AppPath + "@zaptest.zzt") <> "" Then
        Kill AppPath + "@zaptest.zzt"
    End If
    If Dir(AppPath + "@zaptest.szt") <> "" Then
        Kill AppPath + "@zaptest.szt"
    End If
    If Dir(AppPath + "@zapexec.bat") <> "" Then
        Kill AppPath + "@zapexec.bat"
    End If
    
    Sound.StopPlaying
    SaveRecentFiles
    
    End
    
End Sub

Sub SetupPalette()
    Dim x As Long
    Dim r As Byte
    Dim g As Byte
    Dim b As Byte
    Dim PaletteRGB(0 To 15) As RGB
    For x = 0 To 15
        r = (((x And 4) \ 4) * 168) + (((x And 8) \ 8) * 84)
        g = (((x And 2) \ 2) * 168) + (((x And 8) \ 8) * 84)
        b = (((x And 1) \ 1) * 168) + (((x And 8) \ 8) * 84)
        If x = 6 Then g = 84 'special case for brown
        If x > 8 And x < 15 Then g = g - 1 'non-greys
        If x <> 0 Then
            r = r + 3
            g = g + 3
            b = b + 3
        End If
        SetColor x, r, g, b
        PaletteRGB(x).r = r
        PaletteRGB(x).g = g
        PaletteRGB(x).b = b
        frmEdit.SetToolColor x, RGB(r, g, b)
    Next x
    BMP2BRD_CreateLookupTable PaletteRGB()
End Sub

Public Sub DisableForms()
    FormsEnableCount = FormsEnableCount + 1
    frmEdit.Enabled = False
    'frmPalette.Enabled = False
End Sub

Public Sub EnableForms()
    If FormsEnableCount > 1 Then
        FormsEnableCount = FormsEnableCount - 1
    Else
        frmEdit.Enabled = True
        FormsEnableCount = 0
    End If
    'frmPalette.Enabled = True
End Sub

Public Sub SetEditScale(Optional ByVal iNewScale As Single = -1)
    If iNewScale > -1 Then
        iUseScale = iNewScale
    Else
        iNewScale = iUseScale
    End If
    frmEdit.Width = (Screen.TwipsPerPixelX * (FontDCWidth * 60) * iNewScale) + ExWidth
    frmEdit.Height = (Screen.TwipsPerPixelY * (FontDCHeight * 25) * iNewScale) + ExHeight
    frmEdit.picMain.Height = (Screen.TwipsPerPixelY * FontDCHeight * 25 * iNewScale)
    DoEvents
    SetTargetDC frmEdit.picMain.hdc, 60, 25
End Sub

Public Sub DrawBoard()
    Dim x As Long
    Dim y As Long
    Dim ch As Long
    Dim c As Long
    For x = 0 To BOARD_WIDTH - 1
        For y = 0 To BOARD_HEIGHT - 1
            With World
                c = .BoardCol(CurrentBoard, x, y)
                ch = DefaultChar(.BoardID(CurrentBoard, x, y), c)
                SetCharB x, y, CByte(ch), CByte(c)
                If ch <> 32 Then
                    x = x
                End If
            End With
        Next y
    Next x
End Sub

Public Function AppPath() As String
    Dim s As String
    s = App.Path
    If Right$(s, 1) <> "\" Then
        s = s + "\"
    End If
    AppPath = s
End Function

Public Sub SetCurrentBoard(newboard As Long)
    CurrentBoard = newboard
    With BoardInfo
        World.GetBoardInfo CurrentBoard, .xshots, .xdark, .xRestart, .xNorth, .xSouth, .xWest, .xEast, .xTimeLimit
    End With
End Sub

Public Sub ShowAboutWindow()
    Dim aboutmsg As String
    aboutmsg = "ZAP is a windows-based ZZT editor by SaxxonPike - 2oo8" + vbCrLf + vbCrLf + "Huge thanks to Tim Sweeney for the original ZZT." + vbCrLf + "Also, thanks to Jacob Hammond and Kev Vance for their documentation." + vbCrLf + "And let's not forget http://zzt.belsambar.net for continued efforts to preserve ZZT."
    MsgBox aboutmsg, vbInformation, "About " + App.Title + " (revision " + CStr(App.Revision) + ")"
End Sub

Sub SetupSound()
    Dim x As Long
    Sound.LoadSoundResource 0, 102
    Sound.LoadSoundResource 1, 103
    Sound.LoadSoundResource 2, 104
    Sound.LoadSoundResource 4, 105
    Sound.LoadSoundResource 5, 106
    Sound.LoadSoundResource 6, 107
    Sound.LoadSoundResource 7, 108
    Sound.LoadSoundResource 8, 109
    Sound.LoadSoundResource 9, 110
    Sound.LoadSoundResource 3, 111
End Sub
