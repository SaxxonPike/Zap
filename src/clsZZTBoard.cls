VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsZZTBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' **************************************************************
'  clsZZTBoard
'
'  ZZT board class for Visual Basic 6
'  SaxxonPike 2oo7-2oo8
' **************************************************************
'
' ZZT: 1500: 60x25 boards
' SZT: 7680: 96x80 boards (scrollable)
'

' Proposed SuperZZT prodecure:
' * load SuperZZT file using appropriate types
'
' * convert format to original ZZT
'    NOTE: THIS MEANS ALL ASSOCIATED TILE TYPES AS WELL
'          e.g. E_SZTBullet -> E_Bullet, etc.
'
' * set board height and width to 80 and 96 respectively
' * activate super ZZT mode for special routine routing



Option Explicit

Private Const MAXOBJECTS = 151
Private Const DEFAULT = &HFEDC

Private Const TILES_ZZT = 1500
Private Const TILES_SUPERZZT = 7680

Private Type xtiRLE
    xNum As Byte
    xID As Byte
    xCol As Byte
End Type

Private Type xtiStatElement
    xX As Byte
    xY As Byte
    xXStep As Integer
    xYStep As Integer
    xCycle As Integer
    xP1 As Byte
    xP2 As Byte
    xP3 As Byte
    xFollow As Integer
    xLeader As Integer
    xUnderID As Byte
    xUnderColor As Byte
    xPointer As Long
    xInstruction As Integer
    xLength As Integer
    xUnk0 As String * 8
End Type

Private Type xtiSuperStatElement
    xX As Byte
    xY As Byte
    xXStep As Integer
    xYStep As Integer
    xCycle As Integer
    xP1 As Byte
    xP2 As Byte
    xP3 As Byte
    xFollow As Integer
    xLeader As Integer
    xUnderID As Byte
    xUnderColor As Byte
    xPointer As Long
    xInstruction As Integer
    xLength As Integer
End Type

Private Type xtiHeader
    xSize As Integer
    xTitleSize As Byte
    xTitle As String * 50
End Type

Private Type xtiSuperHeader
    xSize As Integer
    xTitleSize As Byte
    xTitle As String * 60
End Type

Private xData() As xtiRLE

Private Type xtiHeader2
    xshots As Byte
    xdark As Byte
    xNorth As Byte
    xSouth As Byte
    xWest As Byte
    xEast As Byte
    xRestart As Byte
    xMessageSize As Byte
    xMessage As String * 58
    xEnterX As Byte
    xEnterY As Byte
    xTimeLimit As Integer
    xUnk0(0 To 15) As Byte
    xStatElements As Integer
End Type

Private Type xtiSuperHeader2
    xshots As Byte
    xNorth As Byte
    xSouth As Byte
    xWest As Byte
    xEast As Byte
    xRestart As Byte
    xMessageSize As Byte
    'xUnk1 As Long
    'xTimeLimit As Integer
    xUnk0(0 To 20) As Byte
    xStatElements As Integer
End Type

Private Type xtiTile
    xColor As Byte
    xType As Byte
End Type

Private Type xtiOOP
    xOOP() As Byte
End Type

Private xBoardData() As xtiTile
Private xBoardHeader As xtiHeader
Private xBoardHeader2 As xtiHeader2
Private xBoardStats() As xtiStatElement

Private xSuperHeader As xtiSuperHeader
Private xSuperHeader2 As xtiSuperHeader2
Private xSuperBoardStats() As xtiSuperStatElement

Private xBoardWidth As Long
Private xBoardHeight As Long
Private xBoardTiles As Long

Private xEmptyStat As xtiStatElement
Private xOOP() As xtiOOP
Private xMaxStat As Long

Public Function BoardName() As String
    BoardName = Left$(xBoardHeader.xTitle, xBoardHeader.xTitleSize)
End Function

Public Sub SetBoardName(bn As String)
    xBoardHeader.xTitle = bn + String(50, 0)
    xBoardHeader.xTitleSize = Len(bn)
    If xBoardHeader.xTitleSize > 50 Then
        xBoardHeader.xTitleSize = 50
    End If
End Sub

Public Function ObjectCharAt(x As Long, y As Long, anim As Long) As Byte
    
    'this function is messy as hell and needs to be cleaned up.
    ' -saxx
    
    Dim z As Long
    'just in case we have something statless, fall back to defaults
    ObjectCharAt = DefaultChar(BoardID(x, y), BoardCol(x, y))
    z = ObjectAt(x, y)
    If z <> -1 Then
        With xBoardStats(z)
            Select Case BoardID(x, y)
                Case E_Object
                    ObjectCharAt = .xP1
                Case E_Pusher
                    If .xXStep < 0 And .xYStep = 0 Then 'left
                        ObjectCharAt = 17
                    ElseIf .xYStep < 0 And .xXStep = 0 Then 'up
                        ObjectCharAt = 30
                    ElseIf .xYStep > 0 And .xXStep = 0 Then 'down
                        ObjectCharAt = 31
                    Else                                'right
                        ObjectCharAt = 16
                    End If
                Case E_Transporter
                    If .xXStep < 0 And .xYStep = 0 Then 'left
                        ObjectCharAt = Asc(Mid$("<(" + Chr$(179) + "(", anim, 1))
                    ElseIf .xYStep < 0 And .xXStep = 0 Then 'up
                        ObjectCharAt = Asc(Mid$("~^" + Chr$(196) + "^", anim, 1))
                    ElseIf .xYStep > 0 And .xXStep = 0 Then 'down
                        ObjectCharAt = Asc(Mid$("_v" + Chr$(196) + "v", anim, 1))
                    ElseIf .xXStep > 0 And .xYStep = 0 Then 'right
                        ObjectCharAt = Asc(Mid$(">)" + Chr$(179) + ")", anim, 1))
                    Else                        'none of the above
                        ObjectCharAt = Asc(Mid$("^" + Chr$(196) + "v_", anim, 1))
                    End If
                Case E_Star, E_Clockwise
                    ObjectCharAt = Asc(Mid$(Chr$(179) + "/" + Chr$(196) + "\", anim, 1))
                Case E_Counter
                    ObjectCharAt = Asc(Mid$("/" + Chr$(179) + "\" + Chr$(196), anim, 1))
                Case E_SpinningGun
                    ObjectCharAt = Asc(Mid$(Chr$(24) + Chr$(26) + Chr$(25) + Chr$(27), anim, 1))
                Case E_DragonPup
                    ObjectCharAt = Asc(Mid$(Chr$(149) + Chr$(148) + Chr$(162) + Chr$(148), anim, 1))
            End Select
        End With
    End If
End Function
Public Function BoardID(x As Long, y As Long) As Byte
    BoardID = xBoardData(x + (y * xBoardWidth)).xType
End Function
Public Function BoardCol(x As Long, y As Long) As Byte
    BoardCol = xBoardData(x + (y * xBoardWidth)).xColor
End Function
Public Sub SetBoardID(x As Long, y As Long, i As Long, bStats As Boolean)
    Dim z As Long
    Dim statnum As Long
    Dim uid As Long
    Dim ucol As Long
    
    'we won't overwrite the main player object. however if we are using
    'a fake or empty, place it under the player instead
    With xBoardStats(0)
        If ((.xX - 1) = x And (.xY - 1) = y) Then
            If i = E_Empty Or i = E_Fake Then
                .xUnderID = i
                .xPointer = 1234
            End If
            Exit Sub
        End If
    End With
    
    'delete any stat objects we've overwritten
    For z = 1 To xBoardHeader2.xStatElements
        With xBoardStats(z)
            If ((.xX - 1) = x And (.xY - 1) = y) Then
                DeleteStat z
                Exit For
            End If
        End With
    Next z

    xBoardData(x + (y * xBoardWidth)).xType = i
End Sub
Public Sub SetBoardCol(x As Long, y As Long, i As Long)
    Dim z As Long
    'we won't overwrite the main player object. however if we are using
    'a fake or empty, place it under the player instead
    With xBoardStats(0)
        If ((.xX - 1) <> x Or (.xY - 1) <> y) Then
            xBoardData(x + (y * xBoardWidth)).xColor = i
        Else
            If .xPointer = 1234 Then 'this was previously set by SetBoardID (hack)
                .xPointer = 0
                .xUnderColor = i
            End If
        End If
    End With
End Sub
Public Sub MovePlayer(x As Long, y As Long)
    If Not (x >= 0 And x < xBoardWidth And y >= 0 And y < xBoardHeight) Then
        Exit Sub
    End If
    With xBoardStats(0)
        If (.xX - 1) = x And (.xY - 1) = y Then
            Exit Sub
        End If
        If .xX > 0 And .xX < (xBoardWidth - 1) And .xY > 0 And .xY < (xBoardHeight - 1) Then
            xBoardData((.xX - 1) + ((.xY - 1) * xBoardWidth)).xColor = .xUnderColor
            xBoardData((.xX - 1) + ((.xY - 1) * xBoardWidth)).xType = .xUnderID
        End If
        Select Case xBoardData(x + (y * xBoardWidth)).xType
            Case E_Empty, E_Fake, E_Floor
                .xUnderColor = BoardCol(x, y)
                .xUnderID = BoardID(x, y)
            Case Else
                .xUnderColor = 0
                .xUnderID = 0
        End Select
        SetBoardID x, y, E_Player, True
        'SetBoardCol x, y, &H1F
        xBoardData(x + (y * xBoardWidth)).xColor = &H1F
        .xX = x + 1
        .xY = y + 1
    End With
    xBoardHeader2.xEnterX = x + 1
    xBoardHeader2.xEnterY = y + 1
End Sub

Public Sub DeleteStat(StatNumber As Long)
    Dim x As Long
    If StatNumber > 0 And StatNumber <= xBoardHeader2.xStatElements Then
        For x = StatNumber To xBoardHeader2.xStatElements - 1
            xBoardStats(x) = xBoardStats(x + 1)
            xOOP(x) = xOOP(x + 1)
        Next x
        xBoardStats(xBoardHeader2.xStatElements) = xEmptyStat
        xOOP(xBoardHeader2.xStatElements).xOOP = ""
        xBoardHeader2.xStatElements = xBoardHeader2.xStatElements - 1
    End If
    
End Sub
Public Function NextFreeStat(x As Byte, y As Byte) As Long
    Dim z As Long
    NextFreeStat = -1
    If x < 0 Or x >= xBoardWidth Or y < 0 Or y >= xBoardHeight Then
        Exit Function
    End If
    For z = 0 To MAXOBJECTS - 1
        With xBoardStats(z)
            'to keep things clean, we need to make sure we replace
            'anything on-screen that matches the coordinates.
            If .xX = 0 Or ((.xX - 1) = x And (.xY - 1) = y) Then
                NextFreeStat = z
                Exit Function
            End If
        End With
    Next z
End Function
Public Function CreateStat(x As Byte, y As Byte) As Long
    Dim StatNumber As Long
    StatNumber = NextFreeStat(x, y)
    If StatNumber >= 0 Then
        With xBoardStats(StatNumber)
            .xX = x + 1
            .xY = y + 1
            .xXStep = 0
            .xYStep = 0
            .xCycle = 1
            .xP1 = 0
            .xP2 = 0
            .xP3 = 0
            .xFollow = 0
            .xLeader = 0
            .xUnderID = 0
            .xUnderColor = 0
            .xPointer = 0
            .xInstruction = 0
            .xLength = 0
        End With
    End If
    If StatNumber > xBoardHeader2.xStatElements Then
        xBoardHeader2.xStatElements = StatNumber
    End If
    CreateStat = StatNumber
End Function

Public Sub LoadBoard(filenumber As Long, ByVal offset As Long)
    Dim xDat As xtiRLE
    Dim xCount As Long
    Dim x As Long
    Dim y As Long
    ReDim xBoardData(0 To TILES_ZZT - 1) As xtiTile
    Get #filenumber, offset, xBoardHeader
    xCount = 0
    xBoardWidth = 60
    xBoardHeight = 25
    Do While xCount < TILES_ZZT
        Get #filenumber, , xDat
        If xDat.xNum = 0 Then 'support ZZT's rollover bug
            y = 256
        Else
            y = xDat.xNum
        End If
        For x = 1 To y
            With xBoardData(xCount)
                .xColor = xDat.xCol
                .xType = xDat.xID
            End With
            xCount = xCount + 1
            If xCount = TILES_ZZT Then
                Exit Do
            End If
        Next x
    Loop
    Get #filenumber, , xBoardHeader2
    ReDim xBoardStats(0 To MAXOBJECTS - 1) As xtiStatElement
    ReDim xOOP(0 To UBound(xBoardStats)) As xtiOOP
    If xBoardHeader2.xStatElements > 150 Or xBoardHeader2.xStatElements < 0 Then
        'this board is corrupted so don't load stats on it
        xBoardHeader2.xStatElements = -1
    Else
        For x = 0 To xBoardHeader2.xStatElements
            Get #filenumber, , xBoardStats(x)
            If xBoardStats(x).xLength > 0 Then
                ReDim xOOP(x).xOOP(0 To xBoardStats(x).xLength - 1) As Byte
                Get #filenumber, , xOOP(x).xOOP
            End If
        Next x
    End If
    
    FixBinds
End Sub

Private Sub FixBinds()
    'this variable y is to count the number of jumps so we don't end up with
    'a clever endless bind... someone out there is bound to try
    'basically what we're doing is simplifying bind chains so objects point
    'directly to the object they are binding
    Dim x As Long
    Dim y As Long
    For x = 0 To xBoardHeader2.xStatElements
        y = 0
        With xBoardStats(x)
            Do While y < 150 And .xLength < 0
                If xBoardStats(-(.xLength)).xLength > 0 Then
                    Exit Do
                End If
                .xLength = xBoardStats(-(.xLength)).xLength
                y = y + 1
            Loop
        End With
    Next x
End Sub

Public Sub LoadSuperBoard(filenumber As Long, ByVal offset As Long)
    Dim xDat As xtiRLE
    Dim xCount As Long
    Dim x As Long
    ReDim xBoardData(0 To TILES_SUPERZZT - 1) As xtiTile
    Get #filenumber, offset, xSuperHeader
    xCount = 0
    xBoardWidth = 96
    xBoardHeight = 80
    xBoardTiles = TILES_SUPERZZT
    Do While xCount < TILES_SUPERZZT
        Get #filenumber, , xDat
        For x = 1 To xDat.xNum
            With xBoardData(xCount)
                .xColor = xDat.xCol
                .xType = SZTtoZZTid(xDat.xID + 0)
            End With
            xCount = xCount + 1
            If xCount = TILES_SUPERZZT Then
                Exit Do
            End If
        Next x
    Loop
    'Debug.Print Loc(filenumber)
    'Debug.Print xSuperHeader2.xshots
    Get #filenumber, , xSuperHeader2
    ReDim xBoardStats(0 To MAXOBJECTS - 1) As xtiStatElement
    ReDim xSuperBoardStats(0 To MAXOBJECTS - 1) As xtiSuperStatElement
    ReDim xOOP(0 To UBound(xSuperBoardStats)) As xtiOOP
    For x = 0 To xSuperHeader2.xStatElements
        'Debug.Print Loc(filenumber)
        Get #filenumber, , xSuperBoardStats(x)
        If xSuperBoardStats(x).xLength > 0 Then
            ReDim xOOP(x).xOOP(0 To xSuperBoardStats(x).xLength - 1) As Byte
            Get #filenumber, , xOOP(x).xOOP
        End If
        With xBoardStats(x)
            .xCycle = xSuperBoardStats(x).xCycle
            .xFollow = xSuperBoardStats(x).xFollow
            .xInstruction = xSuperBoardStats(x).xInstruction
            .xLeader = xSuperBoardStats(x).xLeader
            .xLength = xSuperBoardStats(x).xLength
            .xP1 = xSuperBoardStats(x).xP1
            .xP2 = xSuperBoardStats(x).xP2
            .xP3 = xSuperBoardStats(x).xP3
            .xPointer = xSuperBoardStats(x).xPointer
            .xUnderColor = xSuperBoardStats(x).xUnderColor
            .xUnderID = SZTtoZZTid(xSuperBoardStats(x).xUnderID + 0)
            .xX = xSuperBoardStats(x).xX
            .xXStep = xSuperBoardStats(x).xXStep
            .xY = xSuperBoardStats(x).xY
            .xYStep = xSuperBoardStats(x).xYStep
        End With
    Next x
    xBoardHeader.xSize = xSuperHeader.xSize
    SetBoardName Left$(xSuperHeader.xTitle, xSuperHeader.xTitleSize)
    With xBoardHeader2
        .xdark = 0
        .xEast = xSuperHeader2.xEast
        .xNorth = xSuperHeader2.xNorth
        .xRestart = xSuperHeader2.xRestart
        .xshots = xSuperHeader2.xshots
        .xSouth = xSuperHeader2.xSouth
        .xStatElements = xSuperHeader2.xStatElements
        .xWest = xSuperHeader2.xWest
    End With
    
    FixBinds
End Sub

Public Sub LoadBoardFile(FileName As String)
    Dim f As Long
    f = FreeFile
    Open FileName For Binary As #f
    LoadBoard f, 1
    Close #f
End Sub

Public Sub LoadSuperBoardFile(FileName As String)
    Dim f As Long
    f = FreeFile
    Open FileName For Binary As #f
    LoadSuperBoard f, 1
    Close #f
End Sub

Public Function SaveSuperBoard(filenumber As Long, offset As Long) As Long
    ReDim xData(0) As xtiRLE
    Dim x As Long
    Dim y As Long
    Dim z As Long
    Dim a As Byte
    Dim t As Long
    t = TILES_SUPERZZT
    y = -1
    Do While x < t
        y = y + 1
        If y > 0 Then
            ReDim Preserve xData(0 To y) As xtiRLE
        End If
        With xData(y)
            .xCol = xBoardData(x).xColor
            .xID = xBoardData(x).xType
            a = .xID
            .xID = ZZTtoSZTid(a + 0)
            For z = 1 To 255
                x = x + 1
                If x = t Or z = 255 Then
                    Exit For
                End If
                If xBoardData(x).xColor <> .xCol Or xBoardData(x).xType <> a Then
                    Exit For
                End If
            Next z
            .xNum = z
        End With
    Loop
    With xSuperHeader
        .xTitle = xBoardHeader.xTitle
        .xTitleSize = xBoardHeader.xTitleSize
    End With
    With xSuperHeader2
        .xEast = xBoardHeader2.xEast
        .xNorth = xBoardHeader2.xNorth
        .xRestart = xBoardHeader2.xRestart
        .xshots = xBoardHeader2.xshots
        .xSouth = xBoardHeader2.xSouth
        .xStatElements = xBoardHeader2.xStatElements
        .xWest = xBoardHeader2.xWest
    End With
    
    Put #filenumber, offset, xSuperHeader
    Put #filenumber, , xData
    Put #filenumber, , xSuperHeader2
    
    ReDim xSuperBoardStats(0 To xSuperHeader2.xStatElements) As xtiSuperStatElement
    For z = 0 To xSuperHeader2.xStatElements
        With xSuperBoardStats(z)
            .xCycle = xBoardStats(z).xCycle
            .xFollow = xBoardStats(z).xFollow
            .xInstruction = xBoardStats(z).xInstruction
            .xLeader = xBoardStats(z).xLeader
            .xLength = xBoardStats(z).xLength
            .xP1 = xBoardStats(z).xP1
            .xP2 = xBoardStats(z).xP2
            .xP3 = xBoardStats(z).xP3
            .xPointer = xBoardStats(z).xPointer
            .xUnderColor = xBoardStats(z).xUnderColor
            .xUnderID = ZZTtoSZTid(xBoardStats(z).xUnderID + 0)
            .xX = xBoardStats(z).xX
            .xXStep = xBoardStats(z).xXStep
            .xY = xBoardStats(z).xY
            .xYStep = xBoardStats(z).xYStep
        End With
        Put #filenumber, , xSuperBoardStats(z)
        If xSuperBoardStats(z).xLength > 0 Then
            Put #filenumber, , xOOP(z).xOOP
        End If
    Next z
    With xSuperHeader
        .xSize = (LOF(filenumber) - offset) - 1
    End With
    Put #filenumber, offset, xSuperHeader.xSize
    SaveSuperBoard = xSuperHeader.xSize
End Function

Public Function SaveBoard(filenumber As Long, offset As Long) As Long
    ReDim xData(0) As xtiRLE
    Dim x As Long
    Dim y As Long
    Dim z As Long
    Dim a As Byte
    Dim t As Long
    t = TILES_ZZT
    y = -1
    Do While x < t
        y = y + 1
        If y > 0 Then
            ReDim Preserve xData(0 To y) As xtiRLE
        End If
        With xData(y)
            .xCol = xBoardData(x).xColor
            .xID = xBoardData(x).xType
            a = .xID
            .xID = ClassicZZTElementFilter(a + 0)
            For z = 1 To 255
                x = x + 1
                If x = t Or z = 255 Then
                    Exit For
                End If
                If xBoardData(x).xColor <> .xCol Or xBoardData(x).xType <> a Then
                    Exit For
                End If
            Next z
            .xNum = z
        End With
    Loop
    'xBoardHeader2.xStatElements = 1
    Put #filenumber, offset, xBoardHeader
    Put #filenumber, , xData
    Put #filenumber, , xBoardHeader2
    For z = 0 To xBoardHeader2.xStatElements
        Put #filenumber, , xBoardStats(z)
        If xBoardStats(z).xLength > 0 Then
            Put #filenumber, , xOOP(z).xOOP
        End If
    Next z
    
    xBoardHeader.xSize = (LOF(filenumber) - offset) - 1
    Put #filenumber, offset, xBoardHeader.xSize
    SaveBoard = xBoardHeader.xSize
End Function

Public Sub SaveBoardFile(FileName As String)
    Dim f As Long
    f = FreeFile
    If Dir(FileName) <> "" Then
        Kill FileName
    End If
    Open FileName For Binary As #f
    SaveBoard f, 1
    Close #f
End Sub

Public Sub SaveSuperBoardFile(FileName As String)
    Dim f As Long
    f = FreeFile
    If Dir(FileName) <> "" Then
        Kill FileName
    End If
    Open FileName For Binary As #f
    SaveSuperBoard f, 1
    Close #f
End Sub

Public Sub ResetBoard(ByVal bSuperZZT As Boolean)
    Dim x As Long
    If bSuperZZT Then
        xBoardTiles = TILES_SUPERZZT
        xBoardWidth = 96
        xBoardHeight = 80
    Else
        xBoardTiles = TILES_ZZT
        xBoardWidth = 60
        xBoardHeight = 25
    End If
    ReDim xBoardData(0 To xBoardTiles - 1) As xtiTile
    
    For x = 0 To xBoardTiles - 1
        xBoardData(x).xColor = 0
        xBoardData(x).xType = 0
    Next x
    With xBoardHeader2
        .xEnterX = xBoardWidth \ 2
        .xEnterY = xBoardHeight \ 2
        .xshots = 255
        .xStatElements = 0
    End With
    ReDim xBoardStats(0 To MAXOBJECTS - 1) As xtiStatElement
    With xBoardStats(0)
        .xXStep = 25088
        .xYStep = 1105
        .xCycle = 1
    End With
    MovePlayer xBoardWidth \ 2, xBoardHeight \ 2
    xBoardHeader.xTitle = "Untitled"
    xBoardHeader.xTitleSize = 8
    ReDim xOOP(0 To MAXOBJECTS - 1) As xtiOOP
End Sub

Public Sub GetStatLocation(StatNumber As Long, ByRef xRET As Long, ByRef yRET As Long)
    With xBoardStats(StatNumber)
        xRET = .xX
        yRET = .xY
    End With
End Sub

Public Function StatCount() As Long
    StatCount = xBoardHeader2.xStatElements
End Function

Public Sub GetBoardInfo(ByRef retShots As Byte, ByRef retDark As Byte, ByRef retZap As Byte, ByRef retBoardN As Byte, ByRef retBoardS As Byte, ByRef retBoardW As Byte, ByRef retBoardE As Byte, ByRef retTime As Integer)
    With xBoardHeader2
        retDark = .xdark
        retZap = .xRestart
        retBoardN = .xNorth
        retBoardS = .xSouth
        retBoardW = .xWest
        retBoardE = .xEast
        retTime = .xTimeLimit
        retShots = .xshots
    End With
End Sub

Public Sub SetBoardInfo(retShots As Byte, retDark As Byte, retZap As Byte, retBoardN As Byte, retBoardS As Byte, retBoardW As Byte, retBoardE As Byte, retTime As Integer)
    With xBoardHeader2
        .xdark = retDark
        .xRestart = retZap
        .xNorth = retBoardN
        .xSouth = retBoardS
        .xWest = retBoardW
        .xEast = retBoardE
        .xTimeLimit = retTime
        .xshots = retShots
    End With
End Sub

Public Function ObjectAt(x As Long, y As Long) As Long
    Dim z As Long
    ObjectAt = -1
    For z = 0 To xBoardHeader2.xStatElements
        With xBoardStats(z)
            If (.xX - 1) = x And (.xY - 1) = y Then
                ObjectAt = z
                Exit Function
            End If
        End With
    Next z
End Function

Public Sub EraseObjectAt(x As Long, y As Long)
    Dim z As Long
    Dim a As Long
    Dim b As Long
    z = ObjectAt(x, y)
    If z > 0 Then
        With xBoardStats(z)
            a = .xUnderColor
            b = .xUnderID
        End With
        SetBoardID x, y, b, False
        SetBoardCol x, y, a
    End If
End Sub

Public Sub SetObjectInfo1(ObjectNum As Long, ByRef retXStep As Integer, ByRef retYStep As Integer, ByRef retCycle As Integer, ByRef retP1 As Byte, ByRef retP2 As Byte, ByRef retP3 As Byte)
    With xBoardStats(ObjectNum)
        .xXStep = retXStep
        .xYStep = retYStep
        .xCycle = retCycle
        .xP1 = retP1
        .xP2 = retP2
        .xP3 = retP3
    End With
End Sub

Public Sub SetObjectInfo2(ObjectNum As Long, ByRef retfollow As Integer, ByRef retleader As Integer, ByRef retUnderID As Byte, ByRef retUnderCol As Byte, ByRef retPointer As Long, ByRef retInstruction As Integer, ByRef retLength As Integer)
    With xBoardStats(ObjectNum)
        .xFollow = retfollow
        .xLeader = retleader
        .xUnderID = retUnderID
        .xUnderColor = retUnderCol
        .xPointer = retPointer
        .xInstruction = retInstruction
        If retLength < 0 Then
            'this object now binds something else... so clear the code
            SetObjectOOP ObjectNum, ""
            .xLength = retLength
        End If
    End With
End Sub

Public Sub GetObjectInfo1(ObjectNum As Long, ByRef retXStep As Integer, ByRef retYStep As Integer, ByRef retCycle As Integer, ByRef retP1 As Byte, ByRef retP2 As Byte, ByRef retP3 As Byte)
    With xBoardStats(ObjectNum)
        retXStep = .xXStep
        retYStep = .xYStep
        retCycle = .xCycle
        retP1 = .xP1
        retP2 = .xP2
        retP3 = .xP3
    End With
End Sub

Public Sub GetObjectInfo2(ObjectNum As Long, ByRef retfollow As Integer, ByRef retleader As Integer, ByRef retUnderID As Byte, ByRef retUnderCol As Byte, ByRef retPointer As Long, ByRef retInstruction As Integer, ByRef retLength As Integer)
    With xBoardStats(ObjectNum)
        retfollow = .xFollow
        retleader = .xLeader
        retUnderID = .xUnderID
        retUnderCol = .xUnderColor
        retPointer = .xPointer
        retInstruction = .xInstruction
        retLength = .xLength
    End With
End Sub

Public Sub SetObjectOOP(ObjectNum As Long, ByVal OOPCode As String)
    'convert all the returns to ZZT's form
    OOPCode = Replace$(OOPCode, Chr$(13) + Chr$(10), Chr$(13))
    xOOP(ObjectNum).xOOP() = StrConv(OOPCode, vbFromUnicode)
    xBoardStats(ObjectNum).xLength = UBound(xOOP(ObjectNum).xOOP) + 1
End Sub

Public Sub SetObjectPointer(ObjectNum As Long, NewPointer As Long)
    xBoardStats(ObjectNum).xPointer = NewPointer
End Sub

Public Function ObjectOOP(ObjectNum As Long) As String
    'convert all the returns to windows form
    Dim x As String
    Dim y() As Byte
    If xBoardStats(ObjectNum).xLength <= 0 Then
        ObjectOOP = ""
        Exit Function
    End If
    y() = xOOP(ObjectNum).xOOP()
    x = StrConv(y(), vbUnicode)
    x = Replace$(x, Chr$(13), Chr$(13) + Chr$(10))
    ObjectOOP = x
End Function

Public Function ObjectLength(ObjectNum As Long) As Integer
    ObjectLength = xBoardStats(ObjectNum).xLength
End Function

Public Sub ChangeTiles(fromID As Long, fromColor As Long, toID As Long, toColor As Long)
    Dim x As Long
    For x = 0 To (xBoardTiles - 1)
        If xBoardData(x).xType = fromID Or fromID = 256 Then
            If xBoardData(x).xColor = fromColor Or fromColor = 256 Then
                If toID < 256 Then
                    xBoardData(x).xType = toID
                End If
                If toColor < 256 Then
                    xBoardData(x).xColor = toColor
                End If
            End If
        End If
    Next x
End Sub

Private Sub Class_Initialize()
    ReDim xBoardData(0) As xtiTile
    xBoardHeight = 1
    xBoardWidth = 1
    xBoardTiles = 1
End Sub

Public Function ObjectName(objnumber As Long) As String
    Dim s As String
    If objnumber <= xBoardHeader2.xStatElements Then
        With xBoardStats(objnumber)
            If .xX > 0 And .xY > 0 Then
                If BoardID(.xX - 1, .xY - 1) = E_Object Then
                    s = ObjectOOP(objnumber)
                    If s <> "" Then
                        If Left$(s, 1) <> "@" Then
                            s = ""
                        Else
                            If InStr(s, Chr$(13)) > 0 Then
                                s = Mid$(Left$(s, InStr(s, Chr$(13)) - 1), 2)
                            Else
                                s = Mid$(s, 2)
                            End If
                        End If
                    End If
                End If
            End If
        End With
    End If
    ObjectName = s
End Function

Public Function ObjectCycle(objnumber As Long) As Integer
    ObjectCycle = xBoardStats(objnumber).xCycle
End Function

Public Sub SetObjectLength(objectnumber As Long, newlength As Integer)
    xBoardStats(objectnumber).xLength = newlength
End Sub
