VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsZZTWorld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' **************************************************************
'  clsZZTWorld
'
'  ZZT world class for Visual Basic 6 (uses clsZZTBoard, modMain)
'  SaxxonPike 2oo7-2oo8
' **************************************************************
Option Explicit

Private Type xtiFlag
    xFlagLength As Byte
    xFlagName As String * 20
End Type

Private Type xtiWorld
    xMagic As Integer
    xBoardCount As Integer
    xAmmo As Integer
    xGems As Integer
    xKeys(0 To 6) As Byte
    xHealth As Integer
    xStartBoard As Integer
    xTorches As Integer
    xTorchCycles As Integer
    xEnergyCycles As Integer
    xUnk0 As Integer
    xScore As Integer
    xGameNameLength As Byte
    xGameName As String * 20
    xFlags(0 To 9) As xtiFlag
    xTimePassed As Integer
    xUnk1 As Integer
    xLocked As Byte
End Type

Private Type xtiSuperWorld
    xMagic As Integer
    xBoardCount As Integer
    xAmmo As Integer
    xGems As Integer
    xKeys(0 To 6) As Byte
    xHealth As Integer
    xStartBoard As Integer
    xUnk2 As Integer
    xUnk3 As Integer
    xUnk4 As Integer
    xScore As Integer
    xGameNameLength As Byte
    xGameName As String * 20
    xFlags(0 To 15) As xtiFlag
    xUnk0(0 To 1) As Integer
    xLocked As Byte
    xStones As Integer
    xUnk5(0 To 10) As Byte
    xTimeLeft As Integer
End Type

Private xBoards() As New clsZZTBoard
Private xWorld As xtiWorld
Private xSuperWorld As xtiSuperWorld
Private xUnsavedChanges As Boolean
Private xBoardWidth As Long
Private xBoardHeight As Long
Private xBoardTiles As Long
Private xSuperZZTMode As Boolean

Public Sub SetWorldInfo1(xAmmo As Integer, xGems As Integer, xKeys() As Byte, xHealth As Integer, xTorches As Integer, xTorchCycles As Integer, xEnergyCycles As Integer)
    With xWorld
        .xAmmo = xAmmo
        .xGems = xGems
        .xKeys(0) = xKeys(0)
        .xKeys(1) = xKeys(1)
        .xKeys(2) = xKeys(2)
        .xKeys(3) = xKeys(3)
        .xKeys(4) = xKeys(4)
        .xKeys(5) = xKeys(5)
        .xKeys(6) = xKeys(6)
        .xHealth = xHealth
        .xTorches = xTorches
        .xTorchCycles = xTorchCycles
        .xEnergyCycles = xEnergyCycles
    End With
End Sub

Public Sub SetWorldInfo2(xScore As Integer, xFlagNames() As String, xTimePassed As Integer, xLocked As Byte)
    Dim x As Long
    With xWorld
        .xScore = xScore
        For x = 0 To 9
            .xFlags(x).xFlagName = xFlagNames(x)
            .xFlags(x).xFlagLength = Len(xFlagNames(x))
        Next x
        .xTimePassed = xTimePassed
        .xLocked = xLocked
    End With
    With xSuperWorld
        For x = 10 To 15
            .xFlags(x).xFlagName = xFlagNames(x)
            .xFlags(x).xFlagLength = Len(xFlagNames(x))
        Next x
    End With
End Sub

Public Sub GetWorldInfo1(ByRef xAmmo As Integer, ByRef xGems As Integer, ByRef xKeys() As Byte, ByRef xHealth As Integer, ByRef xTorches As Integer, ByRef xTorchCycles As Integer, ByRef xEnergyCycles As Integer)
    With xWorld
        xAmmo = .xAmmo
        xGems = .xGems
        xKeys(0) = .xKeys(0)
        xKeys(1) = .xKeys(1)
        xKeys(2) = .xKeys(2)
        xKeys(3) = .xKeys(3)
        xKeys(4) = .xKeys(4)
        xKeys(5) = .xKeys(5)
        xKeys(6) = .xKeys(6)
        xHealth = .xHealth
        xTorches = .xTorches
        xTorchCycles = .xTorchCycles
        xEnergyCycles = .xEnergyCycles
    End With
End Sub

Public Sub GetWorldInfo2(ByRef xScore As Integer, ByRef xFlagNames() As String, ByRef xTimePassed As Integer, ByRef xLocked As Byte)
    Dim x As Long
    With xWorld
        xScore = .xScore
        For x = 0 To 9
            xFlagNames(x) = Left$(.xFlags(x).xFlagName, .xFlags(x).xFlagLength)
        Next x
        xTimePassed = .xTimePassed
        xLocked = .xLocked
    End With
    With xSuperWorld
        For x = 10 To 15
            xFlagNames(x) = Left$(.xFlags(x).xFlagName, .xFlags(x).xFlagLength)
        Next x
    End With
End Sub

Public Sub NewWorld(bSuperZZT As Boolean)
    Dim x As Long
    For x = 0 To 255
        Set xBoards(x) = Nothing
    Next x
    ReDim xBoards(0 To 255) As New clsZZTBoard
    For x = 0 To 1
        xBoards(x).ResetBoard bSuperZZT
    Next x
    SetWorldName "Untitled"
    xWorld.xHealth = 100
    xWorld.xAmmo = 0
    xWorld.xEnergyCycles = 0
    xWorld.xGems = 0
    xWorld.xKeys(0) = 0
    xWorld.xKeys(1) = 0
    xWorld.xKeys(2) = 0
    xWorld.xKeys(3) = 0
    xWorld.xKeys(4) = 0
    xWorld.xKeys(5) = 0
    xWorld.xKeys(6) = 0
    xWorld.xLocked = 0
    xWorld.xScore = 0
    xWorld.xStartBoard = 0
    xWorld.xTimePassed = 0
    xWorld.xTorches = 0
    xWorld.xUnk0 = 0
    xWorld.xUnk1 = 0
    If bSuperZZT Then
        xWorld.xMagic = -2
        xBoardHeight = 80
        xBoardWidth = 96
        xBoardTiles = 7680
        xWorld.xTorchCycles = -1
    Else
        xWorld.xMagic = -1
        xBoardHeight = 25
        xBoardWidth = 60
        xBoardTiles = 1500
        xWorld.xTorchCycles = 0
    End If
    xWorld.xBoardCount = 1
    xBoards(0).SetBoardName "Title Screen"
    xSuperZZTMode = bSuperZZT
End Sub

Public Sub MovePlayer(boardnumber As Long, x As Long, y As Long)
    xBoards(boardnumber).MovePlayer x, y
End Sub

Public Function BoardCount() As Long
    BoardCount = xWorld.xBoardCount
End Function

Public Function CreateNewBoard() As Long
    xWorld.xBoardCount = xWorld.xBoardCount + 1
    xBoards(xWorld.xBoardCount).ResetBoard xSuperZZTMode
    CreateNewBoard = xWorld.xBoardCount
End Function

Public Function StatCount(boardnumber As Long) As Long
    StatCount = xBoards(boardnumber).StatCount
End Function

Public Function BoardName(boardnumber As Long) As String
    BoardName = xBoards(boardnumber).BoardName
End Function

Public Function WorldName() As String
    WorldName = Left$(xWorld.xGameName, xWorld.xGameNameLength)
End Function

Public Sub SetWorldName(newname As String)
    xWorld.xGameName = Trim$(newname) + String(20, 0)
    xWorld.xGameNameLength = Len(newname)
    If xWorld.xGameNameLength > 20 Then xWorld.xGameNameLength = 20
End Sub

Public Function BoardID(boardnumber As Long, x As Long, y As Long) As Byte
    If x >= 0 And x < xBoardWidth And y >= 0 And y < xBoardHeight Then
        BoardID = xBoards(boardnumber).BoardID(x, y)
    End If
End Function

Public Function BoardCol(boardnumber As Long, x As Long, y As Long) As Byte
    If x >= 0 And x < xBoardWidth And y >= 0 And y < xBoardHeight Then
        BoardCol = xBoards(boardnumber).BoardCol(x, y)
    End If
End Function

Public Sub SetBoardID(boardnumber As Long, x As Long, y As Long, i As Long)
    If x >= 0 And x < xBoardWidth And y >= 0 And y < xBoardHeight Then
        xBoards(boardnumber).SetBoardID x, y, i, (frmEdit.chkStats.Value = 1)
    End If
    xUnsavedChanges = True
End Sub

Public Sub SetBoardCol(boardnumber As Long, x As Long, y As Long, i As Long)
    If x >= 0 And x < xBoardWidth And y >= 0 And y < xBoardHeight Then
        xBoards(boardnumber).SetBoardCol x, y, i + 0
    End If
    xUnsavedChanges = True
End Sub

Public Sub GetStatLocation(boardnumber As Long, StatNumber As Long, ByRef xRET As Long, ByRef yRET As Long)
    xBoards(boardnumber).GetStatLocation StatNumber, xRET, yRET
End Sub

Public Sub SaveBoard(boardnumber As Long, FileName As String)
    If xSuperZZTMode Then
        xBoards(boardnumber).SaveSuperBoardFile FileName
    Else
        xBoards(boardnumber).SaveBoardFile FileName
    End If
End Sub

Public Sub LoadBoard(boardnumber As Long, FileName As String)
    If xSuperZZTMode Then
        xBoards(boardnumber).LoadSuperBoardFile FileName
    Else
        xBoards(boardnumber).LoadBoardFile FileName
    End If
    xUnsavedChanges = True
End Sub

Public Sub LoadWorld(FileName As String)
    Dim f As Long
    Dim x As Long
    Dim z As Long
    Dim y As Integer
    f = FreeFile
    Open FileName For Binary As #f
    Get #f, 1, y
    If y <> -1 And y <> -2 Then
        If MsgBox("This world doesn't have an identifiable type. Load anyway?", vbYesNo) = vbNo Then
            NewWorld False
            Exit Sub
        End If
    End If
    If y = -2 Then
        NewWorld True
    Else
        NewWorld False
    End If

    If xSuperZZTMode = False Then
        Get #f, 1, xWorld
        z = 513
    Else
        Get #f, 1, xSuperWorld
        ConvertFromSuperWorld
        z = 1025
    End If
    For x = 0 To xWorld.xBoardCount
        Get #f, z, y
        If xSuperZZTMode = False Then
            xBoards(x).LoadBoard f, z
        Else
            xBoards(x).LoadSuperBoard f, z
        End If
        z = z + y + 2
    Next x
    SetCurrentBoard xWorld.xStartBoard + 0
    Close #f
    xUnsavedChanges = False
End Sub

Public Sub SetStartBoard(newboard As Long)
    xWorld.xStartBoard = newboard
End Sub

Public Function StartBoard() As Long
    StartBoard = xWorld.xStartBoard
End Function

Public Sub SaveWorld(FileName As String, Optional bForTest As Boolean = False)
    Dim f As Long
    Dim x As Long
    Dim z As Long
    Dim y As Long
    Dim s As String
    Dim pw As String
    Dim tg(0 To 2) As Byte
    f = FreeFile
    If FileName = "" Then
        Exit Sub
    End If
    If Dir(FileName) <> "" Then
        Kill FileName
    End If
    If InStr(FileName, "\") > 0 Then
        s = Mid$(FileName, InStrRev(FileName, "\") + 1)
        If InStr(s, ".") Then
            s = Left$(s, InStrRev(s, ".") - 1)
        End If
    End If
    
    pw = WorldName
    SetWorldName s
    
    Open FileName For Binary As #f
    If xSuperZZTMode Then
        ConvertToSuperWorld
        Put #f, 1, xSuperWorld
        z = 1025
    Else
        Put #f, 1, xWorld
        z = 513
    End If
    For x = 0 To xWorld.xBoardCount
        If xSuperZZTMode Then
            y = xBoards(x).SaveSuperBoard(f, z)
        Else
            y = xBoards(x).SaveBoard(f, z)
        End If
        z = z + y + 2
    Next x
    tg(0) = 90
    tg(1) = 65
    tg(2) = 80
    If xSuperZZTMode Then
        Put #f, 1022, tg 'zap tag
    Else
        Put #f, 510, tg
    End If
    Close #f
    xUnsavedChanges = False
    If bForTest Then
        SetWorldName pw
    End If
End Sub

Private Sub ConvertFromSuperWorld()
    With xWorld
        .xAmmo = xSuperWorld.xAmmo
        .xBoardCount = xSuperWorld.xBoardCount
        .xFlags(0) = xSuperWorld.xFlags(0)
        .xFlags(1) = xSuperWorld.xFlags(1)
        .xFlags(2) = xSuperWorld.xFlags(2)
        .xFlags(3) = xSuperWorld.xFlags(3)
        .xFlags(4) = xSuperWorld.xFlags(4)
        .xFlags(5) = xSuperWorld.xFlags(5)
        .xFlags(6) = xSuperWorld.xFlags(6)
        .xFlags(7) = xSuperWorld.xFlags(7)
        .xFlags(8) = xSuperWorld.xFlags(8)
        .xFlags(9) = xSuperWorld.xFlags(9)
        .xGameName = xSuperWorld.xGameName
        .xGameNameLength = xSuperWorld.xGameNameLength
        .xGems = xSuperWorld.xGems
        .xHealth = xSuperWorld.xHealth
        .xKeys(0) = xSuperWorld.xKeys(0)
        .xKeys(1) = xSuperWorld.xKeys(1)
        .xKeys(2) = xSuperWorld.xKeys(2)
        .xKeys(3) = xSuperWorld.xKeys(3)
        .xKeys(4) = xSuperWorld.xKeys(4)
        .xKeys(5) = xSuperWorld.xKeys(5)
        .xKeys(6) = xSuperWorld.xKeys(6)
        .xLocked = xSuperWorld.xLocked
        .xMagic = xSuperWorld.xMagic
        .xScore = xSuperWorld.xScore
        .xStartBoard = xSuperWorld.xStartBoard
        .xTimePassed = xSuperWorld.xTimeLeft
        .xTorchCycles = xSuperWorld.xStones
        .xTorches = 0
        .xUnk0 = 0
        .xUnk1 = 0
    End With
End Sub

Private Sub ConvertToSuperWorld()
    With xSuperWorld
        .xAmmo = xWorld.xAmmo
        .xBoardCount = xWorld.xBoardCount
        .xFlags(0) = xWorld.xFlags(0)
        .xFlags(1) = xWorld.xFlags(1)
        .xFlags(2) = xWorld.xFlags(2)
        .xFlags(3) = xWorld.xFlags(3)
        .xFlags(4) = xWorld.xFlags(4)
        .xFlags(5) = xWorld.xFlags(5)
        .xFlags(6) = xWorld.xFlags(6)
        .xFlags(7) = xWorld.xFlags(7)
        .xFlags(8) = xWorld.xFlags(8)
        .xFlags(9) = xWorld.xFlags(9)
        .xGameName = xWorld.xGameName
        .xGameNameLength = xWorld.xGameNameLength
        .xGems = xWorld.xGems
        .xHealth = xWorld.xHealth
        .xKeys(0) = xWorld.xKeys(0)
        .xKeys(1) = xWorld.xKeys(1)
        .xKeys(2) = xWorld.xKeys(2)
        .xKeys(3) = xWorld.xKeys(3)
        .xKeys(4) = xWorld.xKeys(4)
        .xKeys(5) = xWorld.xKeys(5)
        .xKeys(6) = xWorld.xKeys(6)
        .xLocked = xWorld.xLocked
        .xMagic = xWorld.xMagic
        .xScore = xWorld.xScore
        .xStartBoard = xWorld.xStartBoard
        .xStones = xWorld.xTorchCycles
        .xTimeLeft = xWorld.xTimePassed
    End With
End Sub
Public Function ObjectCharAt(boardnumber As Long, x As Long, y As Long, anim As Long) As Byte
    ObjectCharAt = xBoards(boardnumber).ObjectCharAt(x, y, anim)
End Function

Private Sub Class_Initialize()
    ReDim xBoards(0 To 255) As New clsZZTBoard
End Sub

Public Sub GetBoardInfo(boardnumber As Long, ByRef retShots As Byte, ByRef retDark As Byte, ByRef retZap As Byte, ByRef retBoardN As Byte, ByRef retBoardS As Byte, ByRef retBoardW As Byte, ByRef retBoardE As Byte, ByRef retTime As Integer)
    xBoards(boardnumber).GetBoardInfo retShots, retDark, retZap, retBoardN, retBoardS, retBoardW, retBoardE, retTime
End Sub

Public Sub SetBoardName(boardnumber As Long, bn As String)
    xBoards(boardnumber).SetBoardName bn
End Sub

Public Sub SetBoardInfo(boardnumber As Long, retShots As Byte, retDark As Byte, retZap As Byte, retBoardN As Byte, retBoardS As Byte, retBoardW As Byte, retBoardE As Byte, retTime As Integer)
    xBoards(boardnumber).SetBoardInfo retShots, retDark, retZap, retBoardN, retBoardS, retBoardW, retBoardE, retTime
End Sub

Public Sub GetObjectInfo1(boardnumber As Long, ObjectNum As Long, ByRef retXStep As Integer, ByRef retYStep As Integer, ByRef retCycle As Integer, ByRef retP1 As Byte, ByRef retP2 As Byte, ByRef retP3 As Byte)
    xBoards(boardnumber).GetObjectInfo1 ObjectNum, retXStep, retYStep, retCycle, retP1, retP2, retP3
End Sub

Public Sub GetObjectInfo2(boardnumber As Long, ObjectNum As Long, ByRef retfollow As Integer, ByRef retleader As Integer, ByRef retUnderID As Byte, ByRef retUnderCol As Byte, ByRef retPointer As Long, ByRef retInstruction As Integer, ByRef retLength As Integer)
    xBoards(boardnumber).GetObjectInfo2 ObjectNum, retfollow, retleader, retUnderID, retUnderCol, retPointer, retInstruction, retLength
End Sub

Public Sub SetObjectInfo1(boardnumber As Long, ObjectNum As Long, ByRef retXStep As Integer, ByRef retYStep As Integer, ByRef retCycle As Integer, ByRef retP1 As Byte, ByRef retP2 As Byte, ByRef retP3 As Byte)
    xBoards(boardnumber).SetObjectInfo1 ObjectNum, retXStep, retYStep, retCycle, retP1, retP2, retP3
End Sub

Public Sub SetObjectInfo2(boardnumber As Long, ObjectNum As Long, ByRef retfollow As Integer, ByRef retleader As Integer, ByRef retUnderID As Byte, ByRef retUnderCol As Byte, ByRef retPointer As Long, ByRef retInstruction As Integer, ByRef retLength As Integer)
    xBoards(boardnumber).SetObjectInfo2 ObjectNum, retfollow, retleader, retUnderID, retUnderCol, retPointer, retInstruction, retLength
End Sub

Public Sub SetObjectPointer(boardnumber As Long, ObjectNum As Long, NewPointer As Long)
    xBoards(boardnumber).SetObjectPointer ObjectNum, NewPointer
End Sub

Public Function CreateStat(boardnumber As Long, x As Byte, y As Byte, id As Long, Optional StatNumber As Long = -2, Optional underID As Long = 0, Optional underColor As Long = 0, Optional Instruction As Long = 0, Optional OOPCode As String = "") As Long
    CreateStat = xBoards(boardnumber).CreateStat(x, y)
End Function

Public Function NextFreeStat(boardnumber As Long, x As Byte, y As Byte) As Long
    NextFreeStat = xBoards(boardnumber).NextFreeStat(x, y)
End Function

Public Function ObjectAt(boardnumber As Long, x As Long, y As Long) As Long
    ObjectAt = xBoards(boardnumber).ObjectAt(x, y)
End Function

Public Sub EraseObjectAt(boardnumber As Long, x As Long, y As Long)
    xBoards(boardnumber).EraseObjectAt x, y
End Sub

Public Function ObjectOOP(boardnumber As Long, ObjectNum As Long) As String
    ObjectOOP = xBoards(boardnumber).ObjectOOP(ObjectNum)
End Function

Public Sub SetObjectOOP(boardnumber As Long, ObjectNum As Long, OOPCode As String)
    xBoards(boardnumber).SetObjectOOP ObjectNum, OOPCode
End Sub

Public Sub ResetBoard(boardnumber As Long)
    xBoards(boardnumber).ResetBoard xSuperZZTMode
End Sub

Public Function UnsavedChanges() As Boolean
    UnsavedChanges = xUnsavedChanges
End Function

Public Sub ChangeAllTiles(fromID As Long, fromColor As Long, toID As Long, toColor As Long)
    Dim x As Long
    For x = 0 To xWorld.xBoardCount
        ChangeTiles x, fromID, fromColor, toID, toColor
    Next x
End Sub

Public Sub ChangeTiles(boardnumber As Long, fromID As Long, fromColor As Long, toID As Long, toColor As Long)
    xBoards(boardnumber).ChangeTiles fromID, fromColor, toID, toColor
End Sub

Public Function IsSuperZZT() As Boolean
    IsSuperZZT = xSuperZZTMode
End Function

Public Function ObjectName(boardnumber As Long, objectnumber As Long) As String
    ObjectName = xBoards(boardnumber).ObjectName(objectnumber)
End Function

Public Function ObjectLength(boardnumber As Long, objectnumber As Long) As Integer
    ObjectLength = xBoards(boardnumber).ObjectLength(objectnumber)
End Function

Public Sub SetObjectLength(boardnumber As Long, objectnumber As Long, newlength As Integer)
    xBoards(boardnumber).SetObjectLength objectnumber, newlength
End Sub

Public Function ObjectCycle(boardnumber As Long, objectnumber As Long) As Integer
    ObjectCycle = xBoards(boardnumber).ObjectCycle(objectnumber)
End Function

Public Sub DeleteBoard(boardnumber As Long)
    Dim x As Long
    Dim y As Long
    Dim z As Long
    Dim locx As Long
    Dim locy As Long
    Dim p1 As Byte
    Dim p2 As Byte
    Dim p3 As Byte
    Dim xs As Integer
    Dim ys As Integer
    Dim cy As Integer
    Dim sh As Byte
    Dim tm As Integer
    Dim dk As Byte
    Dim bNo As Byte
    Dim bSo As Byte
    Dim bEa As Byte
    Dim bWe As Byte
    Dim za As Byte
    Dim f As Boolean
    Dim bc As Byte
    'decrement board edges and passages above this board#
    If xWorld.xStartBoard >= boardnumber Then
        xWorld.xStartBoard = xWorld.xStartBoard - 1
    End If
    
    If boardnumber > 0 And boardnumber <= xWorld.xBoardCount Then
        For x = 0 To xWorld.xBoardCount
            xBoards(x).GetBoardInfo sh, dk, za, bNo, bSo, bWe, bEa, tm
            If bNo >= boardnumber Then bNo = bNo - 1
            If bSo >= boardnumber Then bSo = bSo - 1
            If bEa >= boardnumber Then bEa = bEa - 1
            If bWe >= boardnumber Then bWe = bWe - 1
            xBoards(x).SetBoardInfo sh, dk, za, bNo, bSo, bWe, bEa, tm
            For y = 0 To xBoards(x).StatCount
                xBoards(x).GetStatLocation y, locx, locy
                If locx > 0 And locy > 0 Then
                    If xBoards(x).BoardID(locx - 1, locy - 1) = E_Passage Then
                        xBoards(x).GetObjectInfo1 y, xs, ys, cy, p1, p2, p3
                        If p3 >= boardnumber Then
                            p3 = p3 - 1
                            xBoards(x).SetObjectInfo1 y, xs, ys, cy, p1, p2, p3
                        End If
                    End If
                End If
            Next y
        Next x
        For x = boardnumber To xWorld.xBoardCount - 1
            Set xBoards(x) = xBoards(x + 1)
        Next x
        Set xBoards(xWorld.xBoardCount) = Nothing
        Set xBoards(xWorld.xBoardCount) = New clsZZTBoard
        xWorld.xBoardCount = xWorld.xBoardCount - 1
    End If
End Sub
