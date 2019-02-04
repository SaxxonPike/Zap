Attribute VB_Name = "modElements"
' ---------------------------------------------------------------------------
' element info library
' ---------------------------------------------------------------------------
'  this module contains all of ZZT's constants: default color, characters,
'  etc.
'
' saxxonpike 2oo7
' ---------------------------------------------------------------------------
Option Explicit

'normal ZZT elements
Public Const E_Empty = 0
Public Const E_BoardEdge = 1
Public Const E_Messenger = 2
Public Const E_Monitor = 3 'ZZT=char 0, SZZT=char 2 (player without keyboard control)
Public Const E_Player = 4
Public Const E_Ammo = 5
Public Const E_Torch = 6 'in SuperZZT this is nonexistant
Public Const E_Gem = 7
Public Const E_Key = 8
Public Const E_Door = 9
Public Const E_Scroll = 10
Public Const E_Passage = 11
Public Const E_Duplicator = 12
Public Const E_Bomb = 13
Public Const E_Energizer = 14
Public Const E_Star = 15 'in SuperZZT use E_SZTStar
Public Const E_Clockwise = 16
Public Const E_Counter = 17
Public Const E_Bullet = 18 'in SuperZZT use E_SZTBullet
Public Const E_Water = 19 'shows as Lava in SuperZZT
Public Const E_Forest = 20
Public Const E_Solid = 21
Public Const E_Normal = 22
Public Const E_Breakable = 23
Public Const E_Boulder = 24
Public Const E_NSslider = 25
Public Const E_EWslider = 26
Public Const E_Fake = 27
Public Const E_Invisible = 28
Public Const E_BlinkWall = 29
Public Const E_Transporter = 30
Public Const E_Line = 31
Public Const E_Ricochet = 32
Public Const E_HBlinkRay = 33 'in SuperZZT use E_SZTHBlinkRay
Public Const E_Bear = 34
Public Const E_Ruffian = 35
Public Const E_Object = 36
Public Const E_Slime = 37
Public Const E_Shark = 38
Public Const E_SpinningGun = 39
Public Const E_Pusher = 40
Public Const E_Lion = 41
Public Const E_Tiger = 42
Public Const E_VBlinkRay = 43 'in SuperZZT use E_SZTVBlinkRay
Public Const E_Head = 44
Public Const E_Segment = 45
Public Const E_46 = 46

'text
Public Const E_BlueText = 47
Public Const E_GreenText = 48
Public Const E_CyanText = 49
Public Const E_RedText = 50
Public Const E_PurpleText = 51
Public Const E_BrownText = 52
Public Const E_WhiteText = 53

'--- superzzt element translation ---
' Note: These are NOT values that are used by SuperZZT. These are values
'       that I have assigned to be used in the Editor. This is because some
'       items retain the same functionality but have different numbered values.
'
Public Const E_Lava = 99
Public Const E_WaterN = 98
Public Const E_WaterS = 97
Public Const E_WaterW = 96
Public Const E_WaterE = 95
Public Const E_Roton = 94
Public Const E_DragonPup = 93
Public Const E_Pairer = 92
Public Const E_Spider = 91
Public Const E_Web = 90
Public Const E_Stone = 89
Public Const E_Floor = 88

'super ZZT "true" element numbers
Public Const E_SZT15 = 15 'old star element UNUSED
Public Const E_SZT18 = 18 'old bullet element UNUSED
Public Const E_SZTLava = 19 'same as E_Water
Public Const E_SZT33 = 33 'old HBlinkRay element UNUSED
Public Const E_SZT43 = 43 'old VBlinkRay element UNUSED
Public Const E_SZTFloor = 47
Public Const E_SZTWaterN = 48
Public Const E_SZTWaterS = 49
Public Const E_SZTWaterW = 50
Public Const E_SZTWaterE = 51
Public Const E_SZT52 = 52
Public Const E_SZT53 = 53
Public Const E_SZT54 = 54
Public Const E_SZT55 = 55
Public Const E_SZT56 = 56
Public Const E_SZT57 = 57
Public Const E_SZT58 = 58
Public Const E_SZTRoton = 59
Public Const E_SZTDragonPup = 60
Public Const E_SZTPairer = 61
Public Const E_SZTSpider = 62
Public Const E_SZTWeb = 63
Public Const E_SZTStone = 64
Public Const E_SZT65 = 65
Public Const E_SZT66 = 66
Public Const E_SZT67 = 67
Public Const E_SZT68 = 68
Public Const E_SZTBullet = 69
Public Const E_SZTHBlinkRay = 70
Public Const E_SZTVBlinkRay = 71
Public Const E_SZTStar = 72
Public Const E_SZTBlueText = 73
Public Const E_SZTGreenText = 74
Public Const E_SZTCyanText = 75
Public Const E_SZTRedText = 76
Public Const E_SZTPurpleText = 77
Public Const E_SZTBrownText = 78
Public Const E_SZTWhiteText = 79

'DefaultChar:
' * Returns the default display character of an element.
Public Function DefaultChar(iElem As Long, Optional iColor As Long, Optional iParam As Long) As Long
    Dim x As Long 'return
    Select Case iElem
    Case E_Empty:               x = &H20
    Case E_BoardEdge:           x = &H45
    Case E_Messenger:           x = 84
    Case E_Monitor:             x = 77
    Case E_Player:              x = 2
    Case E_Ammo:                x = 132
    Case E_Torch:               x = 157
    Case E_Gem:                 x = 4
    Case E_Key:                 x = 12
    Case E_Door:                x = 10
    Case E_Scroll:              x = 232
    Case E_Passage:             x = 240
    Case E_Duplicator:          x = 250
    Case E_Bomb:                x = 11
    Case E_Energizer:           x = 127
    Case E_Star:                x = 47
    Case E_Clockwise:           x = 179
    Case E_Counter:             x = 92
    Case E_Bullet:              x = 248
    Case E_Water:               x = 176
    Case E_Forest:              x = 176
    Case E_Solid:               x = 219
    Case E_Normal:              x = 178
    Case E_Breakable:           x = 177
    Case E_Boulder:             x = 254
    Case E_NSslider:            x = 18
    Case E_EWslider:            x = 29
    Case E_Fake:                x = 178
    Case E_Invisible:           x = 176
    Case E_BlinkWall:           x = 206
    Case E_Transporter:         x = 60
    Case E_Line:                x = LineChar(iParam)
    Case E_Ricochet:            x = 42
    Case E_HBlinkRay:           x = 205
    Case E_Bear:                If World.IsSuperZZT = True Then x = &HEB Else x = 153
    Case E_Ruffian:             x = 5
    Case E_Object:              x = 2
    Case E_Slime:               x = 42
    Case E_Shark:               x = 94
    Case E_SpinningGun:         x = 24
    Case E_Pusher:              x = 31
    Case E_Lion:                x = 234
    Case E_Tiger:               x = 227
    Case E_VBlinkRay:           x = 186
    Case E_Head:                x = 233
    Case E_Segment:             x = 79
    Case E_46:                  x = 63
    Case E_BlueText:            x = iColor
    Case E_GreenText:           x = iColor
    Case E_CyanText:            x = iColor
    Case E_RedText:             x = iColor
    Case E_PurpleText:          x = iColor
    Case E_BrownText:           x = iColor
    Case E_WhiteText:           x = iColor
    
    'superZZT
    Case E_Lava:                x = 111
    Case E_WaterN:              x = 30
    Case E_WaterS:              x = 31
    Case E_WaterW:              x = 17
    Case E_WaterE:              x = 16
    Case E_Roton:               x = &H94
    Case E_DragonPup:           x = &H94
    Case E_Pairer:              x = &HE5
    Case E_Spider:              x = 15
    Case E_Web:                 x = WebChar(iParam)
    Case E_Stone:               x = Int(Rnd * 26) + 65 'random, even in editor
    Case E_Floor:               x = 176
    End Select
    DefaultChar = x
End Function

'DefaultColor:
' * Returns the editor's default display color of an element.
Public Function DefaultColor(iElem As Long, Optional iColor As Long = 0, Optional iTick As Long = 0) As Long
    Dim x As Long 'return
    Select Case iElem
    Case E_Empty:               x = 0
    Case E_BoardEdge:           x = (iColor And &H70) + 15
    Case E_Messenger:           x = (iColor And &H70) + 15
    Case E_Monitor:             x = (iColor And &H70) + 15
    Case E_Player:              x = &H1F
    Case E_Ammo:                x = &H3
    Case E_Torch:               x = &H6
    Case E_Gem:                 x = iColor
    Case E_Key:                 x = iColor
    Case E_Door:                x = ((iColor And &H7) * 16) + 15
    Case E_Scroll:              x = ((iColor + iTick) Mod 7) + 9 'auto change color
    Case E_Passage:             x = ((iColor And &H7) * 16) + 15
    Case E_Duplicator:          x = iColor
    Case E_Bomb:                x = iColor
    Case E_Energizer:           x = 5
    Case E_Star:                x = ((iColor + iTick) Mod 7) + 9 'auto change color
    Case E_Clockwise:           x = iColor
    Case E_Counter:             x = iColor
    Case E_Bullet:              x = 15
    Case E_Water:               x = iColor '&H79
    Case E_Forest:              x = iColor '&H20
    Case E_Solid:               x = iColor
    Case E_Normal:              x = iColor
    Case E_Breakable:           x = iColor
    Case E_Boulder:             x = iColor
    Case E_NSslider:            x = iColor
    Case E_EWslider:            x = iColor
    Case E_Fake:                x = iColor
    Case E_Invisible:           x = iColor
    Case E_BlinkWall:           x = iColor
    Case E_Transporter:         x = iColor
    Case E_Line:                x = iColor
    Case E_Ricochet:            x = 10
    Case E_HBlinkRay:           x = iColor
    Case E_Bear:                If World.IsSuperZZT = True Then x = 2 Else x = 6 'green for SZT, brown for ZZT
    Case E_Ruffian:             x = 13
    Case E_Object:              x = iColor
    Case E_Slime:               x = iColor
    Case E_Shark:               x = 7
    Case E_SpinningGun:         x = iColor
    Case E_Pusher:              x = iColor
    Case E_Lion:                x = 12
    Case E_Tiger:               x = 11
    Case E_VBlinkRay:           x = iColor
    Case E_Head:                x = iColor
    Case E_Segment:             x = iColor
    Case E_46:                  x = iColor
    Case E_BlueText:            x = &H1F
    Case E_GreenText:           x = &H2F
    Case E_CyanText:            x = &H3F
    Case E_RedText:             x = &H4F
    Case E_PurpleText:          x = &H5F
    Case E_BrownText:           x = &H6F
    Case E_WhiteText:           x = &HF
    
    'superZZT
    Case E_Lava:                x = &H4E
    Case E_WaterN:              x = &H19
    Case E_WaterS:              x = &H19
    Case E_WaterW:              x = &H19
    Case E_WaterE:              x = &H19
    Case E_Roton:               x = 13
    Case E_DragonPup:           x = 4
    Case E_Pairer:              x = 1
    Case E_Spider:              x = iColor
    Case E_Web:                 x = iColor
    Case E_Stone:               x = Int(Rnd * 7) + 9
    Case E_Floor:               x = iColor
    End Select
    DefaultColor = x
End Function

'DefaultCycle:
' * Returns the editor's default object cycle
Public Function DefaultCycle(iElem As Long) As Integer
    Dim x As Long 'return
    x = 1
    Select Case iElem
    Case E_Bear:                x = 3
    Case E_BlinkWall:           x = 1
    Case E_Bomb:                x = 6
    Case E_Bullet:              x = 1
    Case E_Clockwise:           x = 3
    Case E_Counter:             x = 2
    Case E_Duplicator:          x = 2
    Case E_Head:                x = 2
    Case E_Lion:                x = 2
    Case E_Messenger:           x = 1
    Case E_Monitor:             x = 1
    Case E_Object:              x = 3
    Case E_Passage:             x = 0
    Case E_Player:              x = 1
    Case E_Pusher:              x = 4
    Case E_Ruffian:             x = 1
    Case E_Scroll:              x = 1
    Case E_Segment:             x = 2
    Case E_Shark:               x = 3
    Case E_Slime:               x = 3
    Case E_SpinningGun:         x = 2
    Case E_Star:                x = 1
    Case E_Tiger:               x = 2
    Case E_Transporter:         x = 2
    'superZZT
    Case E_Roton:               x = 1
    Case E_DragonPup:           x = 2
    Case E_Pairer:              x = 1
    Case E_Spider:              x = 1
    Case E_Stone:               x = 1
    End Select
    DefaultCycle = x
End Function

'DefaultXStep
' * Returns the editor's default X-step
Public Function DefaultXStep(iElem As Long) As Integer
    DefaultXStep = 0
End Function

'DefaultYStep
' * Returns the editor's default Y-step
Public Function DefaultYStep(iElem As Long) As Integer
    Dim x As Long 'return
    x = 0
    Select Case iElem
    Case E_BlinkWall:           x = -1
    Case E_Duplicator:          x = -1
    Case E_Pusher:              x = -1
    Case E_Transporter:         x = -1
    End Select
    DefaultYStep = x
End Function

'DefaultP1
' * Returns the editor's default first parameter value
Public Function DefaultP1(iElem As Long) As Integer
    Dim x As Long 'return
    x = 0
    Select Case iElem
    Case E_Bear:                x = 8
    Case E_BlinkWall:           x = 4
    Case E_Head:                x = 4
    Case E_Lion:                x = 4
    Case E_Object:              x = 1
    Case E_Ruffian:             x = 4
    Case E_Shark:               x = 4
    Case E_SpinningGun:         x = 4
    Case E_Tiger:               x = 4
    'superZZT
    Case E_Roton:               x = 4
    Case E_DragonPup:           x = 4
    Case E_Pairer:              x = 4
    Case E_Spider:              x = 4
    End Select
    DefaultP1 = x
End Function

'DefaultP2
' * Returns the editor's default second parameter value
Public Function DefaultP2(iElem As Long) As Integer
    Dim x As Long 'return
    Select Case iElem
    Case E_BlinkWall:           x = 4
    Case E_Duplicator:          x = 4
    Case E_Head:                x = 4
    Case E_Ruffian:             x = 4
    Case E_Slime:               x = 4
    Case E_SpinningGun:         x = 4
    Case E_Tiger:               x = 4
    'superZZT
    Case E_Roton:               x = 4
    Case E_DragonPup:           x = 4
    End Select
    DefaultP2 = x
End Function

'DefaultP3
' * Returns the editor's default third parameter value
Public Function DefaultP3(iElem As Long) As Integer
    DefaultP3 = 0
End Function

'LineChar
' * Using +1=N, +2=S, +4=W, +8=E (value of 0-15) we return which character to use
'   for E_Line
Public Function LineChar(iParam As Long) As Long
    Dim x As Long
    Select Case iParam
        Case 0: x = 249
        Case 1: x = 208
        Case 2: x = 210
        Case 3: x = 186
        Case 4: x = 181
        Case 5: x = 188
        Case 6: x = 187
        Case 7: x = 185
        Case 8: x = 198
        Case 9: x = 200
        Case 10: x = 201
        Case 11: x = 204
        Case 12: x = 205
        Case 13: x = 202
        Case 14: x = 203
        Case 15: x = 206
    End Select
    LineChar = x
End Function

'WebChar
' * Using +1=N, +2=S, +4=W, +8=E (value of 0-15) we return which character to use
'   for E_Web
Public Function WebChar(iParam As Long) As Long
    Dim x As Long
    Select Case iParam
        Case 0: x = 250
        Case 1: x = 179
        Case 2: x = 179
        Case 3: x = 179
        Case 4: x = 196
        Case 5: x = 217
        Case 6: x = 191
        Case 7: x = 180
        Case 8: x = 196
        Case 9: x = 192
        Case 10: x = 218
        Case 11: x = 195
        Case 12: x = 196
        Case 13: x = 193
        Case 14: x = 194
        Case 15: x = 197
    End Select
    WebChar = x
End Function

'IsLinkedChar
' * Returns whether or not the chosen element changes chars depending on surroundings
Public Function IsLinkedChar(iElem As Long)
    Select Case iElem
        Case E_Web, E_Line
            IsLinkedChar = True
    End Select
End Function

'DefaultP1Name
' * Returns the editor default RANGE,DESCRIPTION for an element
Public Function DefaultP1Name(iElem As Long) As String
    Dim x As String 'return
    x = ""
    Select Case iElem
    Case E_Bear:                x = "0-8,Sensitivity"
    Case E_BlinkWall:           x = "0-8,Start Interval"
    Case E_Bomb:                x = "0-8,Countdown (0=not set)"
    Case E_Bullet:              x = "0-1,Origin (From Enemy)"
    Case E_Clockwise:           x = ""
    Case E_Counter:             x = ""
    Case E_Duplicator:          x = "0-6,Position (6=duplicate)"
    Case E_Head:                x = "0-8,Intelligence"
    Case E_Lion:                x = "0-8,Intellicence"
    Case E_Messenger:           x = ""
    Case E_Monitor:             x = ""
    Case E_Object:              x = "0-255,Object Character"
    Case E_Passage:             x = ""
    Case E_Player:              x = ""
    Case E_Pusher:              x = ""
    Case E_Ruffian:             x = "0-8,Intelligence"
    Case E_Scroll:              x = ""
    Case E_Segment:             x = ""
    Case E_Shark:               x = "0-8,Intelligence"
    Case E_Slime:               x = "0-8,Current Position"
    Case E_SpinningGun:         x = "0-8,Intelligence"
    Case E_Star:                x = "0-1,Origin (From Enemy)"
    Case E_Tiger:               x = "0-8,Intelligence"
    Case E_Transporter:         x = ""
    'superZZT
    Case E_Roton:               x = "0-8,Intelligence"
    Case E_DragonPup:           x = "0-8,Intelligence"
    Case E_Pairer:              x = "0-8,Intelligence"
    Case E_Spider:              x = "0-8,Intelligence"
    End Select
    DefaultP1Name = x
End Function

'DefaultP2Name
' * Returns the editor default RANGE,DESCRIPTION for an element
Public Function DefaultP2Name(iElem As Long) As String
    Dim x As String 'return
    x = ""
    Select Case iElem
    Case E_Bear:                x = ""
    Case E_BlinkWall:           x = "0-8,Fire Interval"
    Case E_Bomb:                x = ""
    Case E_Bullet:              x = ""
    Case E_Clockwise:           x = ""
    Case E_Counter:             x = ""
    Case E_Duplicator:          x = "0-8,Rate"
    Case E_Head:                x = "0-8,Deviance"
    Case E_Lion:                x = ""
    Case E_Messenger:           x = ""
    Case E_Monitor:             x = ""
    Case E_Object:              x = "0-1,#LOCKed? (Unable to receive messages)"
    Case E_Passage:             x = ""
    Case E_Player:              x = ""
    Case E_Pusher:              x = ""
    Case E_Ruffian:             x = "0-8,Resting Time"
    Case E_Scroll:              x = ""
    Case E_Segment:             x = ""
    Case E_Shark:               x = ""
    Case E_Slime:               x = "0-8,Spread Interval"
    Case E_SpinningGun:         x = "0-8,Firing Rate"
    Case E_Star:                x = "0-255,Life Cycles Left"
    Case E_Tiger:               x = "0-8,Firing Rate"
    Case E_Transporter:         x = ""
    'superZZT
    Case E_Roton:               x = "0-8,Switch Rate"
    Case E_DragonPup:           x = "0-8,Switch Rate"
    Case E_Pairer:              x = ""
    Case E_Spider:              x = ""
    End Select
    DefaultP2Name = x
End Function

'DefaultP3Name
' * Returns the editor default RANGE,DESCRIPTION for an element
Public Function DefaultP3Name(iElem As Long) As String
    Dim x As String 'return
    x = ""
    Select Case iElem
    Case E_Bear:                x = ""
    Case E_BlinkWall:           x = "0-8,Current Position"
    Case E_Bomb:                x = ""
    Case E_Bullet:              x = ""
    Case E_Clockwise:           x = ""
    Case E_Counter:             x = ""
    Case E_Duplicator:          x = ""
    Case E_Head:                x = ""
    Case E_Lion:                x = ""
    Case E_Messenger:           x = ""
    Case E_Monitor:             x = ""
    Case E_Object:              x = ""
    Case E_Passage:             x = "DEST,Destination"
    Case E_Player:              x = ""
    Case E_Pusher:              x = ""
    Case E_Ruffian:             x = ""
    Case E_Scroll:              x = ""
    Case E_Segment:             x = ""
    Case E_Shark:               x = ""
    Case E_Slime:               x = ""
    Case E_SpinningGun:         x = "0-1,Firing Type (checked=stars)"
    Case E_Star:                x = ""
    Case E_Tiger:               x = "0-1,Firing Type (checked=stars)"
    Case E_Transporter:         x = ""
    'superZZT
    Case E_Roton:               x = ""
    Case E_DragonPup:           x = ""
    Case E_Pairer:              x = ""
    Case E_Spider:              x = ""
    End Select
    DefaultP3Name = x
End Function

'DefaultStats
' * Returns whether or not an object made in the editor contains stats by default.
Public Function DefaultStats(iElem As Long) As Boolean
    Dim x As Boolean 'return
    x = False
    Select Case iElem
    Case E_Bear:                x = True
    Case E_BlinkWall:           x = True
    Case E_Bomb:                x = True
    Case E_Bullet:              x = True
    Case E_Clockwise:           x = True
    Case E_Counter:             x = True
    Case E_Duplicator:          x = True
    Case E_Head:                x = True
    Case E_Lion:                x = True
    Case E_Messenger:           x = True
    Case E_Monitor:             x = True
    Case E_Object:              x = True
    Case E_Passage:             x = True
    Case E_Player:              x = True
    Case E_Pusher:              x = True
    Case E_Ruffian:             x = True
    Case E_Scroll:              x = True
    Case E_Segment:             x = True
    Case E_Shark:               x = True
    Case E_Slime:               x = True
    Case E_SpinningGun:         x = True
    Case E_Star:                x = True
    Case E_Tiger:               x = True
    Case E_Transporter:         x = True
    'superZZT
    Case E_Roton:               x = True
    Case E_DragonPup:           x = True
    Case E_Pairer:              x = True
    Case E_Spider:              x = True
    Case E_Stone:               x = True
    End Select
    DefaultStats = x
End Function

'SZTtoZZTid
' * Converts a SuperZZT element into a ZZT element number
Public Function SZTtoZZTid(iElem As Long) As Long
    Dim x As Long
    Select Case iElem
    Case E_SZTFloor:            x = E_Floor
    Case E_SZTWaterN:           x = E_WaterN
    Case E_SZTWaterS:           x = E_WaterS
    Case E_SZTWaterW:           x = E_WaterW
    Case E_SZTWaterE:           x = E_WaterE
    Case E_SZTRoton:            x = E_Roton
    Case E_SZTDragonPup:        x = E_DragonPup
    Case E_SZTPairer:           x = E_Pairer
    Case E_SZTSpider:           x = E_Spider
    Case E_SZTWeb:              x = E_Web
    Case E_SZTStone:            x = E_Stone
    Case E_SZTBullet:           x = E_Bullet
    Case E_SZTHBlinkRay:        x = E_HBlinkRay
    Case E_SZTVBlinkRay:        x = E_VBlinkRay
    Case E_SZTStar:             x = E_Star
    Case E_SZTBlueText:         x = E_BlueText
    Case E_SZTGreenText:        x = E_GreenText
    Case E_SZTCyanText:         x = E_CyanText
    Case E_SZTRedText:          x = E_RedText
    Case E_SZTPurpleText:       x = E_PurpleText
    Case E_SZTBrownText:        x = E_BrownText
    Case E_SZTWhiteText:        x = E_WhiteText
    Case E_SZTLava:             x = E_Lava
    Case Else:                  x = iElem
    End Select
    SZTtoZZTid = x
End Function

'SZTtoZZTid
' * Converts a ZZT element into a SuperZZT element number
Public Function ZZTtoSZTid(iElem As Long) As Long
    Dim x As Long
    Select Case iElem
    Case E_Floor:               x = E_SZTFloor
    Case E_WaterN:              x = E_SZTWaterN
    Case E_WaterS:              x = E_SZTWaterS
    Case E_WaterW:              x = E_SZTWaterW
    Case E_WaterE:              x = E_SZTWaterE
    Case E_Roton:               x = E_SZTRoton
    Case E_DragonPup:           x = E_SZTDragonPup
    Case E_Pairer:              x = E_SZTPairer
    Case E_Spider:              x = E_SZTSpider
    Case E_Web:                 x = E_SZTWeb
    Case E_Stone:               x = E_SZTStone
    Case E_Bullet:              x = E_SZTBullet
    Case E_HBlinkRay:           x = E_SZTHBlinkRay
    Case E_VBlinkRay:           x = E_SZTVBlinkRay
    Case E_Star:                x = E_SZTStar
    Case E_BlueText:            x = E_SZTBlueText
    Case E_GreenText:           x = E_SZTGreenText
    Case E_CyanText:            x = E_SZTCyanText
    Case E_RedText:             x = E_SZTRedText
    Case E_PurpleText:          x = E_SZTPurpleText
    Case E_BrownText:           x = E_SZTBrownText
    Case E_WhiteText:           x = E_SZTWhiteText
    Case E_Lava:                x = E_SZTLava
    Case Else:                  x = iElem
    End Select
    ZZTtoSZTid = x
End Function

'ClassicZZTElementFilter
' * Filters out all SuperZZT-exclusive elements.
Public Function ClassicZZTElementFilter(iElem As Long) As Long
    If iElem > E_WhiteText Then
        ClassicZZTElementFilter = 0
    Else
        ClassicZZTElementFilter = iElem
    End If
End Function
