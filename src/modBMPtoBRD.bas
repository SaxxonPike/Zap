Attribute VB_Name = "modBMPtoBRD"
' conversion of 60x25 bitmaps to boards
' saxxonpike 2oo7
'
' what this does is creates a lookup table for colors and their appropriate
' fades. then it takes the input image, pixel by pixel, and finds the color
' that is the least different in the table. the lower-right is ignored as it
' is where we will be sticking the player.
'
' 8 background colors + 16 foreground colors + 4 shades (water/breakable/
' normal/solid) = 512 combinations
'
' element           opacity
' -------           -------
' water             25% (FG*.25)+(BG*.75)
' breakable         50% (FG*.50)+(BG*.50)
' normal            75% (FG*.75)+(BG*.25)
' solid             100% (FG)
'
' when loading pictures, we will use LoadPicture and attach a DC to the bitmap
' then retrieve all pixels using GetPixel. then after we are finished, we will
' delete the DC and the picture objects.

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

Private Type xtiColorMatch
    xID As Byte
    xCol As Byte
    xRed As Byte
    xGreen As Byte
    xBlue As Byte
    xUseColor As Boolean
End Type

Private ColorMatch(0 To 511) As xtiColorMatch

Public Sub BMP2BRD_CreateLookupTable(Pal() As RGB)
    
    Dim x As Long '4 counters
    Dim y As Long
    Dim z As Long
    Dim a As Long
    
    Dim r1 As Byte 'foreground rgb
    Dim g1 As Byte
    Dim b1 As Byte
    Dim r2 As Byte 'background rgb
    Dim g2 As Byte
    Dim b2 As Byte
    
    Dim c As Single 'rgb multipliers
    Dim d As Single 'c=foreground, d=background
    
    Debug.Print "Initializing BMPtoBRD color table..."
    
    a = 0
    For x = 0 To 7 'background
        For y = 0 To 15 'foreground
            For z = 0 To 3 'types
                With ColorMatch(a)
                    Select Case z
                        Case 0
                            .xID = E_Water
                            c = 0.25
                        Case 1
                            .xID = E_Breakable
                            c = 0.5
                        Case 2
                            .xID = E_Normal
                            c = 0.75
                        Case 3
                            .xID = E_Solid
                            c = 1
                    End Select
                    
                    .xUseColor = True
                    
                    'remove multiple solids (bg color 0 only)
                    If z = 3 And x <> 0 Then
                        .xUseColor = False
                    End If
                    
                    'remove other multiples
                    If x > y Then
                        .xUseColor = False
                    End If
                    
                    'no brights on black except solids
                    If (x = 0 And (y > 8 Or y = 7)) And z <> 3 Then
                        .xUseColor = False
                    End If
                    
                    'remove retarded color opposites
                    'blue/yellow
                    If x = 1 And y = 14 Then
                        .xUseColor = False
                    End If
                    'green/purple
                    If x = 2 And (y = 13 Or y = 5) Then
                        .xUseColor = False
                    End If
                    'cyan/red
                    If x = 3 And (y = 12 Or y = 4) Then
                        .xUseColor = False
                    End If
                    'red/cyan
                    If x = 4 And (y = 11 Or y = 3) Then
                        .xUseColor = False
                    End If
                    'purple/green
                    If x = 5 And (y = 10 Or y = 2) Then
                        .xUseColor = False
                    End If
                    'brown/blue
                    If x = 6 And y = 9 Then
                        .xUseColor = False
                    End If
                    'if foreground and background colors match, remove
                    'them EXCEPT black solid
                    If (x = y) And (Not (x = 0 And z = 3)) Then
                        .xUseColor = False
                    End If
                    
                    
                    d = 1 - c
                    r1 = Pal(y).r
                    g1 = Pal(y).g
                    b1 = Pal(y).b
                    r2 = Pal(x).r
                    g2 = Pal(x).g
                    b2 = Pal(x).b
                    .xRed = (r1 * c) + (r2 * d)
                    .xGreen = (g1 * c) + (g2 * d)
                    .xBlue = (b1 * c) + (b2 * d)
                    .xCol = y + (x * 16)
                End With
                a = a + 1
            Next z
        Next y
    Next x

End Sub

Private Function FindClosestColor(r As Long, g As Long, b As Long) As Long
    
    Dim s As Long 'selected color
    Dim xClosest(0 To 1, 0 To 1) As Double
    Dim pc As Boolean
    Dim d As Double 'calculated difference
    Dim x As Long 'loop counter
    Dim y As Long
    Dim z As Long
    
    For x = 0 To UBound(xClosest)
        xClosest(x, 1) = 250000
    Next x
    
    e = 250000 'maximum difference
    
    For x = 0 To 511
        d = 0
        d = d + (((CDbl(r) - CDbl(ColorMatch(x).xRed)) ^ 2) * 1)
        d = d + (((CDbl(g) - CDbl(ColorMatch(x).xGreen)) ^ 2) * 1)
        d = d + (((CDbl(b) - CDbl(ColorMatch(x).xBlue)) ^ 2) * 1)
        d = d ^ 0.5
        
        pc = True
        
        'preserve greys by using only greys
        If Abs(r - g) < 3 And Abs(g - b) < 3 Then
            If ColorMatch(x).xRed <> ColorMatch(x).xGreen Or ColorMatch(x).xGreen <> ColorMatch(x).xBlue Then
                pc = False
            End If
        End If
        
        'add the color to the "closest" color list
        If ColorMatch(x).xUseColor = True And pc = True Then
            For y = 0 To UBound(xClosest)
                If d < xClosest(y, 1) Then
                    For z = UBound(xClosest) To (y + 1) Step -1
                        xClosest(z, 0) = xClosest(z - 1, 0)
                        xClosest(z, 1) = xClosest(z - 1, 1)
                    Next z
                    xClosest(y, 0) = x
                    xClosest(y, 1) = d
                    Exit For
                End If
            Next y
        End If
    Next x
    
    'now randomly pick from that small closest color list
    FindClosestColor = xClosest(Int(Rnd * (UBound(xClosest) + 1)), 0)
    
End Function

Public Function BMP2BRD_LoadToBoard(boardnum As Long, bmpfile As String)
    Dim iPict As IPictureDisp
    Dim iWidth As Long
    Dim iHeight As Long
    Dim iDC As Long
    Dim iObj As Long
    Dim iPix As Long
    Dim x As Long
    Dim y As Long
    Dim c As Long
    Dim d As Long
    
    Dim r As Byte
    Dim g As Byte
    Dim b As Byte
    
    Set frmEdit.Picture = LoadPicture(bmpfile)
    
    If frmEdit.Picture Then
        With frmEdit.Picture
            iWidth = CLng((.Width / 2540 * 1440) / Screen.TwipsPerPixelX)
            iHeight = CLng((.Height / 2540 * 1440) / Screen.TwipsPerPixelY)
        End With
    End If
    If iWidth <> 60 Or (iHeight <> 25 And iHeight <> 50) Then
        MsgBox "The bitmap provided must be 60x25.", vbCritical, "Cannot convert"
        Set iPict = Nothing
        Exit Function
    End If
    'iDC = CreateCompatibleDC(frmEdit.hdc)
    
    iDC = frmEdit.hdc
    
    'iObj = SelectObject(iDC, iPict.Handle)
    'If iObj <> 0 Then
        'DeleteObject iObj
    'End If
    World.MovePlayer boardnum, 59, 24
    For y = 0 To 24
        For x = 0 To 59
            If x <> 59 Or y <> 24 Then
                If iHeight = 25 Then
                    d = GetPixel(iDC, x, y)
                Else
                    d = GetPixel(iDC, x, y * 2)
                End If
                r = (d And &HFF&) \ &H1
                g = (d And &HFF00&) \ &H100
                b = (d And &HFF0000) \ &H10000
                'add noise
                'If r < 240 Then r = r + Int(Rnd * 16)
                'If g < 240 Then g = g + Int(Rnd * 16)
                'If b < 240 Then b = b + Int(Rnd * 16)
                c = FindClosestColor(r + 0, g + 0, b + 0)
                With ColorMatch(c)
                    World.SetBoardID CurrentBoard, x, y, .xID + 0
                    World.SetBoardCol CurrentBoard, x, y, .xCol + 0
                End With
            End If
        Next x
    Next y
    Set frmEdit.Picture = Nothing
End Function

Private Function ColorDistance2(Col1 As RGB, Col2 As RGB) As Long
    Dim r As Double
    Dim g As Double
    Dim b As Double
    Dim r1 As Double
    Dim g1 As Double
    Dim b1 As Double
    Dim r2 As Double
    Dim g2 As Double
    Dim b2 As Double
    Dim rmean As Double
    Dim cd2 As Double
    
    r1 = Col1.r
    g1 = Col1.g
    b1 = Col1.b
    r2 = Col2.r
    g2 = Col2.g
    b2 = Col2.b
    
    rmean = (r1 + r2) / 2
    r = r1 - r2
    g = g1 - g2
    b = b1 - b2
    
    '(((512+rmean)*r*r)>>8) + 4*g*g + (((767-rmean)*b*b)>>8);
    cd2 = (((512 + rmean) * r * r) * 256) + 4 * g * g + (((767 - rmean) * b * b) * 256)
    ColorDistance2 = (cd2 / 16)
End Function

