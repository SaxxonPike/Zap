Attribute VB_Name = "modFont"
' This module uses GDI objects for faster graphics access.
' It does require a font first though -- use a picturebox
' on a form for a square one :)

Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type RGB
    r As Byte
    g As Byte
    b As Byte
End Type

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "USER32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Private FontDC As Long
Private FontBMP As Long
Private BackDC(0 To 1) As Long
Private BackBMP(0 To 1) As Long
Private TargetDC As Long
Private TargetHeight As Long
Private TargetWidth As Long
Private FontHeight As Long
Private fontwidth As Long
Private ScreenRect As RECT
Private PaletteBrushes(0 To 15) As Long
Private CharMap(0 To 79, 0 To 24) As Byte
Private ColMap(0 To 79, 0 To 24) As Byte
Private BlinkStatus As Byte
Private CharBufferDC As Long
Private CharBufferBMP As Long
Private xPaletteColor(0 To 15) As Long
Private xPaletteColorRGB(0 To 15) As RGB
Private xFontBackup(0 To 127, 0 To 255) As Byte

Public Function FontDCHeight()
    FontDCHeight = FontHeight
End Function

Public Function FontDCWidth()
    FontDCWidth = fontwidth
End Function

Public Function PaletteColor(colnum As Long) As Long
    PaletteColor = xPaletteColor(colnum)
End Function

Public Function PaletteColorRGB(colnum As Long) As RGB
    PaletteColorRGB = xPaletteColorRGB(colnum)
End Function

Public Sub LoadBinaryFont(filenumber As Long, FileOffset As Long, xFontHeight As Long, FontCharCount As Long)
    Dim xfont() As Byte
    Dim x As Long
    Dim y As Long
    Dim y2 As Long
    Dim c As Long
    Dim xoffs As Long
    Dim yoffs As Long
    Dim ah As Long 'adjusted height
    
    fontwidth = 8
    FontHeight = xFontHeight
    ReDim xfont(0 To (FontHeight * FontCharCount) - 1) As Byte
    
    ah = 8 - (FontHeight \ 2)
    Get #filenumber, FileOffset, xfont
    For c = 0 To 255
        xoffs = (c Mod 16) * fontwidth
        yoffs = (c \ 16) * FontHeight
        For y = 0 To 15
            y2 = y ' - ah
            For x = 0 To 7
                If y2 < 0 Or y2 >= FontHeight Then
                    SetPixel FontDC, xoffs + x, yoffs + y, vbBlack
                Else
                    If ((xfont((c * FontHeight) + y2) And (2 ^ (7 - x)))) <> 0 Then
                        SetPixel FontDC, xoffs + x, yoffs + y, vbWhite
                    Else
                        SetPixel FontDC, xoffs + x, yoffs + y, vbBlack
                    End If
                End If
            Next x
        Next y
    Next c
    
End Sub

Public Sub SetColor(colnum As Long, r As Byte, g As Byte, b As Byte)
    If PaletteBrushes(colnum) <> 0 Then
        DeleteObject PaletteBrushes(colnum)
    End If
    xPaletteColor(colnum) = RGB(r, g, b)
    xPaletteColorRGB(colnum).r = r
    xPaletteColorRGB(colnum).g = g
    xPaletteColorRGB(colnum).b = b
    PaletteBrushes(colnum) = CreateSolidBrush(xPaletteColor(colnum))
End Sub

Public Sub SetFontDCResource(newDC As Long, resourcenumber As Long, CharWidth As Long, CharHeight As Long)
    Dim x As Long
    Dim y As Long
    Dim r() As Byte
    ClearFontDC
    FontDC = CreateCompatibleDC(newDC)
    FontBMP = CreateCompatibleBitmap(newDC, CharWidth * 16, CharHeight * 16)
    SelectObject FontDC, FontBMP
    'BitBlt FontDC, 0, 0, CharWidth * 16, CharHeight * 16, newDC, 0, 0, vbSrcCopy
    fontwidth = CharWidth
    FontHeight = CharHeight
    CharBufferDC = CreateCompatibleDC(newDC)
    CharBufferBMP = CreateCompatibleBitmap(newDC, CharWidth * 2, CharHeight)
    SelectObject CharBufferDC, CharBufferBMP
    r() = LoadResData(resourcenumber, "CUSTOM")
    For x = 0 To UBound(r)
        SetPixel FontDC, x Mod 128, x \ 128, vbWhite * r(x)
    Next x
End Sub

Public Sub SetFontDC(newDC As Long, CharWidth As Long, CharHeight As Long)
    Dim x As Long
    Dim y As Long
    ClearFontDC
    FontDC = CreateCompatibleDC(newDC)
    FontBMP = CreateCompatibleBitmap(newDC, CharWidth * 16, CharHeight * 16)
    SelectObject FontDC, FontBMP
    BitBlt FontDC, 0, 0, CharWidth * 16, CharHeight * 16, newDC, 0, 0, vbSrcCopy
    fontwidth = CharWidth
    FontHeight = CharHeight
    CharBufferDC = CreateCompatibleDC(newDC)
    CharBufferBMP = CreateCompatibleBitmap(newDC, CharWidth * 2, CharHeight)
    SelectObject CharBufferDC, CharBufferBMP
    For x = 0 To 127
        For y = 0 To 255
            xFontBackup(x, y) = (GetPixel(FontDC, x, y) And 1)
        Next y
    Next x
End Sub

Public Sub RestoreFontDC()
    Dim x As Long
    Dim y As Long
    For x = 0 To 127
        For y = 0 To 255
            SetPixel FontDC, x, y, vbWhite * xFontBackup(x, y)
        Next y
    Next x
    fontwidth = 8
    FontHeight = 16
End Sub

Public Sub SetTargetDC(newDC As Long, Width As Long, Height As Long)
    If FontDC = 0 Then
        Exit Sub
    End If
    ClearTargetDC
    BackDC(0) = CreateCompatibleDC(newDC)
    BackBMP(0) = CreateCompatibleBitmap(newDC, Width * fontwidth, Height * FontHeight)
    BackDC(1) = CreateCompatibleDC(newDC)
    BackBMP(1) = CreateCompatibleBitmap(newDC, Width * fontwidth, Height * FontHeight)
    SelectObject BackDC(0), BackBMP(0)
    SelectObject BackDC(1), BackBMP(1)
    With ScreenRect
        .Left = 0
        .Right = Width * fontwidth
        .Top = 0
        .Bottom = Height * FontHeight
    End With
    TargetDC = newDC
End Sub

Public Sub SetBlinkStatus(bBlink As Boolean)
    If bBlink Then
        BlinkStatus = 0
    Else
        BlinkStatus = 1
    End If
End Sub

Private Sub ClearFontDC()
    If FontDC <> 0 Then
        DeleteDC FontDC
        DeleteObject FontBMP
        DeleteDC CharBufferDC
        DeleteObject CharBufferBMP
    End If
End Sub

Private Sub ClearTargetDC()
    Dim a As Long
    If BackDC(0) <> 0 Then
        For a = 0 To 1
            DeleteDC BackDC(a)
            DeleteObject BackBMP(a)
        Next a
    End If
    TargetDC = 0
    With ScreenRect
        .Left = 0
        .Right = 0
        .Bottom = 0
        .Top = 0
    End With
End Sub

Public Sub ClearDC()
    BitBlt BackDC(0), 0, 0, fontwidth * 60, FontHeight * 25, 0, 0, 0, vbBlackness
    BitBlt BackDC(1), 0, 0, fontwidth * 60, FontHeight * 25, 0, 0, 0, vbBlackness
End Sub

Public Sub UnloadFontSystem()
    Dim x As Long
    ClearTargetDC
    ClearFontDC
    For x = 0 To 15
        If PaletteBrushes(x) <> 0 Then
            DeleteObject PaletteBrushes(x)
        End If
    Next x
End Sub

Public Sub SetCharB(x As Long, y As Long, char As Byte, col As Byte)
    SetChar1 x, y, char, col, BackDC(0), , False
    If (col And 128) Then
        SetChar1 x, y, 0, col, BackDC(1), , False
    Else
        SetChar1 x, y, char, col, BackDC(1), , False
    End If
End Sub

Public Sub SetCharBScale(x As Long, y As Long, char As Byte, col As Byte, iScale As Single)
    SetChar1 x, y, char, col, BackDC(0), iScale, False
    If (col And 128) Then
        SetChar1 x, y, 0, col, BackDC(1), iScale, False
    Else
        SetChar1 x, y, char, col, BackDC(1), iScale, False
    End If
End Sub

Public Sub SetChar1(x As Long, y As Long, ByVal char As Byte, col As Byte, Optional hdc As Long = -1, Optional cScale As Single = 1, Optional CensorBlink As Boolean = True)
    If hdc = -1 Then
        hdc = TargetDC
    End If
    Dim fcol As Byte
    Dim bcol As Byte
    Dim destrect As RECT
    Dim forerect As RECT
    Dim a As Long
    With destrect
        .Left = 0
        .Top = 0
        .Right = fontwidth
        .Bottom = FontHeight
    End With
    With forerect
        .Left = fontwidth
        .Top = 0
        .Right = fontwidth + fontwidth
        .Bottom = FontHeight
    End With
    fcol = col And 15
    bcol = (col And 112) \ 16
    'background color
    FillRect CharBufferDC, destrect, PaletteBrushes(bcol)
    If (BlinkStatus = 1 And col >= 128) And CensorBlink = True Then
        char = 0
    End If
    If (bcol <> fcol) Then
        'invert for mask (R)
        BitBlt CharBufferDC, fontwidth, 0, fontwidth, FontHeight, FontDC, (char Mod 16) * fontwidth, (char \ 16) * FontHeight, vbNotSrcCopy
        'apply mask to background color (L)
        BitBlt CharBufferDC, 0, 0, fontwidth, FontHeight, CharBufferDC, fontwidth, 0, vbSrcAnd
        'foreground color (R)
        FillRect CharBufferDC, forerect, PaletteBrushes(fcol)
        BitBlt CharBufferDC, fontwidth, 0, fontwidth, FontHeight, FontDC, (char Mod 16) * fontwidth, (char \ 16) * FontHeight, vbSrcAnd
        'copy (L)
        BitBlt CharBufferDC, 0, 0, fontwidth, FontHeight, CharBufferDC, fontwidth, 0, vbSrcPaint
    End If
    'put it on our target DC as well
    If cScale = 1 Then
        BitBlt hdc, x * fontwidth, y * FontHeight, fontwidth, FontHeight, CharBufferDC, 0, 0, vbSrcCopy
    Else
        StretchBlt hdc, x * fontwidth, y * FontHeight, fontwidth * cScale, FontHeight * cScale, CharBufferDC, 0, 0, fontwidth, FontHeight, vbSrcCopy
    End If
    CharMap(x, y) = char
    ColMap(x, y) = col
End Sub

Public Sub RenderDC(Optional ByVal xTargetDC As Long = 0)
    If xTargetDC = 0 Then
        xTargetDC = TargetDC
    End If
    BitBlt xTargetDC, 0, 0, fontwidth * 60, FontHeight * 25, BackDC(BlinkStatus), 0, 0, vbSrcCopy
End Sub

Public Sub RenderDCScale(fScale As Single, Optional ByVal xTargetDC As Long = 0)
    If xTargetDC = 0 Then
        xTargetDC = TargetDC
    End If
    If fScale = 1 Then
        RenderDC xTargetDC
        Exit Sub
    End If
    StretchBlt xTargetDC, 0, 0, fontwidth * 60 * fScale, FontHeight * 25 * fScale, BackDC(BlinkStatus), 0, 0, fontwidth * 60, FontHeight * 25, vbSrcCopy
End Sub

Public Sub RenderCharDC(x As Long, y As Long, Optional iScale As Single = 1)
    If TargetDC = 0 Then
        Exit Sub
    End If
    If iScale = 1 Then
        BitBlt TargetDC, x * fontwidth, y * FontHeight, fontwidth, FontHeight, BackDC(BlinkStatus), x * fontwidth, y * FontHeight, vbSrcCopy
    ElseIf iScale > 0 Then
        StretchBlt TargetDC, x * fontwidth * iScale, y * FontHeight * iScale, fontwidth * iScale, FontHeight * iScale, BackDC(BlinkStatus), x * fontwidth, y * FontHeight, fontwidth, FontHeight, vbSrcCopy
    End If
End Sub

Public Sub RenderBlinksOnlyDC(Optional iScale As Single = 1)
    Dim x As Long
    Dim y As Long
    For y = 0 To 24
        For x = 0 To 59
            If (ColMap(x, y) And 128) <> 0 Then
                RenderCharDC x, y, iScale
            End If
        Next x
    Next y
End Sub

Public Function DisplayCharMap(x As Long, y As Long) As Byte
    DisplayCharMap = CharMap(x, y)
End Function

Public Function DisplayColorMap(x As Long, y As Long) As Byte
    DisplayColorMap = ColMap(x, y)
End Function

