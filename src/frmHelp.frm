VERSION 5.00
Begin VB.Form frmHelp 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00800000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ZAP Help"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5295
   BeginProperty Font 
      Name            =   "Terminal"
      Size            =   9
      Charset         =   255
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0000FFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   527
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   353
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar VScroll1 
      Height          =   7905
      LargeChange     =   10
      Left            =   5040
      Max             =   100
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type LinkType
    lLinkLine As Long
    lLinkText As String
    lLinkDestination As String
End Type
Private xLinks() As LinkType
Private xHelp() As String

Public Sub ShowHelp(sHelpFile As String)
    Dim bHelpBuffer() As Byte
    Dim sCurrentString As String
    Dim lCurrentHelp As Long
    Dim lCurrentLink As Long
    Dim lOffset As Long
    ReDim xHelp(0 To 1) As String
    ReDim xLinks(0) As LinkType
    lCurrentHelp = 1
    lCurrentLink = -1
    lOffset = 0
    If Help.LoadHelp(sHelpFile, bHelpBuffer()) = False Then
        Exit Sub
    End If
    Do While lOffset <= UBound(bHelpBuffer)
        If bHelpBuffer(lOffset) <> 13 And bHelpBuffer(lOffset) <> 10 Then
            xHelp(lCurrentHelp) = xHelp(lCurrentHelp) + Chr$(bHelpBuffer(lOffset))
        ElseIf bHelpBuffer(lOffset) = 13 Then
            lCurrentHelp = lCurrentHelp + 1
            ReDim Preserve xHelp(0 To lCurrentHelp) As String
        End If
        lOffset = lOffset + 1
    Loop
    For lCurrentHelp = 0 To UBound(xHelp)
        Select Case Left$(xHelp(lCurrentHelp), 1)
            Case "!" 'help link
                lCurrentLink = lCurrentLink + 1
                ReDim Preserve xLinks(0 To lCurrentLink) As LinkType
                If InStr(xHelp(lCurrentHelp), ";") > 0 Then
                    xLinks(lCurrentLink).lLinkDestination = Mid$(Left$(xHelp(lCurrentHelp), InStr(xHelp(lCurrentHelp), ";") - 1), 2)
                    xLinks(lCurrentLink).lLinkText = Mid$(xHelp(lCurrentHelp), InStr(xHelp(lCurrentHelp), ";") + 1)
                Else
                    xLinks(lCurrentLink).lLinkText = xHelp(lCurrentHelp)
                End If
                xLinks(lCurrentLink).lLinkLine = lCurrentHelp
                xHelp(lCurrentHelp) = "!" + xLinks(lCurrentLink).lLinkText
        End Select
    Next lCurrentHelp
    VScroll1.Value = 0
    DrawHelp
    Me.Visible = True
End Sub

Sub DrawHelp()
    Dim x As Long
    Dim s As String
    If UBound(xHelp) > (Me.ScaleHeight \ 16) Then
        VScroll1.Enabled = True
        VScroll1.Max = UBound(xHelp) - (Me.ScaleHeight \ 16)
    Else
        VScroll1.Value = 0
        VScroll1.Enabled = False
    End If
    Me.Cls
    For x = 0 To UBound(xHelp)
        Me.CurrentX = 8
        Me.CurrentY = ((-16 * (VScroll1.Value)) + (x * 16))
        If Me.CurrentY < Me.ScaleHeight And Me.CurrentY > -16 Then
            Select Case Left$(xHelp(x), 1)
                Case "!"
                    Me.ForeColor = vbMagenta
                    Me.Print " > ";
                    Me.ForeColor = vbWhite
                    Me.Font.Underline = True
                    Me.Print Mid$(xHelp(x), 2)
                    Me.Font.Underline = False
                Case "$", ":"
                    s = xHelp(x)
                    If Left$(s, 1) = "$" Then
                        s = Mid$(s, 2)
                    ElseIf InStr(s, ";") > 0 Then
                        s = Mid$(s, InStr(s, ";") + 1)
                    Else
                        s = ""
                    End If
                    Me.ForeColor = vbWhite
                    Me.CurrentX = ((Me.ScaleWidth - VScroll1.Width) \ 2) - ((Len(s) * 8) \ 2)
                    Me.Print s
                Case Else
                    Me.ForeColor = vbYellow
                    Print xHelp(x)
            End Select
        End If
    Next x
End Sub

Private Sub Form_Load()
    Me.Icon = frmEdit.Icon
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim x1 As Long
    Dim Dest As String
    For x1 = 0 To UBound(xLinks)
        If xLinks(x1).lLinkLine = (Int(Y) \ 16) + VScroll1.Value Then
            Dest = xLinks(x1).lLinkDestination
        End If
    Next x1
    If Dest <> "" Then
        Debug.Print "HELP navigation:", Dest
        If Left$(Dest, 1) = "-" Then
            ShowHelp Mid$(Dest, 2)
        Else
            Dest = ":" + UCase$(Dest) + ";"
            For x1 = 0 To UBound(xHelp)
                If xHelp(x1) <> "" Then
                    If InStr(UCase$(xHelp(x1)), Dest) = 1 Then
                        VScroll1.Value = x1
                        Exit For
                    End If
                End If
            Next x1
        End If
    End If
End Sub

Private Sub VScroll1_Change()
    DrawHelp
End Sub

Private Sub VScroll1_Scroll()
    VScroll1_Change
End Sub
