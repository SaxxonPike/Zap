VERSION 5.00
Begin VB.Form frmFillStyle 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Fill Pattern"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Index           =   1
      Left            =   2040
      Max             =   8
      Min             =   1
      TabIndex        =   4
      Top             =   4200
      Value           =   1
      Width           =   1695
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Index           =   0
      Left            =   2040
      Max             =   8
      Min             =   1
      TabIndex        =   3
      Top             =   3840
      Value           =   1
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   4560
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   120
      ScaleHeight     =   8
      ScaleMode       =   0  'User
      ScaleWidth      =   8
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "height"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "width"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   3840
      Width           =   1815
   End
End
Attribute VB_Name = "frmFillStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim x As Long
    Dim y As Long
    HScroll1(0).Value = 4
    HScroll1(1).Value = 4
    DisableForms
    For x = 0 To 7
        For y = 0 To 7
            FloodFillPattern(x, y) = False
        Next y
    Next x
    bFillPattern = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    EnableForms
End Sub

Private Sub HScroll1_Change(Index As Integer)
    Label1(Index).Caption = CStr(HScroll1(Index).Value)
    Select Case Index
        Case 0
            Label1(Index).Caption = "Width: " + Label1(Index).Caption
            lFillSizeX = HScroll1(0)
        Case 1
            Label1(Index).Caption = "Height: " + Label1(Index).Caption
            lFillSizeY = HScroll1(1)
    End Select
    RefreshGrid
End Sub

Private Sub HScroll1_Scroll(Index As Integer)
    HScroll1_Change Index
End Sub

Private Sub RefreshGrid()
    Dim x As Long
    Dim y As Long
    Picture1.Cls
    Picture1.ScaleWidth = HScroll1(0).Value
    Picture1.ScaleHeight = HScroll1(1).Value
    For x = 0 To Picture1.ScaleWidth - 1
        For y = 0 To Picture1.ScaleHeight - 1
            If FloodFillPattern(x, y) Then
                Picture1.Line (x, y)-(x + 1, y + 1), vbRed, BF
            End If
        Next y
    Next x
    For x = 1 To Picture1.ScaleWidth - 1
        Picture1.Line (x, 0)-(x, Picture1.ScaleHeight)
    Next x
    For y = 1 To Picture1.ScaleHeight - 1
        Picture1.Line (0, y)-(Picture1.ScaleWidth, y)
    Next y
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If x >= 0 And x <= Picture1.ScaleWidth And y >= 0 And y <= Picture1.ScaleHeight Then
        FloodFillPattern(Int(x), Int(y)) = Not FloodFillPattern(Int(x), Int(y))
        RefreshGrid
    End If
End Sub
