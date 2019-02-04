VERSION 5.00
Begin VB.Form frmBoardInfo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Board Information"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3960
   ControlBox      =   0   'False
   Icon            =   "frmBoardInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   3960
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   3120
      Width           =   3015
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   2895
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   3735
      Begin VB.CheckBox chkDark 
         Caption         =   "Dark"
         Height          =   255
         Left            =   2400
         TabIndex        =   8
         ToolTipText     =   "Regular ZZT only: determines if the board is dark - the player will need torches to see in it."
         Top             =   2520
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   3
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "This determines where the player will go if they walk off the board on the East side."
         Top             =   2040
         Width           =   2775
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   2
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "This determines where the player will go if they walk off the board on the West side."
         Top             =   1680
         Width           =   2775
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "This determines where the player will go if they walk off the board on the South side."
         Top             =   1320
         Width           =   2775
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "This determines where the player will go if they walk off the board on the North side."
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox txtTime 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2520
         TabIndex        =   2
         Text            =   "0"
         ToolTipText     =   "Time limit of the board, in seconds."
         Top             =   480
         Width           =   495
      End
      Begin VB.CheckBox chkRestart 
         Caption         =   "Restart when zapped"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "This determines whether a player will have its position reset when it is damaged."
         Top             =   2520
         Width           =   1815
      End
      Begin VB.TextBox txtShots 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   1
         Text            =   "255"
         ToolTipText     =   "Maximum number of player shots that can simultaneously exist on the board."
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtTitle 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         MaxLength       =   40
         TabIndex        =   0
         Text            =   "Untitled"
         Top             =   0
         Width           =   3495
      End
      Begin VB.Label Label3 
         Caption         =   "Board E"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   16
         Top             =   2070
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Board W"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   1710
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Board S"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   1350
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Board N"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   990
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Time Limit:"
         Height          =   255
         Left            =   1560
         TabIndex        =   12
         Top             =   510
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Shots:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   510
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmBoardInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkDark_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub

Private Sub chkRestart_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub

Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim x As Long
    Dim y As Long
    Dim xshots As Byte
    Dim xtime As Integer
    Dim xBN As Byte
    Dim xBS As Byte
    Dim xBE As Byte
    Dim xBW As Byte
    Dim xzap As Byte
    Dim xdark As Byte
    
    'frmEdit.Enabled = False
    'frmPalette.Enabled = False
    DisableForms
    'populate board listing
    Combo1(0).AddItem "*None*"
    For x = 1 To World.BoardCount
        Combo1(0).AddItem CStr(x) + ": " + World.BoardName(x)
    Next x
    For x = 1 To 3
        For y = 0 To World.BoardCount
            Combo1(x).AddItem Combo1(0).List(y)
        Next y
    Next x
    txtTitle.Text = World.BoardName(CurrentBoard)
    World.GetBoardInfo CurrentBoard, xshots, xdark, xzap, xBN, xBS, xBW, xBE, xtime
    txtShots = xshots
    If xdark <> 0 Then chkDark.Value = 1
    If xzap <> 0 Then chkRestart.Value = 1
    If xBN < Combo1(0).ListCount Then Combo1(0).ListIndex = xBN Else Combo1(0).ListIndex = CurrentBoard
    If xBS < Combo1(1).ListCount Then Combo1(1).ListIndex = xBS Else Combo1(1).ListIndex = CurrentBoard
    If xBW < Combo1(2).ListCount Then Combo1(2).ListIndex = xBW Else Combo1(2).ListIndex = CurrentBoard
    If xBE < Combo1(3).ListCount Then Combo1(3).ListIndex = xBE Else Combo1(3).ListIndex = CurrentBoard
    txtTime = xtime
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'save changes and refresh edit form
    Dim c As Long
    Dim xshots As Byte
    Dim xtime As Integer
    Dim xBN As Byte
    Dim xBS As Byte
    Dim xBE As Byte
    Dim xBW As Byte
    Dim xzap As Byte
    Dim xdark As Byte
    
    World.SetBoardName CurrentBoard, txtTitle.Text
    
    c = CLng(txtShots)
    If c > 255 Then c = 255
    If c < 0 Then c = 0
    xshots = c
    
    c = CLng(txtTime)
    If c > 32767 Then c = 32767
    If c < 0 Then c = 0
    xtime = c
    
    If chkRestart.Value = 1 Then
        xzap = 255
    End If
    
    If chkDark.Value = 1 Then
        xdark = 255
    End If
    
    xBN = Combo1(0).ListIndex
    xBS = Combo1(1).ListIndex
    xBW = Combo1(2).ListIndex
    xBE = Combo1(3).ListIndex
    
    World.SetBoardInfo CurrentBoard, xshots, xdark, xzap, xBN, xBS, xBW, xBE, xtime
    
    'frmEdit.Enabled = True
    'frmPalette.Enabled = True
    EnableForms
    frmEdit.ConstructMenus
    frmEdit.SetMode
    
End Sub

Private Sub txtShots_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub

Private Sub txtTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub

Private Sub txtTitle_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub
