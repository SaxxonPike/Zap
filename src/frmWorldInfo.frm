VERSION 5.00
Begin VB.Form frmWorldInfo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "World Information"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5175
   ControlBox      =   0   'False
   Icon            =   "frmWorldInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   5175
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Player"
      Height          =   1815
      Left            =   120
      TabIndex        =   36
      Top             =   3720
      Width           =   4935
      Begin VB.CheckBox chkKeys 
         Caption         =   "Check2"
         Height          =   255
         Index           =   6
         Left            =   2910
         TabIndex        =   33
         ToolTipText     =   "Player's keys. These allow the player to unlock doors."
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox chkKeys 
         Caption         =   "Check2"
         Height          =   255
         Index           =   5
         Left            =   2550
         TabIndex        =   32
         ToolTipText     =   "Player's keys. These allow the player to unlock doors."
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox chkKeys 
         Caption         =   "Check2"
         Height          =   255
         Index           =   4
         Left            =   2190
         TabIndex        =   31
         ToolTipText     =   "Player's keys. These allow the player to unlock doors."
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox chkKeys 
         Caption         =   "Check2"
         Height          =   255
         Index           =   3
         Left            =   1830
         TabIndex        =   30
         ToolTipText     =   "Player's keys. These allow the player to unlock doors."
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox chkKeys 
         Caption         =   "Check2"
         Height          =   255
         Index           =   2
         Left            =   1470
         TabIndex        =   29
         ToolTipText     =   "Player's keys. These allow the player to unlock doors."
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox chkKeys 
         Caption         =   "Check2"
         Height          =   255
         Index           =   1
         Left            =   1110
         TabIndex        =   28
         ToolTipText     =   "Player's keys. These allow the player to unlock doors."
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox chkKeys 
         Caption         =   "Check2"
         Height          =   255
         Index           =   0
         Left            =   750
         TabIndex        =   27
         ToolTipText     =   "Player's keys. These allow the player to unlock doors."
         Top             =   1320
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   135
         Index           =   6
         Left            =   2880
         ScaleHeight     =   105
         ScaleWidth      =   225
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   1590
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H80000008&
         Height          =   135
         Index           =   5
         Left            =   2520
         ScaleHeight     =   105
         ScaleWidth      =   225
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   1590
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         ForeColor       =   &H80000008&
         Height          =   135
         Index           =   4
         Left            =   2160
         ScaleHeight     =   105
         ScaleWidth      =   225
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   1590
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         ForeColor       =   &H80000008&
         Height          =   135
         Index           =   3
         Left            =   1800
         ScaleHeight     =   105
         ScaleWidth      =   225
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   1590
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         ForeColor       =   &H80000008&
         Height          =   135
         Index           =   2
         Left            =   1440
         ScaleHeight     =   105
         ScaleWidth      =   225
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   1590
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         ForeColor       =   &H80000008&
         Height          =   135
         Index           =   1
         Left            =   1080
         ScaleHeight     =   105
         ScaleWidth      =   225
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   1590
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H80000008&
         Height          =   135
         Index           =   0
         Left            =   720
         ScaleHeight     =   105
         ScaleWidth      =   225
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   1590
         Width           =   255
      End
      Begin VB.TextBox txtPlayer 
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
         Index           =   7
         Left            =   4080
         TabIndex        =   26
         Text            =   "0"
         ToolTipText     =   "This determines how many game cycles the player has left to be energized."
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtPlayer 
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
         Index           =   6
         Left            =   4080
         TabIndex        =   25
         Text            =   "0"
         ToolTipText     =   "This determines how many game cycles the player has left on his torch."
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtPlayer 
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
         Index           =   5
         Left            =   2160
         TabIndex        =   24
         Text            =   "0"
         ToolTipText     =   "Player's time. This is best kept zero even if some boards have time limits."
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtPlayer 
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
         Index           =   4
         Left            =   2160
         TabIndex        =   23
         Text            =   "0"
         ToolTipText     =   "Player's score."
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtPlayer 
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
         Index           =   3
         Left            =   2160
         TabIndex        =   22
         Text            =   "0"
         ToolTipText     =   "Player's torches. Torches allow the player to see on dark boards."
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtPlayer 
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
         Index           =   2
         Left            =   720
         TabIndex        =   21
         Text            =   "0"
         ToolTipText     =   "Player's gems. Generally these are used for currency in games."
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtPlayer 
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
         Index           =   1
         Left            =   720
         TabIndex        =   20
         Text            =   "0"
         ToolTipText     =   "Player's ammunition. This allows the player to shoot bullets."
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtPlayer 
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
         Index           =   0
         Left            =   720
         TabIndex        =   19
         Text            =   "100"
         ToolTipText     =   "The player's health. When this reaches 0, the game ends."
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Keys"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   45
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Energy Cycles"
         Height          =   255
         Index           =   7
         Left            =   2880
         TabIndex        =   44
         Top             =   630
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Torch Cycles"
         Height          =   255
         Index           =   6
         Left            =   3000
         TabIndex        =   43
         Top             =   270
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Time"
         Height          =   255
         Index           =   5
         Left            =   1560
         TabIndex        =   42
         Top             =   990
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Score"
         Height          =   255
         Index           =   4
         Left            =   1560
         TabIndex        =   41
         Top             =   630
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Torch"
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   40
         Top             =   270
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Gems"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   39
         Top             =   990
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Ammo"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   38
         Top             =   630
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Health"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   37
         Top             =   270
         Width           =   615
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   1080
      TabIndex        =   34
      Top             =   5640
      Width           =   3015
   End
   Begin VB.CheckBox chkSave 
      Caption         =   "Saved Game"
      Height          =   255
      Left            =   3720
      TabIndex        =   1
      ToolTipText     =   $"frmWorldInfo.frx":038A
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtGameName 
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
      Left            =   1200
      MaxLength       =   20
      TabIndex        =   0
      ToolTipText     =   "Game Title. This is what's shown on the title screen in place of the game's file name."
      Top             =   120
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "Flags"
      Height          =   3135
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   4935
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFFFFF&
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
         Index           =   15
         Left            =   2520
         MaxLength       =   20
         TabIndex        =   18
         ToolTipText     =   "Flags keep track of things not visible to the player. In SuperZZT, preceding a flag with ""Z"" changes the Z-element text."
         Top             =   2760
         Width           =   2295
      End
      Begin VB.TextBox Text2 
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
         Index           =   14
         Left            =   2520
         MaxLength       =   20
         TabIndex        =   17
         ToolTipText     =   "Flags keep track of things not visible to the player. In SuperZZT, preceding a flag with ""Z"" changes the Z-element text."
         Top             =   2400
         Width           =   2295
      End
      Begin VB.TextBox Text2 
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
         Index           =   13
         Left            =   2520
         MaxLength       =   20
         TabIndex        =   16
         ToolTipText     =   "Flags keep track of things not visible to the player. In SuperZZT, preceding a flag with ""Z"" changes the Z-element text."
         Top             =   2040
         Width           =   2295
      End
      Begin VB.TextBox Text2 
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
         Index           =   12
         Left            =   2520
         MaxLength       =   20
         TabIndex        =   15
         ToolTipText     =   "Flags keep track of things not visible to the player. In SuperZZT, preceding a flag with ""Z"" changes the Z-element text."
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox Text2 
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
         Index           =   11
         Left            =   2520
         MaxLength       =   20
         TabIndex        =   14
         ToolTipText     =   "Flags keep track of things not visible to the player. In SuperZZT, preceding a flag with ""Z"" changes the Z-element text."
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox Text2 
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
         Index           =   10
         Left            =   2520
         MaxLength       =   20
         TabIndex        =   13
         ToolTipText     =   "Flags keep track of things not visible to the player. In SuperZZT, preceding a flag with ""Z"" changes the Z-element text."
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox Text2 
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
         Index           =   9
         Left            =   2520
         MaxLength       =   20
         TabIndex        =   12
         ToolTipText     =   "Flags keep track of things not visible to the player. In SuperZZT, preceding a flag with ""Z"" changes the Z-element text."
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox Text2 
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
         Index           =   8
         Left            =   2520
         MaxLength       =   20
         TabIndex        =   11
         ToolTipText     =   "Flags keep track of things not visible to the player. In SuperZZT, preceding a flag with ""Z"" changes the Z-element text."
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox Text2 
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
         Index           =   7
         Left            =   120
         MaxLength       =   20
         TabIndex        =   10
         ToolTipText     =   "Flags keep track of things not visible to the player. In SuperZZT, preceding a flag with ""Z"" changes the Z-element text."
         Top             =   2760
         Width           =   2295
      End
      Begin VB.TextBox Text2 
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
         Index           =   6
         Left            =   120
         MaxLength       =   20
         TabIndex        =   9
         ToolTipText     =   "Flags keep track of things not visible to the player. In SuperZZT, preceding a flag with ""Z"" changes the Z-element text."
         Top             =   2400
         Width           =   2295
      End
      Begin VB.TextBox Text2 
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
         Index           =   5
         Left            =   120
         MaxLength       =   20
         TabIndex        =   8
         ToolTipText     =   "Flags keep track of things not visible to the player. In SuperZZT, preceding a flag with ""Z"" changes the Z-element text."
         Top             =   2040
         Width           =   2295
      End
      Begin VB.TextBox Text2 
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
         Index           =   4
         Left            =   120
         MaxLength       =   20
         TabIndex        =   7
         ToolTipText     =   "Flags keep track of things not visible to the player. In SuperZZT, preceding a flag with ""Z"" changes the Z-element text."
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox Text2 
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
         Index           =   3
         Left            =   120
         MaxLength       =   20
         TabIndex        =   6
         ToolTipText     =   "Flags keep track of things not visible to the player. In SuperZZT, preceding a flag with ""Z"" changes the Z-element text."
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox Text2 
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
         Index           =   2
         Left            =   120
         MaxLength       =   20
         TabIndex        =   5
         ToolTipText     =   "Flags keep track of things not visible to the player. In SuperZZT, preceding a flag with ""Z"" changes the Z-element text."
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox Text2 
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
         Index           =   1
         Left            =   120
         MaxLength       =   20
         TabIndex        =   4
         ToolTipText     =   "Flags keep track of things not visible to the player. In SuperZZT, preceding a flag with ""Z"" changes the Z-element text."
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox Text2 
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
         Index           =   0
         Left            =   120
         MaxLength       =   20
         TabIndex        =   2
         ToolTipText     =   "Flags keep track of things not visible to the player. In SuperZZT, preceding a flag with ""Z"" changes the Z-element text."
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Game Name"
      Height          =   255
      Left            =   120
      TabIndex        =   35
      Top             =   150
      Width           =   975
   End
End
Attribute VB_Name = "frmWorldInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type xtiWorldBuffer
    xAmmo As Integer
    xGems As Integer
    xKeys(0 To 6) As Byte
    xHealth As Integer
    xTorches As Integer
    xTorchCycles As Integer
    xEnergyCycles As Integer
    xScore As Integer
    xGameName As String * 20
    xFlags(0 To 15) As String
    xTimePassed As Integer
    xLocked As Byte
End Type
Private WorldBuffer As xtiWorldBuffer

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim x As Long
    'frmEdit.Enabled = False
    'frmPalette.Enabled = False
    DisableForms
    For x = 10 To 15
        Text2(x).Visible = World.IsSuperZZT
    Next x
    With WorldBuffer
        World.GetWorldInfo1 .xAmmo, .xGems, .xKeys(), .xHealth, .xTorches, .xTorchCycles, .xEnergyCycles
        World.GetWorldInfo2 .xScore, .xFlags(), .xTimePassed, .xLocked
        txtPlayer(0) = .xHealth
        txtPlayer(1) = .xAmmo
        txtPlayer(2) = .xGems
        txtPlayer(3) = .xTorches
        txtPlayer(4) = .xScore
        txtPlayer(5) = .xTimePassed
        txtPlayer(6) = .xTorchCycles
        txtPlayer(7) = .xEnergyCycles
        For x = 0 To 6
            If .xKeys(x) <> 0 Then
                chkKeys(x).Value = 1
            End If
        Next x
        For x = 0 To 15
            Text2(x) = .xFlags(x)
        Next x
        txtGameName = World.WorldName
        If .xLocked <> 0 Then
            chkSave.Value = 1
        End If
    End With
    If World.IsSuperZZT Then
        Label2(6) = "Z elements"
        Label2(3) = "Unused"
        txtPlayer(3).ToolTipText = "This value is said to be unused by SuperZZT. It is still offered here for debug purposes."
        txtPlayer(6).ToolTipText = "'Z' elements. These are shown on the left side of the screen as a number. If set to -1, it is not visible at all."
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim x As Long
    With WorldBuffer
        .xHealth = Val(txtPlayer(0))
        .xAmmo = Val(txtPlayer(1))
        .xGems = Val(txtPlayer(2))
        .xTorches = Val(txtPlayer(3))
        .xScore = Val(txtPlayer(4))
        .xTimePassed = Val(txtPlayer(5))
        .xTorchCycles = Val(txtPlayer(6))
        .xEnergyCycles = Val(txtPlayer(7))
        For x = 0 To 6
            If chkKeys(x).Value <> 0 Then
                .xKeys(x) = 255
            End If
        Next x
        For x = 0 To 15
            .xFlags(x) = Text2(x)
        Next x
        World.SetWorldName txtGameName
        If chkSave.Value <> 0 Then
            .xLocked = 1
        End If
        World.SetWorldInfo1 .xAmmo, .xGems, .xKeys(), .xHealth, .xTorches, .xTorchCycles, .xEnergyCycles
        World.SetWorldInfo2 .xScore, .xFlags(), .xTimePassed, .xLocked
    End With
    
    'frmEdit.Enabled = True
    'frmPalette.Enabled = True
    EnableForms
    frmEdit.SetMode
End Sub
