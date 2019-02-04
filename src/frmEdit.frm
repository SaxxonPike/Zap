VERSION 5.00
Begin VB.Form frmEdit 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7575
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   7560
   DrawStyle       =   2  'Dot
   Icon            =   "frmEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   7575
   ScaleWidth      =   7560
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar VScroll 
      Enabled         =   0   'False
      Height          =   6255
      LargeChange     =   12
      Left            =   7200
      Max             =   55
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar HScroll 
      Enabled         =   0   'False
      Height          =   255
      LargeChange     =   30
      Left            =   0
      Max             =   36
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   6000
      Width           =   7215
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6240
      ScaleHeight     =   375
      ScaleWidth      =   900
      TabIndex        =   26
      Top             =   4800
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tools"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      Top             =   6240
      Width           =   7455
      Begin VB.VScrollBar vsRBU 
         Height          =   255
         Left            =   3240
         Max             =   0
         Min             =   10
         TabIndex        =   34
         Top             =   1005
         Width           =   255
      End
      Begin VB.CheckBox lblObjCount 
         Caption         =   "objects"
         Height          =   300
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Free objects remaining - when toggled, will show objects on-screen"
         Top             =   630
         Width           =   615
      End
      Begin VB.CommandButton cmdWorldInfo 
         Caption         =   "W"
         Height          =   375
         Left            =   3960
         TabIndex        =   29
         TabStop         =   0   'False
         ToolTipText     =   "World Information"
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton cmdBoardInfo 
         Caption         =   "B"
         Height          =   375
         Left            =   3960
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "Board Information"
         Top             =   240
         Width           =   375
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IntegralHeight  =   0   'False
         ItemData        =   "frmEdit.frx":038A
         Left            =   120
         List            =   "frmEdit.frx":046C
         Style           =   2  'Dropdown List
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   240
         Width           =   3735
      End
      Begin VB.CheckBox chkStats 
         Caption         =   "S"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Has stats."
         Top             =   720
         Width           =   375
      End
      Begin VB.CheckBox chkDefaultColor 
         Caption         =   "D"
         Height          =   255
         Left            =   600
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Default color."
         Top             =   720
         Width           =   375
      End
      Begin VB.CheckBox chkBlinks 
         Caption         =   "B"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Allow blinks. (set background to bottom color row)"
         Top             =   960
         Value           =   1  'Checked
         Width           =   375
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00000000&
         Height          =   300
         Left            =   1860
         ScaleHeight     =   1
         ScaleMode       =   0  'User
         ScaleWidth      =   10
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Tile Buffer"
         Top             =   960
         Width           =   1255
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00000000&
         Height          =   300
         Left            =   1860
         ScaleHeight     =   1
         ScaleMode       =   0  'User
         ScaleWidth      =   10
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Shortcuts"
         Top             =   630
         Width           =   1255
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00000000&
         Height          =   540
         Left            =   1080
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   41
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Character display - this shows you in both enlarged and tile view what the selected tile looks like."
         Top             =   690
         Width           =   675
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         DrawMode        =   6  'Mask Pen Not
         Height          =   375
         Index           =   15
         Left            =   6960
         ScaleHeight     =   1
         ScaleMode       =   0  'User
         ScaleWidth      =   1
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "15: White"
         Top             =   600
         Width           =   375
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         DrawMode        =   6  'Mask Pen Not
         Height          =   375
         Index           =   14
         Left            =   6600
         ScaleHeight     =   1
         ScaleMode       =   0  'User
         ScaleWidth      =   1
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "14: Yellow"
         Top             =   600
         Width           =   375
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         DrawMode        =   6  'Mask Pen Not
         Height          =   375
         Index           =   13
         Left            =   6240
         ScaleHeight     =   1
         ScaleMode       =   0  'User
         ScaleWidth      =   1
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "13: Light Purple"
         Top             =   600
         Width           =   375
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         DrawMode        =   6  'Mask Pen Not
         Height          =   375
         Index           =   12
         Left            =   5880
         ScaleHeight     =   1
         ScaleMode       =   0  'User
         ScaleWidth      =   1
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "12: Light Red"
         Top             =   600
         Width           =   375
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         DrawMode        =   6  'Mask Pen Not
         Height          =   375
         Index           =   11
         Left            =   5520
         ScaleHeight     =   1
         ScaleMode       =   0  'User
         ScaleWidth      =   1
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "11: Light Cyan"
         Top             =   600
         Width           =   375
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         DrawMode        =   6  'Mask Pen Not
         Height          =   375
         Index           =   10
         Left            =   5160
         ScaleHeight     =   1
         ScaleMode       =   0  'User
         ScaleWidth      =   1
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "10: Light Green"
         Top             =   600
         Width           =   375
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         DrawMode        =   6  'Mask Pen Not
         Height          =   375
         Index           =   9
         Left            =   4800
         ScaleHeight     =   1
         ScaleMode       =   0  'User
         ScaleWidth      =   1
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "09: Light Blue"
         Top             =   600
         Width           =   375
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         DrawMode        =   6  'Mask Pen Not
         Height          =   375
         Index           =   8
         Left            =   4440
         ScaleHeight     =   1
         ScaleMode       =   0  'User
         ScaleWidth      =   1
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "08: Dark Gray"
         Top             =   600
         Width           =   375
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         DrawMode        =   6  'Mask Pen Not
         Height          =   375
         Index           =   7
         Left            =   6960
         ScaleHeight     =   1
         ScaleMode       =   0  'User
         ScaleWidth      =   1
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "07: Light Gray"
         Top             =   240
         Width           =   375
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         DrawMode        =   6  'Mask Pen Not
         Height          =   375
         Index           =   6
         Left            =   6600
         ScaleHeight     =   1
         ScaleMode       =   0  'User
         ScaleWidth      =   1
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "06: Dark Brown"
         Top             =   240
         Width           =   375
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         DrawMode        =   6  'Mask Pen Not
         Height          =   375
         Index           =   5
         Left            =   6240
         ScaleHeight     =   1
         ScaleMode       =   0  'User
         ScaleWidth      =   1
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "05: Dark Purple"
         Top             =   240
         Width           =   375
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         DrawMode        =   6  'Mask Pen Not
         Height          =   375
         Index           =   4
         Left            =   5880
         ScaleHeight     =   1
         ScaleMode       =   0  'User
         ScaleWidth      =   1
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "04: Dark Red"
         Top             =   240
         Width           =   375
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         DrawMode        =   6  'Mask Pen Not
         Height          =   375
         Index           =   3
         Left            =   5520
         ScaleHeight     =   1
         ScaleMode       =   0  'User
         ScaleWidth      =   1
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "03: Dark Cyan"
         Top             =   240
         Width           =   375
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         DrawMode        =   6  'Mask Pen Not
         Height          =   375
         Index           =   2
         Left            =   5160
         ScaleHeight     =   1
         ScaleMode       =   0  'User
         ScaleWidth      =   1
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "02: Dark Green"
         Top             =   240
         Width           =   375
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         DrawMode        =   6  'Mask Pen Not
         Height          =   375
         Index           =   1
         Left            =   4800
         ScaleHeight     =   1
         ScaleMode       =   0  'User
         ScaleWidth      =   1
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "01: Dark Blue"
         Top             =   240
         Width           =   375
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         DrawMode        =   6  'Mask Pen Not
         Height          =   375
         Index           =   0
         Left            =   4440
         ScaleHeight     =   1
         ScaleMode       =   0  'User
         ScaleWidth      =   1
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "00: Black"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblRBU 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   3510
         TabIndex        =   35
         ToolTipText     =   "Random Buffer Use - determines how many tiles in the backbuffer will be used (randomly) while painting."
         Top             =   1005
         Width           =   345
      End
      Begin VB.Label lblCoord 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "coord"
         Height          =   255
         Left            =   5520
         TabIndex        =   32
         ToolTipText     =   "Coordinates and, if applicable, [object number]"
         Top             =   1005
         Width           =   1815
      End
      Begin VB.Label lblColor 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "color"
         Height          =   255
         Left            =   3960
         TabIndex        =   25
         ToolTipText     =   "Currently selected color"
         Top             =   1005
         Width           =   1455
      End
   End
   Begin VB.Timer tmrMain 
      Interval        =   80
      Left            =   1680
      Top             =   120
   End
   Begin VB.PictureBox picMain 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DrawStyle       =   2  'Dot
      Height          =   6000
      Left            =   0
      MousePointer    =   2  'Cross
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   400
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   0
      Width           =   7200
      Begin VB.Timer tmrCursor 
         Interval        =   5
         Left            =   2280
         Top             =   120
      End
   End
   Begin VB.PictureBox picExport 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DrawStyle       =   2  'Dot
      Height          =   6000
      Left            =   0
      MousePointer    =   2  'Cross
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   400
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   480
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   7200
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "[&N] New ZZT World"
      End
      Begin VB.Menu mnuFileNewSZT 
         Caption         =   "[&U] New Super ZZT World"
      End
      Begin VB.Menu mnuFileLoad 
         Caption         =   "[&L] Load World"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "[&S] Save World"
      End
      Begin VB.Menu mnuFileRecent2 
         Caption         =   "[&C] Recent Files"
         Begin VB.Menu mnuFileRecent 
            Caption         =   "-"
            Index           =   0
         End
      End
      Begin VB.Menu mnuFile5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileTest 
         Caption         =   "[&E] Test World"
      End
      Begin VB.Menu mnuFileTest2 
         Caption         =   "[&R] Test World (with font+palette)"
      End
      Begin VB.Menu mnuFile3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileFont 
         Caption         =   "[&F] Fonts..."
         Begin VB.Menu mnuFileLoadFont 
            Caption         =   "[&L] Load Font"
         End
         Begin VB.Menu mnuFileUnloadFont 
            Caption         =   "[&U] Unload Font"
         End
      End
      Begin VB.Menu mnuFilePalette 
         Caption         =   "[&P] Palettes..."
         Begin VB.Menu mnuFileLoadPalette 
            Caption         =   "[&L] Load Palette"
         End
         Begin VB.Menu mnuFileUnloadPalette 
            Caption         =   "[&U] Unload Palette"
         End
      End
      Begin VB.Menu mnuFile4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileTransfer 
         Caption         =   "[&T] Transfer Board"
         Begin VB.Menu mnuFileTransfer1 
            Caption         =   "[&I] Import"
            Begin VB.Menu mnuFileTransferIB 
               Caption         =   "[&B] from Board File"
            End
            Begin VB.Menu mnuFileTransferIW 
               Caption         =   "[&W] from Another World"
            End
            Begin VB.Menu mnuFileTransferII 
               Caption         =   "[&I] from 60x25 Image"
            End
         End
         Begin VB.Menu mnuFileTransfer2 
            Caption         =   "[&E] Export"
            Begin VB.Menu mnuFileTransferEB 
               Caption         =   "[&B] to Board File"
            End
            Begin VB.Menu mnuFileTransferES 
               Caption         =   "[&S] to Screenshot"
            End
         End
      End
      Begin VB.Menu mnuFile2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "[&X] Exit"
      End
   End
   Begin VB.Menu mnuPaste 
      Caption         =   "Paste"
      Visible         =   0   'False
      Begin VB.Menu mnuPasteFlipH 
         Caption         =   "[&H] Flip Horizontally"
      End
      Begin VB.Menu mnuPasteFlipV 
         Caption         =   "[&V] Flip Vertically"
      End
      Begin VB.Menu mnuPaste1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPasteExecute 
         Caption         =   "[&P] Paste Here"
      End
   End
   Begin VB.Menu mnuFunctions 
      Caption         =   "Functions"
      Visible         =   0   'False
      Begin VB.Menu mnuFunctionsCopy 
         Caption         =   "[&C] Copy this tile to Buffer"
      End
      Begin VB.Menu mnuFunctionsEditStats 
         Caption         =   "[&E] Edit Tile..."
      End
      Begin VB.Menu mnuFunctionsFlood 
         Caption         =   "[&X] Flood fill"
      End
      Begin VB.Menu mnuFunctionsFloodPattern 
         Caption         =   "[&L] Flood fill pattern"
      End
      Begin VB.Menu mnuFunctions1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFunctionsArea 
         Caption         =   "[&A] Area..."
         Begin VB.Menu mnuFunctionsCutRect 
            Caption         =   "[&Z] Cut Area Rectangle"
         End
         Begin VB.Menu mnuFunctionsCopyRect 
            Caption         =   "[&C] Copy Area Rectangle"
         End
         Begin VB.Menu mnuFunctionsPasteRect 
            Caption         =   "[&V] Paste Area Rectangle"
         End
         Begin VB.Menu mnuFunctionsMoveRect 
            Caption         =   "[&M] Move Area Rectangle"
         End
         Begin VB.Menu mnuFunctionsFill 
            Caption         =   "[&X] Fill Area Rectangle"
         End
         Begin VB.Menu mnuFunctionsFillPattern 
            Caption         =   "[&L] Fill Pattern Rectangle"
         End
         Begin VB.Menu mnuFunctionsDeselect 
            Caption         =   "[&D] Deselect Area Rectangle"
         End
         Begin VB.Menu mnuFunctionsOutline 
            Caption         =   "[&S] Outline Area Rectangle"
         End
         Begin VB.Menu mnuFunctionDeleteObjectsRect 
            Caption         =   "[&R] Delete All Objects in Rectangle"
         End
      End
      Begin VB.Menu mnuFunctions4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFunctionsEditPattern 
         Caption         =   "[&P] Edit pattern..."
      End
      Begin VB.Menu mnuFunctions2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuItems 
         Caption         =   "[&1] Elements: Items"
         Begin VB.Menu mnuItems1 
            Caption         =   "[&Z] Player"
            Index           =   0
         End
         Begin VB.Menu mnuItems1 
            Caption         =   "[&X] Player CLONE"
            Index           =   1
         End
         Begin VB.Menu mnuItems1 
            Caption         =   "[&A] Ammo"
            Index           =   2
         End
         Begin VB.Menu mnuItems1 
            Caption         =   "[&T] Torch"
            Index           =   3
         End
         Begin VB.Menu mnuItems1 
            Caption         =   "[&G] Gem"
            Index           =   4
         End
         Begin VB.Menu mnuItems1 
            Caption         =   "[&K] Key"
            Index           =   5
         End
         Begin VB.Menu mnuItems1 
            Caption         =   "[&D] Door"
            Index           =   6
         End
         Begin VB.Menu mnuItems1 
            Caption         =   "[&S] Scroll"
            Index           =   7
         End
         Begin VB.Menu mnuItems1 
            Caption         =   "[&P] Passage"
            Index           =   8
         End
         Begin VB.Menu mnuItems1 
            Caption         =   "[&U] Duplicator"
            Index           =   9
         End
         Begin VB.Menu mnuItems1 
            Caption         =   "[&B] Bomb"
            Index           =   10
         End
         Begin VB.Menu mnuItems1 
            Caption         =   "[&E] Energizer"
            Index           =   11
         End
         Begin VB.Menu mnuItems1 
            Caption         =   "[&1] Clockwise"
            Index           =   12
         End
         Begin VB.Menu mnuItems1 
            Caption         =   "[&2] Counter"
            Index           =   13
         End
         Begin VB.Menu mnuItems1 
            Caption         =   "[&O] Stone of Power"
            Index           =   14
         End
      End
      Begin VB.Menu mnuCreatures 
         Caption         =   "[&2] Elements: Creatures"
         Begin VB.Menu mnuCreatures1 
            Caption         =   "[&B] Bear"
            Index           =   0
         End
         Begin VB.Menu mnuCreatures1 
            Caption         =   "[&R] Ruffian"
            Index           =   1
         End
         Begin VB.Menu mnuCreatures1 
            Caption         =   "[&O] Object"
            Index           =   2
         End
         Begin VB.Menu mnuCreatures1 
            Caption         =   "[&V] Slime"
            Index           =   3
         End
         Begin VB.Menu mnuCreatures1 
            Caption         =   "[&Y] Shark"
            Index           =   4
         End
         Begin VB.Menu mnuCreatures1 
            Caption         =   "[&G] Spinning Gun"
            Index           =   5
         End
         Begin VB.Menu mnuCreatures1 
            Caption         =   "[&P] Pusher"
            Index           =   6
         End
         Begin VB.Menu mnuCreatures1 
            Caption         =   "[&L] Lion"
            Index           =   7
         End
         Begin VB.Menu mnuCreatures1 
            Caption         =   "[&T] Tiger"
            Index           =   8
         End
         Begin VB.Menu mnuCreatures1 
            Caption         =   "[&H] Head"
            Index           =   9
         End
         Begin VB.Menu mnuCreatures1 
            Caption         =   "[&S] Segment"
            Index           =   10
         End
         Begin VB.Menu mnuCreatures1 
            Caption         =   "[&U] Bullet"
            Index           =   11
         End
         Begin VB.Menu mnuCreatures1 
            Caption         =   "[&A] Star"
            Index           =   12
         End
         Begin VB.Menu mnuCreatures1 
            Caption         =   "[&N] Roton"
            Index           =   13
         End
         Begin VB.Menu mnuCreatures1 
            Caption         =   "[&D] Dragon Pup"
            Index           =   14
         End
         Begin VB.Menu mnuCreatures1 
            Caption         =   "[&I] Pairer"
            Index           =   15
         End
         Begin VB.Menu mnuCreatures1 
            Caption         =   "[&E] Spider"
            Index           =   16
         End
      End
      Begin VB.Menu mnuTerrain 
         Caption         =   "[&3] Elements: Terrain"
         Begin VB.Menu mnuTerrain1 
            Caption         =   "[&W] Water"
            Index           =   0
         End
         Begin VB.Menu mnuTerrain1 
            Caption         =   "[&F] Forest"
            Index           =   1
         End
         Begin VB.Menu mnuTerrain1 
            Caption         =   "[&S] Solid"
            Index           =   2
         End
         Begin VB.Menu mnuTerrain1 
            Caption         =   "[&N] Normal"
            Index           =   3
         End
         Begin VB.Menu mnuTerrain1 
            Caption         =   "[&B] Breakable"
            Index           =   4
         End
         Begin VB.Menu mnuTerrain1 
            Caption         =   "[&O] Boulder"
            Index           =   5
         End
         Begin VB.Menu mnuTerrain1 
            Caption         =   "[&1] Slider (NS)"
            Index           =   6
         End
         Begin VB.Menu mnuTerrain1 
            Caption         =   "[&2] Slider (EW)"
            Index           =   7
         End
         Begin VB.Menu mnuTerrain1 
            Caption         =   "[&A] Fake"
            Index           =   8
         End
         Begin VB.Menu mnuTerrain1 
            Caption         =   "[&I] Invisible"
            Index           =   9
         End
         Begin VB.Menu mnuTerrain1 
            Caption         =   "[&L] Blink Wall"
            Index           =   10
         End
         Begin VB.Menu mnuTerrain1 
            Caption         =   "[&T] Transporter"
            Index           =   11
         End
         Begin VB.Menu mnuTerrain1 
            Caption         =   "[&R] Ricochet"
            Index           =   12
         End
         Begin VB.Menu mnuTerrain1 
            Caption         =   "[&E] Board Edge"
            Index           =   13
         End
         Begin VB.Menu mnuTerrain1 
            Caption         =   "[&M] Monitor"
            Index           =   14
         End
         Begin VB.Menu mnuTerrain1 
            Caption         =   "[&H] Horiz Blink"
            Index           =   15
         End
         Begin VB.Menu mnuTerrain1 
            Caption         =   "[&V] Vert Blink"
            Index           =   16
         End
         Begin VB.Menu mnuTerrain1 
            Caption         =   "[&D] Dead Smiley"
            Index           =   17
         End
         Begin VB.Menu mnuTerrain1 
            Caption         =   "[&X] Empty"
            Index           =   18
         End
         Begin VB.Menu mnuTerrain1 
            Caption         =   "[&Z] Floor"
            Index           =   19
         End
         Begin VB.Menu mnuTerrain1 
            Caption         =   "[&8] Water N"
            Index           =   20
         End
         Begin VB.Menu mnuTerrain1 
            Caption         =   "[&5] Water S"
            Index           =   21
         End
         Begin VB.Menu mnuTerrain1 
            Caption         =   "[&4] Water W"
            Index           =   22
         End
         Begin VB.Menu mnuTerrain1 
            Caption         =   "[&6] Water E"
            Index           =   23
         End
      End
      Begin VB.Menu mnuObjLib 
         Caption         =   "[&5] Object Libraries"
         Visible         =   0   'False
         Begin VB.Menu mnuObjLibList 
            Caption         =   "-"
            Index           =   0
         End
      End
      Begin VB.Menu mnuFunctions3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFunctionsFillAll 
         Caption         =   "[&Z] Fill Entire Board"
      End
      Begin VB.Menu mnuFunctionsBoardInfo 
         Caption         =   "[&I] Board Information"
      End
      Begin VB.Menu mnuFunctionsWorldInfo 
         Caption         =   "[&W] World Information"
      End
   End
   Begin VB.Menu mnuText 
      Caption         =   "T&ext"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsGraphics 
         Caption         =   "[&G] Graphics"
         Begin VB.Menu mnuOptionsGraphics2X 
            Caption         =   "[&2] 2x Graphics Mode (slower)"
         End
         Begin VB.Menu mnuOptionsGrid 
            Caption         =   "[&G] Show Grid"
         End
      End
      Begin VB.Menu mnuOptionsSZTView 
         Caption         =   "[&V] Show player viewport for Super ZZT"
      End
      Begin VB.Menu mnuOptions1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptionsAnimate 
         Caption         =   "[&N] Animate scrolls, conveyors, etc"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuOptionsCoords 
         Caption         =   "[&1] Base one coordinates"
      End
      Begin VB.Menu mnuOptionsContinuous 
         Caption         =   "[&C] Continuous drawing"
      End
   End
   Begin VB.Menu mnuOptimize 
      Caption         =   "O&ptimize"
      Begin VB.Menu mnuOptimizeLocal 
         Caption         =   "[&B] Board (this board only)"
         Begin VB.Menu mnuOptimizeEmptiesLocal 
            Caption         =   "[&1] Convert empties to color 00"
         End
         Begin VB.Menu mnuOptimizeBindsLocal 
            Caption         =   "[&2] Simplify #BINDs"
         End
         Begin VB.Menu mnuOptimizePointersLocal 
            Caption         =   "[&3] Convert stat 'pointers' to 0"
         End
      End
      Begin VB.Menu mnuOptimizeWorldwide 
         Caption         =   "[&W] Board (worldwide)"
         Begin VB.Menu mnuOptimizeEmpties 
            Caption         =   "[&1] Convert empties to color 00"
         End
         Begin VB.Menu mnuOptimizeBinds 
            Caption         =   "[&2] Simplify #BINDs"
         End
         Begin VB.Menu mnuOptimizePointers 
            Caption         =   "[&3] Convert stat 'pointers' to 0"
         End
      End
      Begin VB.Menu mnuOptimizeDelete 
         Caption         =   "[&D] Delete unused boards"
      End
      Begin VB.Menu mnuOptimizeAll 
         Caption         =   "[&A] Perform All Optimizations"
      End
   End
   Begin VB.Menu mnuBoardTitle 
      Caption         =   "Untitled"
      Begin VB.Menu mnuBoards 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpEditor 
         Caption         =   "[&E] Editor"
      End
      Begin VB.Menu mnuHelpLangRef 
         Caption         =   "[&R] Language Reference"
      End
      Begin VB.Menu mnuHelpOOP 
         Caption         =   "[&O] ZZT OOP"
      End
      Begin VB.Menu mnuHelp1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "[&A] About..."
      End
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ZAP zzt editor
'saxxonpike 2oo7-2oo8

Option Explicit

'for putting the board menu on the right side
Private Type MENUITEMINFO
   cbSize As Long
   fMask As Long
   fType As Long
   fState As Long
   wID As Long
   hSubMenu As Long
   hbmpChecked As Long
   hbmpUnchecked As Long
   dwItemData As Long
   dwTypeData As String
   cch As Long
End Type
Private Const MF_STRING = &H0&
Private Const MF_HELP = &H4000&
Private Const MFS_DEFAULT = &H1000&
Private Const MIIM_ID = &H2
Private Const MIIM_SUBMENU = &H4
Private Const MIIM_TYPE = &H10
Private Const MIIM_DATA = &H20
Private Declare Function GetMenu Lib "USER32" (ByVal hWnd As Long) As Long
Private Declare Function GetMenuItemInfo Lib "USER32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function SetMenuItemInfo Lib "USER32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function DrawMenuBar Lib "USER32" (ByVal hWnd As Long) As Long

'shell
Private Declare Function ShellExecute Lib "SHELL32.DLL" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetForegroundWindow Lib "USER32" () As Long

'tools
Private xSelectedElement As Long
Private xSelectedFCol As Long
Private xSelectedBCol As Long
Private xForcedCol As Long
Private xChar As Long
Private xGotFocus As Boolean
Private xBufferElements(0 To 9) As Byte
Private xBufferElements2(0 To 9) As Byte
Private xBufferColors(0 To 9) As Byte
Private xBufferStats(0 To 9) As xtiBufferStats
Private xBufferChars(0 To 9) As Byte
Private bSetStats As Boolean
Private bAutoSetElement As Boolean
Private bIsSetByProgram As Boolean
Private xBoardWidth As Long
Private xBoardHeight As Long
Private bChangingElement As Boolean
Private xFontFile As String
Private xPaletteFile As String
Private xFontCommand As String
Private xPaletteCommand As String
Private bBaseOneCoords As Boolean
Private bAnimateStuff As Boolean

Private xCopySizeX As Long
Private xCopySizeY As Long
Private xCopyObjects() As xtiBufferStats
Private xCopyBuffer() As xtiSelectionType
Private xCopyDispBuffer() As xtiANSIPair
Private xCopyObjectCount As Long

Private xMaxObjects As Long
Private xViewPortX As Long
Private xViewPortY As Long
Private xRBUCount As Long

Private Type xtiANSIPair
    xCol As Byte
    xChar As Byte
End Type

Private Type xtiSelectionType
    xColor As Byte
    xID As Byte
    xObjectRef As Byte
End Type

'editor
Dim XRatio As Double
Dim YRatio As Double
Dim xBlink As Boolean
Dim OldMouseX As Long
Dim OldMouseY As Long
Dim MouseX As Long
Dim MouseY As Long
Dim MousePosX As Long
Dim MousePosY As Long
Dim OldMousePosX As Long
Dim OldMousePosY As Long
Dim ShiftPosX As Long
Dim ShiftPosY As Long
Dim VertMode As Long
Dim ModeType As Long
Dim SelectionX1 As Long
Dim SelectionX2 As Long
Dim SelectionY1 As Long
Dim SelectionY2 As Long
Dim CtrlPosX As Long
Dim CtrlPosY As Long
Dim CtrlMode As Long
Dim tickcount As Long
Dim AnimCount As Long
Dim AnimCountBlink As Long
Dim ModeName As String
Dim FloodFillGrid(0 To 95, 0 To 79) As Byte
Dim TabDrawMode As Boolean
Dim bContinuousDrawing As Boolean
Dim bDrawPoint As Boolean

Const MODE_EDIT = 0
Const MODE_TEXT = 1
Const MODE_OVERVIEW = 2
Const MODE_PASTE = 3
Const MODE_VIEWPORT = 4

Const fontwidth = 8
Const FontHeight = 12
Const PALFilter = "All Openable Palettes (*.ACT, *.PLD)|*.ACT;*.PLD|All Files|*.*"
Const FONTFilter = "Font Mania executable fonts (*.COM)|*.COM|8-pixel width Binary fonts (*.BIN)|*.BIN|All Files|*.*"
Const OFNFilter = "All Openable Files|*.ZZT;*.SZT;*.SAV|ZZT World Files (*.ZZT)|*.ZZT|Super ZZT World Files (*.SZT)|*.SZT|Save Files (*.SAV)|*.SAV|All Files|*.*"
Const OFNFilterS1 = "All World Files|*.ZZT;*.SAV|ZZT World Files (*.ZZT)|*.ZZT|Save Files (*.SAV)|*.SAV|All Files|*.*"
Const OFNFilterS2 = "All World Files|*.SZT;*.SAV|Super ZZT World Files (*.SZT)|*.SZT|Save Files (*.SAV)|*.SAV|All Files|*.*"
Const BRDFilter = "Board Files (*.BRD)|*.BRD|All Files|*.*"
Const BMPFilter = "Bitmap Files (*.BMP)|*.BMP|All Files|*.*"

Private Sub cmdBoardInfo_Click()
    frmBoardInfo.Show
End Sub

Private Sub cmdWorldInfo_Click()
    frmWorldInfo.Show
End Sub

Private Sub Combo1_DropDown()
    MainStatBuffer.xIsFilled = False
End Sub

Private Sub Form_Resize()
    picMain.Width = iUseScale * (FontDCWidth * 60) * 15
    picMain.Height = iUseScale * (FontDCHeight * 25) * 15
    picMain.ScaleWidth = 480
    picMain.ScaleHeight = 400
    HScroll.Width = picMain.Width
    HScroll.Top = picMain.Height
    VScroll.Height = picMain.Height + HScroll.Height
    VScroll.Left = picMain.Width
    Frame1.Width = picMain.Width + VScroll.Width
    Frame1.Top = picMain.Height + HScroll.Height
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If ForegroundWindow = Me.hWnd Then
        If picMain.Enabled = True Then
            picMain.SetFocus
        End If
    End If
End Sub

Private Sub HScroll_Change()
    RefreshBoard
End Sub

Private Sub HScroll_Scroll()
    HScroll_Change
End Sub

Private Sub lblRBU_Click()
    vsRBU = 0
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileLoadFont_Click()
    Dim NewFN As String
    CD_ShowOpen_Save Me.hWnd, OFN_EXPLORER Or OFN_FILEMUSTEXIST, NewFN, "", "", "Import Font", FONTFilter, "", True
    If NewFN <> "" Then
        LoadFont NewFN
    End If
    RefreshBoard
End Sub

Private Sub LoadFont(ByVal NewFN As String)
    Dim f As Long
    Dim s As Integer
    Dim h As Byte
    Dim x As String * 10
    If NewFN = "" Then
        Exit Sub
    End If
    If Dir(NewFN) = "" Then
        NewFN = AppPath + NewFN
        If Dir(NewFN) = "" Then
            Exit Sub
        End If
    End If
    xFontFile = NewFN
    If InStr(xFontFile, "\") > 0 Then
        xFontFile = Mid$(xFontFile, InStrRev(xFontFile, "\") + 1)
    End If
    
    f = FreeFile
    Open NewFN For Binary As #f
    If Right$(UCase$(NewFN), 4) = ".COM" Then
        Get #f, 9, x
        If x <> "FONT MANIA" Then
            If MsgBox("This doesn't appear to be a FONT MANIA font. Attempt anyway?", vbYesNo, "Confirmation") = vbNo Then
                Close #f
                Exit Sub
            End If
        End If
        Get #f, 3, s
        Get #f, 6, h
        If s < LOF(f) And s > 0 And h > 0 Then
            LoadBinaryFont f, s + 1, h + 0, 256
        End If
    Else
        s = LOF(f) / 256
        If (LOF(f) Mod 16 <> 0) Or s > 16 Or s < 8 Then
            MsgBox "This font is either not a fixed-width font, an invalid height, or cannot be processed.", vbCritical, "Oops."
            Close #f
            Exit Sub
        End If
        LoadBinaryFont f, 1, s + 0, 256
    End If
    Close #f
    picMain.Height = Screen.TwipsPerPixelY * 25 * FontDCHeight
    SetEditScale
End Sub

Private Sub mnuFileLoadPalette_Click()
    Dim NewFN As String
    CD_ShowOpen_Save Me.hWnd, OFN_EXPLORER Or OFN_FILEMUSTEXIST, NewFN, "", "", "Load Palette", PALFilter, "", True
    If NewFN <> "" Then
        LoadPalette NewFN
    End If
    RefreshBoard
End Sub

Private Sub LoadPalette(NewFN As String)
    Dim zpal(0 To 191) As Byte
    Dim dpal(0 To 15) As RGB
    Dim f As Long
    Dim c As Long
    
    If NewFN = "" Then
        Exit Sub
    End If
    If Dir(NewFN) = "" Then
        NewFN = AppPath + NewFN
        If Dir(NewFN) = "" Then
            Exit Sub
        End If
    End If
    xPaletteFile = NewFN
    If InStr(xPaletteFile, "\") > 0 Then
        xPaletteFile = Mid$(xPaletteFile, InStrRev(xPaletteFile, "\") + 1)
    End If
    
    f = FreeFile
    If Right$(UCase$(NewFN), 4) = ".PLD" Then
        Open NewFN For Binary As #f
        If LOF(f) <> 192 Then
            If LOF(f) <> 48 Then
                MsgBox "This PLD file cannot be used.", vbCritical, "Oops."
                Close #f
                Exit Sub
            End If
        End If
        Get #f, 1, zpal
        If LOF(f) = 192 Then
            For c = 0 To 7
                SetColor c + 0, zpal(((c + 0) * 3) + 0) * 4, zpal(((c + 0) * 3) + 1) * 4, zpal(((c + 0) * 3) + 2) * 4
                SetColor c + 8, zpal(((c + 56) * 3) + 0) * 4, zpal(((c + 56) * 3) + 1) * 4, zpal(((c + 56) * 3) + 2) * 4
                SetToolColor c + 0, PaletteColor(c + 0)
                SetToolColor c + 8, PaletteColor(c + 8)
            Next c
        Else
            For c = 0 To 15
                SetColor c, zpal((c * 3) + 0), zpal((c * 3) + 1), zpal((c * 3) + 2)
                SetToolColor c, PaletteColor(c)
            Next c
        End If
        For c = 0 To 15
            dpal(c) = PaletteColorRGB(c)
        Next c
        BMP2BRD_CreateLookupTable dpal()
    ElseIf Right$(UCase$(NewFN), 4) = ".ACT" Then
        MsgBox "Support has not been added for the .ACT palette format. (yet.)"
    Else
        MsgBox "Unrecognized file format.", , "Oops."
    End If
    Close #f
End Sub


Private Sub mnuFileNewSZT_Click()
    If MsgBox("Make a new " + "Super ZZT" + " world? All unsaved changes will be lost.", vbYesNo, "Confirmation") = vbYes Then
        CurrentBoard = 0
        World.NewWorld True
        SetupScrollers
        RefreshBoard
        ConstructMenus
    End If
End Sub

Private Sub mnuFileRecent_Click(Index As Integer)
    LoadFile RecentFiles(Index - 1)
End Sub

Private Sub mnuFileTest_Click()
    TestWorld False
End Sub

Private Sub mnuFileTest2_Click()
    TestWorld True
End Sub

Private Sub TestWorld(bUseAddons As Boolean)
    Dim f As Long
    
    'verify that the proper executables do exist and write temp world
    If World.IsSuperZZT Then
        If Dir(AppPath + "SuperZ.exe") = "" Then
            MsgBox "SuperZ.exe could not be found in the application folder."
            Exit Sub
        Else
            World.SetStartBoard CurrentBoard
            World.SaveWorld AppPath + "@zaptest.szt", True
        End If
    Else
        If Dir(AppPath + "ZZT.exe") = "" Then
            MsgBox "ZZT.exe could not be found in the application folder."
            Exit Sub
        Else
            World.SetStartBoard CurrentBoard
            World.SaveWorld AppPath + "@zaptest.zzt", True
        End If
    End If
    
    'now create the batch file
    f = FreeFile
    Open AppPath + "@zapexec.bat" For Output As #f
    Print #f, "@echo off"
    Print #f, "cls"
    If bUseAddons Then
        If xFontFile <> "" Or xPaletteFile <> "" Then
            'Print #f, "echo " + String$(60, 205)
            'Print #f, "echo To properly see fonts and palettes, make sure to press"
            'Print #f, "echo ALT+ENTER at this time. Please note that Windows XP"
            'Print #f, "echo does tend to have problems with font and palette changes."
            'Print #f, "echo " + String$(60, 205)
            'Print #f, "pause"
        End If
        If xFontCommand <> "" And xFontFile <> "" Then
            Print #f, Replace(UCase$(xFontCommand), "%N", xFontFile)
        End If
        If xPaletteCommand <> "" And xPaletteFile <> "" Then
            Print #f, Replace(UCase$(xPaletteCommand), "%N", xFontFile)
        End If
    End If
    If World.IsSuperZZT Then
        Print #f, "superz.exe @zaptest.szt"
    Else
        Print #f, "zzt.exe @zaptest.zzt"
    End If
    Close #f
    
    ShellExecute Me.hWnd, vbNullString, "command.com", "/c @zapexec.bat", AppPath, 1

End Sub

Private Sub mnuFileTransferII_Click()
    Dim NewFN As String
    CD_ShowOpen_Save Me.hWnd, OFN_EXPLORER Or OFN_FILEMUSTEXIST, NewFN, "", "", "Generate BRD from 60x25 Image", BMPFilter, "", True
    If NewFN <> "" Then
        BMP2BRD_LoadToBoard CurrentBoard, NewFN
    End If
    RefreshBoard
End Sub

Private Sub mnuFileUnloadFont_Click()
    xFontFile = ""
    RestoreFontDC
    picMain.Height = Screen.TwipsPerPixelY * 25 * FontDCHeight
    SetEditScale
    RefreshBoard
End Sub

Private Sub mnuFileUnloadPalette_Click()
    xPaletteFile = ""
    SetupPalette
    RefreshBoard
End Sub

Private Sub mnuFunctionsBoardInfo_Click()
    frmBoardInfo.Visible = True
End Sub

Private Sub mnuFunctionsCopyRect_Click()
    GetSelection False
End Sub

Private Sub mnuFunctionsCutRect_Click()
    GetSelection True
End Sub

Private Sub mnuFunctionsEditPattern_Click()
    Load frmFillStyle
    frmFillStyle.Show
End Sub

Private Sub mnuFunctionsFillPattern_Click()
    Dim x1 As Long
    Dim y1 As Long
    bSkipStats = True
    If Not bFillPattern Then
        mnuFunctionsEditPattern_Click
        Exit Sub
    Else
        For x1 = SelectionX1 To SelectionX2
            For y1 = SelectionY1 To SelectionY2
                If x1 >= 0 And y1 >= 0 And x1 < xBoardWidth And y1 < xBoardHeight Then
                    If FloodFillPattern((x1 - SelectionX1) Mod lFillSizeX, (y1 - SelectionY1) Mod lFillSizeY) Then
                        SetTile x1, y1
                    End If
                End If
            Next y1
        Next x1
        RefreshBoard
    End If
    bSkipStats = False
End Sub

Private Sub mnuFunctionsFloodPattern_Click()
    If Not bFillPattern Then
        mnuFunctionsEditPattern_Click
        Exit Sub
    Else
        FloodFill True 'use pattern
    End If
End Sub

Private Sub mnuFunctionsMoveRect_Click()
    GetSelection True
    If xCopySizeX > 0 And xCopySizeY > 0 Then
        PutSelection
    End If
End Sub

Private Sub mnuFunctionsPasteRect_Click()
    PutSelection
End Sub

Private Sub mnuFunctionsWorldInfo_Click()
    frmWorldInfo.Visible = True
End Sub

Private Sub mnuHelpEditor_Click()
    frmHelp.ShowHelp "EDITOR"
End Sub

Private Sub mnuHelpLangRef_Click()
    frmHelp.ShowHelp "LANGREF"
End Sub

Private Sub mnuHelpOOP_Click()
    frmHelp.ShowHelp "LANG"
End Sub

Private Sub mnuOptimizeAll_Click()
    Dim z As Long
    mnuOptimizeDelete_Click
    For z = 0 To World.BoardCount
        OptimizeBinds z
        OptimizeEmpties z
        OptimizePointers z
    Next z
    MsgBox "All optimizations have been applied.", vbInformation, "Success"
End Sub

Private Sub mnuOptimizeBinds_Click()
    Dim z As Long
    Dim a As Long
    For z = 0 To World.BoardCount
        a = a + OptimizeBinds(z)
    Next z
    MsgBox "Bytes saved: " + CStr(a), vbInformation, "Results"
End Sub

Private Sub OptimizeEmpties(boardnum As Long)
    World.ChangeTiles boardnum, E_Empty, 256, 256, 0
End Sub

Private Sub OptimizePointers(boardnum As Long)
    'the only real point to this is to make it so that the ZZT file will
    'compress better in an archived format
    Dim x As Long
    For x = 1 To World.StatCount(boardnum)
        World.SetObjectPointer boardnum, x, 0
    Next x
End Sub

Private Function OptimizeBinds(boardnum As Long) As Long
    Dim x As Long
    Dim y As String
    Dim a As Long
    Dim xCount As Long
    Dim ycount As Long
    Dim objectlist() As String
    Dim objectref() As Long
    
    'build object name list
    a = 0
    ReDim objectlist(0) As String
    ReDim objectref(0) As Long
    For x = 1 To World.StatCount(boardnum)
        y = World.ObjectName(boardnum, x)
        If y <> "" Then
            a = a + 1
            ReDim Preserve objectlist(0 To a) As String
            ReDim Preserve objectref(0 To a) As Long
            objectlist(a) = UCase$(y)
            objectref(a) = x
        End If
    Next x
    
    'now assign binds if necessary (and only if we have named objects)
    'by removing all its code and setting the bind parameter
    If UBound(objectlist) > 0 Then
        For x = 1 To World.StatCount(boardnum)
            y = UCase$(World.ObjectOOP(boardnum, x))
            If Left$(y, 6) = "#BIND " Then
                If InStr(y, Chr$(13)) > 0 Then
                    y = Left$(y, InStr(y, Chr$(13)) - 1)
                End If
                For a = 1 To UBound(objectlist)
                    If objectlist(a) = Mid$(y, 7) Then
                        'clear code and set the reference
                        OptimizeBinds = OptimizeBinds + World.ObjectLength(boardnum, x)
                        World.SetObjectOOP boardnum, x, ""
                        World.SetObjectLength boardnum, x, (-objectref(a))
                        Exit For
                    End If
                Next a
            End If
        Next x
    End If
End Function

Private Sub mnuOptimizeBindsLocal_Click()
    OptimizeBinds CurrentBoard
End Sub

Private Sub mnuOptimizeDelete_Click()
    Dim bBoardUsed(0 To 255) As Byte
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
    z = World.BoardCount
    bBoardUsed(World.StartBoard) = 1
    For x = z + 1 To 255
        bBoardUsed(x) = 3
    Next x
    Do
        f = False
        For x = 0 To z
            If bBoardUsed(x) = 1 Then
                World.GetBoardInfo x, sh, dk, za, bNo, bSo, bWe, bEa, tm
                If bNo <> 0 And bBoardUsed(bNo) = 0 Then bBoardUsed(bNo) = 1: f = True
                If bSo <> 0 And bBoardUsed(bSo) = 0 Then bBoardUsed(bSo) = 1: f = True
                If bEa <> 0 And bBoardUsed(bEa) = 0 Then bBoardUsed(bEa) = 1: f = True
                If bWe <> 0 And bBoardUsed(bWe) = 0 Then bBoardUsed(bWe) = 1: f = True
                For y = 0 To World.StatCount(x)
                    World.GetObjectInfo1 x, y, xs, ys, cy, p1, p2, p3
                    World.GetStatLocation x, y, locx, locy
                    If locx > 0 And locy > 0 Then
                        If World.BoardID(x, locx - 1, locy - 1) = E_Passage Then
                            If bBoardUsed(p3) = 0 Then
                                bBoardUsed(p3) = 1
                                f = True
                            End If
                        End If
                    End If
                Next y
                'x = x - 1
                bBoardUsed(x) = 2
            End If
        Next x
    Loop While f
    For x = 1 To z
        If bBoardUsed(x) = 0 Then
            For y = x To z - 1
                bBoardUsed(y) = bBoardUsed(y + 1)
            Next y
            bc = bc + 1
            World.DeleteBoard x
            If CurrentBoard >= x Then
                CurrentBoard = CurrentBoard - 1
            End If
            z = World.BoardCount
            If x > z Then
                Exit For
            End If
            x = x - 1
        End If
    Next x
    MsgBox "Unused boards deleted: " + CStr(bc), vbInformation, "Delete Unused Boards"
    'CurrentBoard = 0
    ConstructMenus
    RefreshBoard
End Sub

Private Sub mnuOptimizeEmptiesLocal_Click()
    OptimizeEmpties CurrentBoard
End Sub

Private Sub mnuOptimizePointers_Click()
    Dim z As Long
    For z = 0 To World.BoardCount
        OptimizePointers z
    Next z
End Sub

Private Sub mnuOptimizePointersLocal_Click()
    OptimizePointers CurrentBoard
End Sub

Private Sub mnuOptionsAnimate_Click()
    mnuOptionsAnimate.Checked = Not mnuOptionsAnimate.Checked
    bAnimateStuff = mnuOptionsAnimate.Checked
End Sub

Private Sub mnuOptionsContinuous_Click()
    mnuOptionsContinuous.Checked = Not mnuOptionsContinuous.Checked
    bContinuousDrawing = mnuOptionsContinuous.Checked
End Sub

Private Sub mnuOptionsCoords_Click()
    mnuOptionsCoords.Checked = Not mnuOptionsCoords.Checked
    bBaseOneCoords = mnuOptionsCoords.Checked
End Sub

Private Sub mnuOptionsSZTView_Click()
    mnuOptionsSZTView.Checked = (Not mnuOptionsSZTView.Checked)
End Sub

Private Sub mnuPasteFlipH_Click()
    FlipSelectionH
End Sub

Private Sub mnuPasteFlipV_Click()
    FlipSelectionV
End Sub

Private Sub picMain_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode >= 16 And KeyCode <= 18 Then
        Exit Sub 'ignore plain shift/ctrl/alt
    End If
    If Shift = 4 Then
        Dim x As Long
        Dim a As Byte
        Dim b(0 To 3) As Byte
        Dim y As Integer
        World.GetBoardInfo CurrentBoard, a, a, a, b(0), b(1), b(2), b(3), y
    End If
    Debug.Print "frmedit.form_keydown", KeyCode, Shift
    Select Case ModeType
        'editmode
        Case MODE_EDIT
            Select Case KeyCode
                Case Asc("C")
                    If Shift = 3 Then
                        mnuFunctionsCopy_Click
                    ElseIf Shift = 2 Then
                        GetSelection False
                    ElseIf Shift = 1 Then
                        NextColorB
                    ElseIf Shift = 0 Then
                        NextColorF
                    End If
                    ClearColors
                Case Asc("D")
                    If Shift = 0 Then
                        chkDefaultColor.Value = Abs(chkDefaultColor.Value - 1)
                    End If
                Case Asc("V")
                    If Shift = 3 Then
                        SetTile MouseX, MouseY
                    ElseIf Shift = 2 Then
                        PutSelection
                    ElseIf Shift = 0 Then
                        chkBlinks.Value = Abs(chkBlinks.Value - 1)
                    End If
                Case Asc("T")
                    If Shift = 0 Then
                        PopupMenu mnuFileTransfer
                    ElseIf Shift = 2 Then
                        mnuFileTest_Click
                    End If
                Case Asc("Z")
                    If Shift = 0 Then
                        If MsgBox("Really clear board?", vbYesNo) = vbYes Then
                            World.ResetBoard CurrentBoard
                            RefreshBoard
                        End If
                    End If
                Case Asc("L")
                    If Shift = 0 Then
                        mnuFileLoad_Click
                    End If
                Case Asc("X")
                    If Shift = 3 Then 'move
                        GetSelection True
                        If xCopySizeX > 0 And xCopySizeY > 0 Then
                            PutSelection
                        End If
                    ElseIf Shift = 2 Then 'cut
                        GetSelection True
                    ElseIf Shift = 0 Then
                        mnuFunctionsFlood_Click
                    End If
                Case Asc("S")
                    If Shift = 0 Then
                        mnuFileSave_Click
                    End If
                Case Asc("B")
                    If Shift = 0 Then 'board select
                        PopupMenu mnuBoardTitle
                    ElseIf Shift = 1 Then 'go to start board
                        SetCurrentBoard World.StartBoard
                        RefreshBoard
                    End If
                Case Asc("I")
                    If Shift = 0 Then
                        frmBoardInfo.Show
                    ElseIf Shift = 1 Then
                        frmWorldInfo.Show
                    End If
                Case 13 'enter
                    If Shift = 0 Then
                        PopupMenu mnuFunctions
                    End If
                Case 27 'esc
                    mnuFunctionsDeselect_Click
                    RefreshBoard
                Case 32 'space
                    SetTile MousePosX, MousePosY
                Case 112 'f1: items
                    If Shift = 0 Then
                        PopupMenu mnuItems
                    End If
                Case 113 'f2: creatures
                    If Shift = 0 Then
                        PopupMenu mnuCreatures
                    End If
                Case 114 'f3: terrain
                    If Shift = 0 Then
                        PopupMenu mnuTerrain
                    End If
                Case 116 'f5: object libraries
                    If Shift = 0 Then
                        PopupMenu mnuObjLib
                    End If
                Case 192 'tilde
                    If Shift = 0 Then
                        RefreshBoard
                    End If
                Case 33 'page up
                    If Shift = 0 Then
                        If CurrentBoard > 0 Then
                            SetCurrentBoard CurrentBoard - 1
                        Else
                            SetCurrentBoard World.BoardCount + 0
                        End If
                        RefreshBoard
                    End If
                Case 34 'page down
                    If Shift = 0 Then
                        If CurrentBoard < World.BoardCount Then
                            SetCurrentBoard CurrentBoard + 1
                        Else
                            SetCurrentBoard 0
                        End If
                        RefreshBoard
                    End If
                Case 37
                    If Shift = 4 Then 'board switch west link
                        If BoardInfo.xWest > 0 Then
                            SetCurrentBoard BoardInfo.xWest + 0
                            RefreshBoard
                        End If
                    ElseIf Shift = 0 Then
                        MouseX = MouseX - 1
                    End If
                Case 38
                    If Shift = 4 Then 'board switch north link
                        If BoardInfo.xNorth > 0 Then
                            SetCurrentBoard BoardInfo.xNorth + 0
                            RefreshBoard
                        End If
                    ElseIf Shift = 0 Then
                        MouseY = MouseY - 1
                    End If
                Case 39
                    If Shift = 4 Then 'board switch east link
                        If BoardInfo.xEast > 0 Then
                            SetCurrentBoard BoardInfo.xEast + 0
                            RefreshBoard
                        End If
                    ElseIf Shift = 0 Then
                        MouseX = MouseX + 1
                    End If
                Case 40
                    If Shift = 4 Then 'board switch south link
                        If BoardInfo.xSouth > 0 Then
                            SetCurrentBoard BoardInfo.xSouth + 0
                            RefreshBoard
                        End If
                    ElseIf Shift = 0 Then
                        MouseY = MouseY + 1
                    End If
                Case Asc("N")
                    If Shift = 0 Then
                        mnuFileNew_Click
                    ElseIf Shift = 1 Then
                        mnuFileNewSZT_Click
                    End If
                Case Asc("H")
                    If Shift = 0 Then
                        mnuHelpEditor_Click
                    End If
                Case 9 'tab
                    TabDrawMode = Not TabDrawMode
                    DoMode
            End Select
            If MouseX <> OldMouseX Or MouseY <> OldMouseY Then
                If MouseX = 60 Then
                    MouseX = 59
                    If World.IsSuperZZT Then
                        If HScroll.Value < HScroll.Max Then
                            HScroll.Value = HScroll.Value + 1
                        End If
                    End If
                End If
                If MouseY = 25 Then
                    MouseY = 24
                    If World.IsSuperZZT Then
                        If VScroll.Value < VScroll.Max Then
                            VScroll.Value = VScroll.Value + 1
                        End If
                    End If
                End If
                If MouseX = -1 Then
                    MouseX = 0
                    If World.IsSuperZZT Then
                        If HScroll.Value > HScroll.Min Then
                            HScroll.Value = HScroll.Value - 1
                        End If
                    End If
                End If
                If MouseY = -1 Then
                    MouseY = 0
                    If World.IsSuperZZT Then
                        If VScroll.Value > VScroll.Min Then
                            VScroll.Value = VScroll.Value - 1
                        End If
                    End If
                End If
                RefreshCursor
                OldMouseX = MouseX
                OldMouseY = MouseY
            End If
            If TabDrawMode = True Then
                SetTile MousePosX, MousePosY
            End If
        Case MODE_PASTE
            Select Case KeyCode
                Case Asc("H")
                    FlipSelectionH
                Case Asc("V")
                    FlipSelectionV
            End Select
            RefreshCursor
        'textmode
        Case MODE_TEXT
            Debug.Print KeyCode, Shift
            Select Case KeyCode
                Case 8, 37 'left
                    MouseX = MouseX - 1
                Case 38 'up
                    MouseY = MouseY - 1
                Case 39 'right
                    MouseX = MouseX + 1
                Case 40 'down
                    MouseY = MouseY + 1
                Case 9 'tab also cancels text edit
                    ModeType = MODE_EDIT
                    DoMode
            End Select
            If MouseX <> OldMouseX Or MouseY <> OldMouseY Then
                If MouseX = 60 Then
                    MouseX = 0
                    MouseY = MouseY + 1
                End If
                If MouseY = 25 Then
                    MouseY = 0
                End If
                If MouseX = -1 Then
                    MouseX = 59
                    MouseY = MouseY - 1
                End If
                If MouseY = -1 Then
                    MouseY = 24
                End If
                RefreshCursor
                OldMouseX = MouseX
                OldMouseY = MouseY
            End If
    End Select
    
    'universals
    Select Case KeyCode
        Case 115 'f4: text
            If ModeType = MODE_EDIT Then
                ModeType = MODE_TEXT
            Else
                ModeType = MODE_EDIT
            End If
            DoMode
        Case 46, 8 'delete & backspace key
            If ModeType = MODE_EDIT Or ModeType = MODE_TEXT Then
                SetTile MouseX + HScroll, MouseY + VScroll, 0, 0, True
                RefreshTile MouseX + HScroll, MouseY + VScroll
            End If
        Case 27
            If ModeType <> MODE_EDIT Then
                ModeType = MODE_EDIT
                DoMode
                RefreshBoard
            End If
    End Select
End Sub

Private Sub picMain_KeyPress(KeyAscii As Integer)
    'the difference between this and KeyDown is, this one will return the
    'ascii value that the user wanted to type and not the key code. very handy
    'for text entry
    'also note: keydown is called before keypress
    Dim c As Long
    If ModeType = MODE_TEXT And KeyAscii > 31 Then
        c = 46 + (SelectedColor And 7)
        If c = 46 Then c = 53
        SetTile MouseX + HScroll.Value, MouseY + VScroll.Value, c + 0, KeyAscii + 0
        MouseX = MouseX + 1
        If MouseX = 60 Then
            MouseX = 0
            MouseY = MouseY + 1
        End If
        If MouseY = 25 Then
            MouseY = 0
        End If
        RefreshCursor
    End If
End Sub

Private Sub Form_Load()
    
    ExWidth = (Me.Width - Me.ScaleWidth) + VScroll.Width
    ExHeight = (Me.Height - Me.ScaleHeight) + Frame1.Height + HScroll.Height
    'SetEditScale 1
    DoMode
    mnuFunctionsDeselect_Click
    RefreshBoard
    ConstructMenus
    
    SetElementByName "normal"
    xSelectedFCol = 15
    xSelectedBCol = 0
    
    'default to ZZT mode
    xBoardHeight = 25
    xBoardWidth = 60
    
    SelectionX1 = -1
    SelectionY1 = -1
    SelectionX2 = -1
    SelectionY2 = -1
    
    bAnimateStuff = True
    xMaxObjects = 150
    
    DoEvents
End Sub

Public Sub LoadConfig()
    Dim f As Long
    Dim s As String
    Dim t As String
    Dim d As Long
    
    'shortcut defaults
    xBufferElements2(0) = E_Solid
    xBufferElements2(1) = E_Normal
    xBufferElements2(2) = E_Breakable
    xBufferElements2(3) = E_Water
    xBufferElements2(4) = E_Empty
    xBufferElements2(5) = E_Invisible
    xBufferElements2(6) = E_Fake
    xBufferElements2(7) = E_Forest
    xBufferElements2(8) = E_Line
    xBufferElements2(9) = E_Player
    
    If Dir(AppPath + "zap.cfg") <> "" Then
        f = FreeFile
        Open AppPath + "zap.cfg" For Input As #f
        Do While Not EOF(f)
            Line Input #f, s
            'remove comments
            If InStr(s, "//") > 0 Then
                s = Left$(s, InStr(s, "//") - 1)
            End If
            If InStr(s, "=") > 0 Then
                d = InStr(s, "=")
                t = Mid$(s, d + 1)
                Select Case UCase$(Left$(s, d - 1))
                    Case "SHORTCUT0": xBufferElements2(0) = ElementNumFromName(t)
                    Case "SHORTCUT1": xBufferElements2(1) = ElementNumFromName(t)
                    Case "SHORTCUT2": xBufferElements2(2) = ElementNumFromName(t)
                    Case "SHORTCUT3": xBufferElements2(3) = ElementNumFromName(t)
                    Case "SHORTCUT4": xBufferElements2(4) = ElementNumFromName(t)
                    Case "SHORTCUT5": xBufferElements2(5) = ElementNumFromName(t)
                    Case "SHORTCUT6": xBufferElements2(6) = ElementNumFromName(t)
                    Case "SHORTCUT7": xBufferElements2(7) = ElementNumFromName(t)
                    Case "SHORTCUT8": xBufferElements2(8) = ElementNumFromName(t)
                    Case "SHORTCUT9": xBufferElements2(9) = ElementNumFromName(t)
                    Case "FONT": LoadFont t
                    Case "PALETTE": LoadPalette t
                    Case "FONTCOMMAND": xFontCommand = t
                    Case "PALETTECOMMAND": xPaletteCommand = t
                    Case "BASEONE": bBaseOneCoords = CBool(t): mnuOptionsCoords.Checked = bBaseOneCoords
                    Case "CONTINUOUSLINES": bContinuousDrawing = CBool(t): mnuOptionsContinuous.Checked = bContinuousDrawing
                    Case "DOUBLESIZE": If CBool(t) = True Then mnuOptionsGraphics2X_Click
                    Case "ANIMATESTUFF": If CBool(t) = False Then mnuOptionsAnimate_Click
                End Select
            End If
        Loop
    End If
    
    RefreshBoard
End Sub

Sub RefreshCursor()
    Dim x As Long
    If ModeType = MODE_PASTE Then
        'special cursor handling for paste/move mode

    Else
        If mnuOptionsGrid.Checked = True Then
            RefreshBoard
            picMain.Line (0, MouseY * 16)-(480, MouseY * 16), &H666666
            picMain.Line (0, (MouseY + 1) * 16)-(480, (MouseY + 1) * 16), &H666666
            picMain.Line (MouseX * 8, 0)-(MouseX * 8, 400), &H666666
            picMain.Line ((MouseX + 1) * 8, 0)-((MouseX + 1) * 8, 400), &H666666
        Else
            RenderCharDC OldMouseX, OldMouseY, iUseScale
            RenderCharDC MouseX, MouseY, iUseScale
        End If
        picMain.Line (MouseX * 8, MouseY * 16)-((MouseX * 8) + 7, (MouseY * 16) + 15), vbWhite, B
        If SelectionX1 > -1 Then
            If CtrlMode = 1 Then
                picMain.Line ((SelectionX1) * 8, (SelectionY1) * 16)-(((SelectionX2) * 8) + 7, ((SelectionY2) * 16) + 15), vbWhite, B
            Else
                picMain.Line ((SelectionX1 - HScroll.Value) * 8, (SelectionY1 - VScroll.Value) * 16)-(((SelectionX2 - HScroll.Value) * 8) + 7, ((SelectionY2 - VScroll.Value) * 16) + 15), vbYellow, B
            End If
        End If
    End If
    
    OldMousePosX = MousePosX
    OldMousePosY = MousePosY
    'MousePosX = MouseX + HScroll.Value
    'MousePosY = MouseY + VScroll.Value
    If bBaseOneCoords Then
        lblCoord = "(" + CStr(MousePosX + 1) + ", " + CStr(MousePosY + 1) + ")"
    Else
        lblCoord = "(" + CStr(MousePosX) + ", " + CStr(MousePosY) + ")"
    End If
    x = World.ObjectAt(CurrentBoard, MousePosX, MousePosY)
    If x >= 0 Then
        lblCoord = lblCoord + " [" + CStr(x) + "]"
    End If
End Sub

Private Sub DrawBlock(x As Long, y As Long)
    'draw all the tiles on the OUTSIDE of the coordinate
    DrawChar x + 1, y
    DrawChar x - 1, y
    DrawChar x, y + 1
    DrawChar x, y - 1
End Sub

Private Sub picMain_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 16 Or KeyCode = 17 Then
        picMain_MouseMove 0, 64, MouseX * 8, MouseY * 16
    End If
End Sub

Private Sub picMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If ForegroundWindow <> Me.hWnd Then
        Exit Sub
    End If
    If ModeType = MODE_EDIT Then
        If Button = 1 Then
            bDrawPoint = True
        End If
        If Button = 2 Then
            PopupMenu mnuFunctions
        End If
        If Button = 4 Then
            mnuFunctionsCopy_Click
        End If
        If Shift = 1 And ShiftPosX = -1 Then
            ShiftPosX = MouseX
            ShiftPosY = MouseY
        End If
        If Shift = 0 And ShiftPosX <> -1 Then
            ShiftPosX = -1
            ShiftPosY = -1
            VertMode = 0
        End If
    ElseIf ModeType = MODE_TEXT Then
        If Button = 1 Then
            MouseX = x \ 8
            MouseY = y \ 16
            OldMouseX = MouseX
            OldMouseY = MouseY
            RefreshCursor
        End If
    ElseIf ModeType = MODE_PASTE Then
        If Button = 1 Then
            PasteSelection MousePosX, MousePosY
        ElseIf Button = 2 Then
            PopupMenu mnuPaste
        End If
    End If
End Sub

Private Sub DrawChar(ByVal x As Long, ByVal y As Long)
    Dim iParam As Long
    Dim xChar As Byte
    Dim xCol As Byte
    Dim xType As Byte
    Dim AnimCount2 As Long
    Dim DispX As Long
    Dim DispY As Long
    DispX = x
    DispY = y
    x = x + HScroll.Value
    y = y + VScroll.Value
    If Not (x >= 0 And x < xBoardWidth And y >= 0 And y < xBoardHeight) Then
        Exit Sub
    End If
    If Not (DispX >= 0 And DispX < 60 And DispY >= 0 And DispY < 25) Then
        Exit Sub
    End If
    xType = World.BoardID(CurrentBoard, x, y)
    xCol = World.BoardCol(CurrentBoard, x, y)
    If xType = E_Empty Then
        xCol = 0
    End If
    If IsLinkedChar(xType + 0) Then
        iParam = 15 'assume a line connection until proven otherwise
        If x > 0 Then If World.BoardID(CurrentBoard, x - 1, y) <> xType Then iParam = iParam - 4 'east
        If x < (xBoardWidth - 1) Then If World.BoardID(CurrentBoard, x + 1, y) <> xType Then iParam = iParam - 8 'west
        If y > 0 Then If World.BoardID(CurrentBoard, x, y - 1) <> xType Then iParam = iParam - 1 'north
        If y < (xBoardHeight - 1) Then If World.BoardID(CurrentBoard, x, y + 1) <> xType Then iParam = iParam - 2 'south
    End If
    'also process special colors or characters at this time
    xChar = DefaultChar(xType + 0, xCol + 0, iParam)
    If xType = E_Object Or xType = E_Pusher Then
        xChar = World.ObjectCharAt(CurrentBoard, x, y, 0)
    End If
    If xType = E_Counter Or xType = E_Star Or xType = E_SpinningGun Then
        'animates every 2 ticks
        AnimCount2 = (World.ObjectAt(CurrentBoard, x, y) Mod 2) + AnimCount
        xChar = World.ObjectCharAt(CurrentBoard, x, y, ((AnimCount2 \ 2) Mod 4) + 1)
    End If
    If xType = E_Transporter Then
        If World.ObjectCycle(CurrentBoard, World.ObjectAt(CurrentBoard, x, y)) > 0 Then
            AnimCount2 = (World.ObjectAt(CurrentBoard, x, y) Mod World.ObjectCycle(CurrentBoard, World.ObjectAt(CurrentBoard, x, y))) + AnimCount
            xChar = World.ObjectCharAt(CurrentBoard, x, y, ((AnimCount2 \ World.ObjectCycle(CurrentBoard, World.ObjectAt(CurrentBoard, x, y))) Mod 4) + 1)
        End If
    End If
    If xType = E_Clockwise Then
        'animates every 3 ticks
        AnimCount2 = (World.ObjectAt(CurrentBoard, x, y) Mod 3) + AnimCount
        xChar = World.ObjectCharAt(CurrentBoard, x, y, ((AnimCount2 \ 3) Mod 4) + 1)
    End If
    If xType = E_DragonPup Then
        iParam = World.ObjectCycle(CurrentBoard, World.ObjectAt(CurrentBoard, x, y)) + 0
        AnimCount2 = World.ObjectAt(CurrentBoard, x, y) + ((AnimCount \ iParam) * iParam)
        xChar = World.ObjectCharAt(CurrentBoard, x, y, (AnimCount2 Mod 4) + 1)
    End If
    If xType = E_Star Or xType = E_Scroll Then
        xCol = DefaultColor(xType + 0, World.BoardCol(CurrentBoard, x, y), AnimCount)
    End If
    If xType >= E_BlueText And xType <= E_WhiteText Then
        xCol = DefaultColor(xType + 0)
    End If
    If xType = E_Stone Then
        xCol = DefaultColor(xType + 0)
    End If
    If xType = E_Player Then
        If World.ObjectAt(CurrentBoard, x, y) >= 0 Then
            xCol = &H1F 'because stat players are ALWAYS white on blue
        End If
    End If
    
    SetCharB DispX, DispY, xChar, xCol
End Sub

Private Sub picMain_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Dim a As Long
    Dim b As Long
    Dim z As Long
    
    x = Int(x)
    y = Int(y)
    
    If ForegroundWindow <> Me.hWnd Then
        Exit Sub
    End If
    
    If picMain.Enabled = True Then
        picMain.SetFocus
    End If
    
    'do not move cursors on text mode
    If ModeType = MODE_TEXT Then
        Exit Sub
    End If
    
    'if we didn't change characters, don't bother updating
    If (x \ 8) = OldMouseX And (y \ 16) = OldMouseY And Shift = 0 Then
        Exit Sub
    End If
    
    MouseX = (x \ 8)
    MouseY = (y \ 16)
    OldMousePosX = MousePosX
    OldMousePosY = MousePosY
    MousePosX = (MouseX + HScroll.Value)
    MousePosY = (MouseY + VScroll.Value)
    
    Select Case ModeType
        Case MODE_EDIT
            'SHIFT ***
            
            If (Shift And 1) = 0 Then
                VertMode = 0
            End If
            If (Shift And 1) And ShiftPosX = -1 Then
                ShiftPosX = MouseX
                ShiftPosY = MouseY
            End If
            If (Shift And 1) And ShiftPosX <> -1 Then
                If MouseX <> ShiftPosX And VertMode = 0 Then
                    VertMode = 1
                    Debug.Print "Horizontal"
                End If
                If MouseY <> ShiftPosY And VertMode = 0 Then
                    VertMode = 2
                    Debug.Print "Vertical"
                End If
                If VertMode = 1 Then
                    MouseY = ShiftPosY
                    MousePosY = ShiftPosY + VScroll.Value
                End If
                If VertMode = 2 Then
                    MouseX = ShiftPosX
                    MousePosX = ShiftPosX + HScroll.Value
                End If
            End If
            If ((Shift And 1) = 0) And ShiftPosX <> -1 Then
                ShiftPosX = -1
                ShiftPosY = -1
            End If
            
            ' CTRL ***
            
            If (Shift And 2) And CtrlMode <> 1 Then
                'pressed ctrl
                CtrlPosX = MouseX
                CtrlPosY = MouseY
                SelectionX1 = MouseX
                SelectionY1 = MouseY
                SelectionX2 = MouseX
                SelectionY2 = MouseY
                CtrlMode = 1
            End If
            If (Shift And 2) = 0 And CtrlMode = 1 Then
                'released ctrl
                CtrlMode = 2
                SelectionX1 = SelectionX1 + HScroll.Value
                SelectionX2 = SelectionX2 + HScroll.Value
                SelectionY1 = SelectionY1 + VScroll.Value
                SelectionY2 = SelectionY2 + VScroll.Value
                RefreshBoard
            End If
            If (Shift And 2) And CtrlMode = 1 Then
                'move selection
                SelectionX1 = CtrlPosX
                SelectionY1 = CtrlPosY
                SelectionX2 = MouseX
                SelectionY2 = MouseY
                If SelectionX2 < SelectionX1 Then
                    a = SelectionX1
                    SelectionX1 = SelectionX2
                    SelectionX2 = a
                End If
                If SelectionY2 < SelectionY1 Then
                    a = SelectionY1
                    SelectionY1 = SelectionY2
                    SelectionY2 = a
                End If
                Debug.Print SelectionX1, SelectionY1, SelectionX2, SelectionY2
            End If
            If ((Shift And 2) = 0) And CtrlMode = 1 Then
                If SelectionX1 = CtrlPosX And SelectionX2 = CtrlPosX Then
                    If SelectionY1 = CtrlPosY And SelectionY2 = CtrlPosY Then
                        'no selection made
                        CtrlPosX = -1
                        CtrlPosY = -1
                        SelectionX1 = -1
                        SelectionX2 = -1
                        SelectionY1 = -1
                        SelectionY2 = -1
                        CtrlMode = 0
                    End If
                End If
            End If
            
            If Button = 1 Or TabDrawMode = True Then
                DrawPoint
            End If
        
        Case MODE_PASTE
            TabDrawMode = False
            CtrlMode = 0
            RefreshPaste
    End Select
End Sub

Sub RefreshPaste()
    Dim a As Long
    Dim b As Long
    Dim z As Long
    'Exit Sub
    'erase old location
    For b = 0 To xCopySizeY - 1
        If OldMouseY + b < 25 Then
            For a = 0 To xCopySizeX - 1
                If OldMouseX + a < 60 Then
                    DrawChar OldMouseX + a, OldMouseY + b
                End If
            Next a
        End If
    Next b
    
    z = 1
    For b = 0 To xCopySizeY - 1
        If MouseY + b < 25 Then
            For a = 0 To xCopySizeX - 1
                z = (a + (b * xCopySizeX) + 1)
                If MouseX + a < 60 Then
                    SetCharB MouseX + a, MouseY + b, xCopyDispBuffer(z).xChar, xCopyDispBuffer(z).xCol
                End If
            Next a
        End If
    Next b
    DrawMethod
    picMain.Line (MouseX * 8, MouseY * 16)-((MouseX + xCopySizeX) * 8, (MouseY + xCopySizeY) * 16), vbYellow, B
    OldMouseX = MouseX
    OldMouseY = MouseY
End Sub

Private Sub RefreshTile(x1 As Long, y1 As Long)
    If ModeType = MODE_PASTE Then
        If Not ((x1 < MousePosX) Or (y1 < MousePosY) Or (x1 > (MousePosX + xCopySizeX)) Or (y1 > (MousePosY + xCopySizeY))) Then
            Exit Sub
        End If
    End If
    RenderCharDC x1 - HScroll.Value, y1 - VScroll.Value, iUseScale
End Sub

Private Sub SetTileLine(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long)
    'line drawing routine converted from:
    'http://www.cprogramming.com/tutorial/tut3.html
    If x1 = x2 And y1 = y2 Then
        SetTile x1, y1
        RefreshTile x1, y1
        Exit Sub
    End If
    Dim u As Long
    Dim s As Long
    Dim v As Long
    Dim d1x As Long
    Dim d1y As Long
    Dim d2x As Long
    Dim d2y As Long
    Dim m As Long
    Dim n As Long
    Dim i As Integer
    u = x2 - x1
    v = y2 - y1
    d1x = Sgn(u)
    d1y = Sgn(v)
    d2x = Sgn(u)
    d2y = 0
    m = Abs(u)
    n = Abs(v)
    If m < n Then
        d2x = 0
        d2y = Sgn(v)
        m = Abs(v)
        n = Abs(u)
    End If
    s = m / 2
    i = 0
    Do While i <= m
        SetTile x1, y1
        RefreshTile x1, y1
        s = s + n
        If s >= m Then
            s = s - m
            x1 = x1 + d1x
            y1 = y1 + d1y
        Else
            x1 = x1 + d2x
            y1 = y1 + d2y
        End If
        i = i + 1
    Loop
End Sub

Private Sub SetTile(x As Long, y As Long, Optional i As Long = -1, Optional c As Long = -1, Optional DisableStat As Boolean)
    Dim z As Long
    Dim a As Long
    Dim b As Long
    Dim d As Long
    Dim e As Long
    Dim f As Long
    Dim tSE As Long
    Dim tSC As Long
    tSE = xSelectedElement
    tSC = xForcedCol
    If xRBUCount > 0 Then
        'we're using random buffer units
        a = Int(Rnd * xRBUCount)
        tSE = xBufferElements(a)
        tSC = xBufferColors(a)
    End If
    If x < 0 Or x >= xBoardWidth Or y < 0 Or y >= xBoardHeight Then
        Exit Sub
    End If
    If Combo1.ItemData(Combo1.ListIndex) = 1 And i = -1 Then
        World.GetStatLocation CurrentBoard, 0, z, a
        World.MovePlayer CurrentBoard, x, y
        DrawChar (z - 1) - HScroll.Value, (a - 1) - VScroll.Value
    Else
        If i = -1 Then i = tSE
        If c = -1 Then c = tSC
        If chkStats.Value = 1 And DisableStat = False Then
            'create a stat object
            If World.NextFreeStat(CurrentBoard, x + 0, y + 0) >= 0 Then
                World.EraseObjectAt CurrentBoard, x, y
                d = World.BoardID(CurrentBoard, x, y)
                e = World.BoardCol(CurrentBoard, x, y)
                World.SetBoardID CurrentBoard, x, y, i
                b = World.CreateStat(CurrentBoard, x + 0, y + 0, i)
                If b >= 0 Then
                    With MainStatBuffer
                        World.SetObjectInfo1 CurrentBoard, b, .xXStep, .xYStep, .xCycle + 0, .xP1, .xP2, .xP3
                        World.SetObjectOOP CurrentBoard, b, .xOOP
                        Select Case d
                            'everything can be placed on these
                            Case E_Empty, E_Fake, E_Floor, E_WaterN, E_WaterS, E_WaterE, E_WaterW, E_Web
                                World.SetObjectInfo2 CurrentBoard, b, .xFollow, .xLeader, d + 0, e + 0, .xPointer, .xInstruction, .xLength
                                f = 1
                            Case E_Water, E_Lava 'bullets, stars and sharks can be on water
                                If i = E_Shark Or i = E_Bullet Or i = E_Star Then
                                    World.SetObjectInfo2 CurrentBoard, b, .xFollow, .xLeader, d + 0, e + 0, .xPointer, .xInstruction, .xLength
                                    f = 1
                                Else
                                    World.SetObjectInfo2 CurrentBoard, b, .xFollow, .xLeader, 0, 0, .xPointer, .xInstruction, .xLength
                                End If
                            Case Else
                                World.SetObjectInfo2 CurrentBoard, b, .xFollow, .xLeader, 0, 0, .xPointer, .xInstruction, .xLength
                        End Select
                    End With
                End If
                If ((c And &H70) = 0) And f = 1 And d <> E_Empty Then
                    'the object was successfully placed on top of a fake, etc
                    'so if we are using an object with a background color of
                    'zero, we borrow what's under it
                    c = c Or (World.BoardCol(CurrentBoard, x, y) And &H70)
                End If
                World.SetBoardCol CurrentBoard, x, y, c
                
                'some items are configured at the time of first placement
                'so let's do that here
                If DefaultStats(i) = True And bSkipStats = False Then
                    Select Case i
                        Case E_Segment, E_Monitor, E_Player, E_Messenger, E_Bomb, E_Clockwise, E_Counter, E_Stone
                            MainStatBuffer.xIsFilled = True
                        Case Else
                            If MainStatBuffer.xIsFilled = False Then
                                mnuFunctionsEditStats_Click
                                bSetStats = True
                            End If
                    End Select
                End If
                
            End If
        Else
            World.SetBoardID CurrentBoard, x, y, i
            World.SetBoardCol CurrentBoard, x, y, c
        End If
    End If
    DrawChar x - HScroll.Value, y - VScroll.Value
    DrawBlock x - HScroll.Value, y - VScroll.Value
    DrawMethod
End Sub

Private Sub picMain_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    'Dim x As Long
    For x = 1 To Data.Files.Count
        LoadFile Data.Files(1), (x > 1)
    Next x
End Sub

Public Sub LoadFile(fname As String, Optional MakeNewBoards As Boolean)
    Dim ext As String
    If Dir(fname) = "" Then
        Exit Sub
    End If
    If InStr(fname, ".") > 0 Then
        ext = UCase$(Mid$(fname, InStrRev(fname, ".") + 1))
    End If
    Select Case ext
        Case "ZZL", "ZZM"
            MsgBox "Code library '" + ObjectLibrary.OpenLibrary(fname) + "' has been loaded."
            Exit Sub
        Case "SAV", "ZZT", "SZT"
            If World.UnsavedChanges Then
                If MsgBox("Load " + fname + "?" + vbCrLf + "All unsaved changes will be lost.", vbYesNo) = vbNo Then
                    Exit Sub
                End If
            End If
            World.LoadWorld fname
            ConstructMenus
            SetupScrollers
            RefreshBoard
        Case "BRD"
            If MakeNewBoards Then
                World.LoadBoard World.CreateNewBoard, fname
            Else
                World.LoadBoard CurrentBoard, fname
            End If
            ConstructMenus
            RefreshBoard
        Case Else
            MsgBox "Unrecognized format:" + vbCrLf + fname, vbCritical, "Cannot load this file."
            Exit Sub
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    bRunning = False
End Sub

Private Sub mnuBoards_Click(Index As Integer)
    If Index = World.BoardCount + 3 Then
        SetCurrentBoard World.CreateNewBoard
        ConstructMenus
    ElseIf Index = World.BoardCount + 4 Then
        If CurrentBoard > 0 Then
            World.DeleteBoard CurrentBoard
            CurrentBoard = CurrentBoard - 1
            ConstructMenus
        Else
            MsgBox "The title screen cannot be deleted.", vbInformation, "Delete Board"
        End If
    Else
        SetCurrentBoard Index - 1
    End If
    RefreshBoard
End Sub

Private Sub mnuCreatures1_Click(Index As Integer)
    Select Case Index
        Case 0: SetElement2 "bear", True
        Case 1: SetElement2 "ruffian", True
        Case 2: SetElement2 "object", True
        Case 3: SetElement2 "slime", True
        Case 4: SetElement2 "shark", True
        Case 5: SetElement2 "spinning gun", True
        Case 6: SetElement2 "pusher", True
        Case 7: SetElement2 "lion", True
        Case 8: SetElement2 "tiger", True
        Case 9: SetElement2 "centipede head", True
        Case 10: SetElement2 "centipede segment", True
        Case 11: SetElement2 "bullet", True
        Case 12: SetElement2 "star", True
        Case 13: SetElement2 "roton", True
        Case 14: SetElement2 "dragon pup", True
        Case 15: SetElement2 "pairer", True
        Case 16: SetElement2 "spider", True
    End Select
End Sub

Private Sub mnuFileLoad_Click()
    Dim NewFN As String
    CD_ShowOpen_Save Me.hWnd, OFN_EXPLORER Or OFN_FILEMUSTEXIST, NewFN, "", "", "Open ZZT World", OFNFilter, "", True
    If NewFN <> "" Then
        LoadFile NewFN
        AddRecentFile NewFN
    End If
    RefreshBoard
    ConstructMenus
End Sub

Private Sub mnuFileNew_Click()
    If MsgBox("Make a new " + "Classic ZZT" + " world? All unsaved changes will be lost.", vbYesNo, "Confirmation") = vbYes Then
        CurrentBoard = 0
        World.NewWorld False
        SetupScrollers
        RefreshBoard
        ConstructMenus
    End If
End Sub

Private Sub mnuFileSave_Click()
    Dim NewFN As String
    If World.IsSuperZZT Then
        NewFN = World.WorldName + ".szt"
        CD_ShowOpen_Save Me.hWnd, OFN_EXPLORER Or OFN_OVERWRITEPROMPT Or OFN_PATHMUSTEXIST, NewFN, "", "", "Save Super ZZT World", OFNFilterS2, "", False
    Else
        NewFN = World.WorldName + ".zzt"
        CD_ShowOpen_Save Me.hWnd, OFN_EXPLORER Or OFN_OVERWRITEPROMPT Or OFN_PATHMUSTEXIST, NewFN, "", "", "Save ZZT World", OFNFilterS1, "", False
    End If
    
    If NewFN <> "" Then
        World.SetStartBoard CurrentBoard
        World.SaveWorld NewFN
        DoMode
    End If
End Sub

Private Sub mnuFileTransferEB_Click()
    Dim NewFN As String
    CD_ShowOpen_Save Me.hWnd, OFN_EXPLORER Or OFN_OVERWRITEPROMPT Or OFN_PATHMUSTEXIST, NewFN, "", "", "Save BRD Board to File", BRDFilter, "", False
    If NewFN <> "" Then
        World.SaveBoard CurrentBoard, NewFN
    End If
End Sub

Private Sub mnuFileTransferES_Click()
    Dim NewFN As String
    NewFN = World.BoardName(CurrentBoard) + ".bmp"
    CD_ShowOpen_Save Me.hWnd, OFN_EXPLORER Or OFN_OVERWRITEPROMPT Or OFN_PATHMUSTEXIST, NewFN, "", "", "Save BMP screenshot", BMPFilter, "", False
    If NewFN <> "" Then
        DoEvents
        RenderDC picExport.hdc
        SavePicture frmEdit.picExport.Image, NewFN
    End If
End Sub

Private Sub mnuFileTransferIB_Click()
    Dim NewFN As String
    CD_ShowOpen_Save Me.hWnd, OFN_EXPLORER Or OFN_FILEMUSTEXIST, NewFN, "", "", "Load BRD Board from File", BRDFilter, "", True
    If NewFN <> "" Then
        LoadFile NewFN, False
    End If
    ConstructMenus
    RefreshBoard
End Sub

Private Sub mnuFileTransferIW_Click()
    MsgBox "Still under construction..."
End Sub

Private Sub mnuFunctionDeleteObjectsRect_Click()
    Dim x1 As Long
    Dim y1 As Long
    For x1 = SelectionX1 To SelectionX2
        For y1 = SelectionY1 To SelectionY2
            If x1 >= 0 And y1 >= 0 And x1 < xBoardWidth And y1 < xBoardHeight Then
                World.EraseObjectAt CurrentBoard, x1, y1
                DrawChar x1, y1
            End If
        Next y1
    Next x1
    DrawMethod
End Sub

Private Sub mnuFunctionsCopy_Click()
    FillStats MousePosX, MousePosY
    MainStatBuffer.xIsFilled = True
    AddToBuffer World.BoardID(CurrentBoard, MousePosX, MousePosY), World.BoardCol(CurrentBoard, MousePosX, MousePosY), World.ObjectCharAt(CurrentBoard, MousePosX, MousePosY, 1)
    RefreshPreview
End Sub

Private Sub mnuFunctionsDeselect_Click()
    CtrlMode = 0
    CtrlPosX = -1
    CtrlPosY = -1
    SelectionX1 = -1
    SelectionX2 = -1
    SelectionY1 = -1
    SelectionY2 = -1
End Sub

Private Sub mnuFunctionsEditStats_Click()
    Dim x As Long
    x = World.ObjectAt(CurrentBoard, MousePosX, MousePosY)
    If x >= 0 Then
        'FillStats MouseX, MouseY
        frmStatInfo.Show
        frmStatInfo.SetObjectNumber x
    End If
    bSetStats = True
End Sub

Private Sub mnuFunctionsFill_Click()
    Dim x1 As Long
    Dim y1 As Long
    bSkipStats = True
    For x1 = SelectionX1 To SelectionX2
        For y1 = SelectionY1 To SelectionY2
            If x1 >= 0 And y1 >= 0 And x1 < xBoardWidth And y1 < xBoardHeight Then
                SetTile x1, y1
            End If
        Next y1
    Next x1
    RefreshBoard
    bSkipStats = False
End Sub

Function GetFloodPatternAt(x As Long, y As Long) As Boolean
    Dim x1 As Long
    Dim y1 As Long
    If lFillSizeX = 0 Or lFillSizeY = 0 Then
        GetFloodPatternAt = True
        Exit Function
    End If
    x1 = (x - MousePosX) Mod lFillSizeX
    If x1 < 0 Then x1 = x1 + lFillSizeX
    y1 = (y - MousePosX) Mod lFillSizeY
    If y1 < 0 Then y1 = y1 + lFillSizeY
    GetFloodPatternAt = FloodFillPattern(x1, y1)
End Function

Private Sub mnuFunctionsFillAll_Click()
    Dim x1 As Long
    Dim y1 As Long
    bSkipStats = True
    For x1 = 0 To xBoardWidth - 1
        For y1 = 0 To xBoardHeight - 1
            SetTile x1, y1
        Next y1
    Next x1
    bSkipStats = False
End Sub

Private Sub mnuFunctionsFlood_Click()
    FloodFill False 'no pattern
End Sub

Sub FloodFill(bUsePattern As Boolean)
    Dim x1 As Long
    Dim y1 As Long
    Dim iCol As Long
    Dim iID As Long
    bSkipStats = True
    iCol = World.BoardCol(CurrentBoard, MousePosX, MousePosY)
    iID = World.BoardID(CurrentBoard, MousePosX, MousePosY)
    Dim z As Boolean
    'determine eligible tiles
    For x1 = 0 To xBoardWidth - 1
        For y1 = 0 To xBoardHeight - 1
            If Not (World.BoardCol(CurrentBoard, x1, y1) = iCol And World.BoardID(CurrentBoard, x1, y1) = iID) Then
                FloodFillGrid(x1, y1) = 250
            Else
                FloodFillGrid(x1, y1) = 0
            End If
        Next y1
    Next x1
    If (Not bUsePattern) Or GetFloodPatternAt(MousePosX, MousePosY) Then
        SetTile MousePosX, MousePosY
    End If
    FloodFillGrid(MousePosX, MousePosY) = 1
    Do
        z = True
        For x1 = 0 To xBoardWidth - 1
            For y1 = 0 To xBoardHeight - 1
                If FloodFillGrid(x1, y1) = 1 Then
                    If x1 > 0 Then
                        'If World.BoardCol(CurrentBoard, x1 - 1, y1) = iCol And World.BoardID(CurrentBoard, x1 - 1, y1) = iID Then
                        If FloodFillGrid(x1 - 1, y1) < 1 Then
                            FloodFillGrid(x1 - 1, y1) = FloodFillGrid(x1 - 1, y1) + 1
                            If (Not bUsePattern) Or (GetFloodPatternAt(x1 - 1, y1)) Then
                                SetTile x1 - 1, y1
                            End If
                            z = False
                        End If
                    End If
                    If x1 < (xBoardWidth - 1) Then
                        If FloodFillGrid(x1 + 1, y1) < 1 Then
                            FloodFillGrid(x1 + 1, y1) = FloodFillGrid(x1 + 1, y1) + 1
                            If (Not bUsePattern) Or (GetFloodPatternAt(x1 + 1, y1)) Then
                                SetTile x1 + 1, y1
                            End If
                            z = False
                        End If
                    End If
                    If y1 > 0 Then
                        If FloodFillGrid(x1, y1 - 1) < 1 Then
                            FloodFillGrid(x1, y1 - 1) = FloodFillGrid(x1, y1 - 1) + 1
                            If (Not bUsePattern) Or (GetFloodPatternAt(x1, y1 - 1)) Then
                                SetTile x1, y1 - 1
                            End If
                            z = False
                        End If
                    End If
                    If y1 < (xBoardHeight - 1) Then
                        If FloodFillGrid(x1, y1 + 1) < 1 Then
                            FloodFillGrid(x1, y1 + 1) = FloodFillGrid(x1, y1 + 1) + 1
                            If (Not bUsePattern) Or (GetFloodPatternAt(x1, y1 + 1)) Then
                                SetTile x1, y1 + 1
                            End If
                            z = False
                        End If
                    End If
                    FloodFillGrid(x1, y1) = 2 'done with this block
                End If
            Next y1
        Next x1
    Loop While z = False
    RefreshBoard
    bSkipStats = False
End Sub

Private Sub mnuFunctionsOutline_Click()
    Dim x1 As Long
    Dim y1 As Long
    For x1 = SelectionX1 To SelectionX2
        If x1 >= 0 And y1 >= 0 And x1 < xBoardWidth And y1 < xBoardHeight Then
            SetTile x1, SelectionY1
            SetTile x1, SelectionY2
        End If
    Next x1
    For y1 = SelectionY1 To SelectionY2
        If x1 >= 0 And y1 >= 0 And x1 < xBoardWidth And y1 < xBoardHeight Then
            SetTile SelectionX1, y1
            SetTile SelectionX2, y1
        End If
    Next y1
    DrawMethod
End Sub

Private Sub mnuItems1_Click(Index As Integer)
    Select Case Index
        Case 0: SetElement2 "Player", True
        Case 1: SetElement2 "Player" + " clone", True
        Case 2: SetElement2 "ammo"
        Case 3: SetElement2 "torch"
        Case 4: SetElement2 "gem"
        Case 5: SetElement2 "key"
        Case 6: SetElement2 "door"
        Case 7: SetElement2 "scroll", True
        Case 8: SetElement2 "passage", True
        Case 9: SetElement2 "duplicator", True
        Case 10: SetElement2 "bomb", True
        Case 11: SetElement2 "energizer"
        Case 12: SetElement2 "conveyor: clockwise", True
        Case 13: SetElement2 "conveyor: counter cw", True
    End Select
End Sub

Private Sub mnuTerrain1_Click(Index As Integer)
    Select Case Index
        Case 0: SetElement2 "water"
        Case 1: SetElement2 "forest"
        Case 2: SetElement2 "solid"
        Case 3: SetElement2 "normal"
        Case 4: SetElement2 "breakable wall"
        Case 5: SetElement2 "boulder"
        Case 6: SetElement2 "slider: NS"
        Case 7: SetElement2 "slider: EW"
        Case 8: SetElement2 "fake wall"
        Case 9: SetElement2 "invisible wall"
        Case 10: SetElement2 "blink wall", True
        Case 11: SetElement2 "transporter", True
        Case 12: SetElement2 "ricochet"
        Case 13: SetElement2 "board edge"
        Case 14: SetElement2 "monitor", True
        Case 15: SetElement2 "blink ray: horizontal"
        Case 16: SetElement2 "blink ray: vertical"
        Case 17: SetElement2 "player"
        Case 18: SetElement2 "empty"
        Case 19: SetElement2 "floor"
        Case 20: SetElement2 "water n"
        Case 21: SetElement2 "water s"
        Case 22: SetElement2 "water w"
        Case 23: SetElement2 "water e"
    End Select
End Sub

Private Sub SetElement2(eName As String, Optional eStat As Boolean = False)
    Dim x As Long
    SetElementByName eName
    chkStats.Value = Abs(eStat)
    MainStatBuffer.xIsFilled = False
End Sub

Public Sub SetMode(Optional ByVal xmodeName As String = "")
    Dim mnuItemInfo As MENUITEMINFO, hMenu As Long
    Dim BuffStr As String * 80
    If xmodeName <> "" Then
        ModeName = xmodeName
    End If
    If TabDrawMode Then
        ModeName = ModeName + " (drawing)"
    End If
    Me.Caption = App.ProductName + " :: " + ModeName + " :: " + World.WorldName
    mnuBoardTitle.Caption = CStr(CurrentBoard) + ": " + World.BoardName(CurrentBoard)
    
    'shove the 5th menu all the way to the right side
    'for some reason, changing the form's title removes this
    Me.Visible = True
    hMenu = GetMenu(Me.hWnd)
    BuffStr = Space(80)
    With mnuItemInfo
         .cbSize = Len(mnuItemInfo)
         .dwTypeData = BuffStr & Chr(0)
         .fType = MF_STRING
         .cch = Len(mnuItemInfo.dwTypeData)
         .fState = MFS_DEFAULT
         .fMask = MIIM_ID Or MIIM_DATA Or MIIM_TYPE Or MIIM_SUBMENU
    End With
    If GetMenuItemInfo(hMenu, 4, True, mnuItemInfo) = 0 Then
        'menu error
    Else
        mnuItemInfo.fType = mnuItemInfo.fType Or MF_HELP
        If SetMenuItemInfo(hMenu, 4, True, mnuItemInfo) = 0 Then
            'menu error
        End If
    End If
    DrawMenuBar Me.hWnd
End Sub

Private Sub DoMode()
    Select Case ModeType
        Case MODE_EDIT: SetMode "EDIT"
        Case MODE_TEXT: SetMode "TEXT ENTRY"
        Case MODE_OVERVIEW: SetMode "OVERVIEW"
        Case MODE_PASTE: SetMode "PASTE"
        Case MODE_VIEWPORT: SetMode "VIEWPORT"
    End Select
End Sub

Public Sub SimulateKey(KeyCode As Integer, Shift As Integer)
    'Form_KeyDown KeyCode, Shift
    picMain_KeyDown KeyCode, Shift
End Sub

Public Sub RefreshBoard()
    Dim x As Long
    Dim y As Long
    For x = 0 To (xBoardWidth - 1)
        For y = 0 To (xBoardHeight - 1)
            DrawChar x, y
        Next y
    Next x
    DrawMethod
    SetMode
End Sub

Public Sub ConstructMenus()
    Dim i As Integer
    
    ' BOARD menu
    mnuBoards(0).Visible = True
    For i = 1 To mnuBoards.UBound
        Unload mnuBoards(i)
    Next i
    For i = 1 To World.BoardCount + 4
        Load mnuBoards(i)
        If i < World.BoardCount + 2 Then
            mnuBoards(i).Caption = CStr(i - 1) + ": " + World.BoardName(i - 1)
        ElseIf i = World.BoardCount + 2 Then
            mnuBoards(i).Caption = "-"
        ElseIf i = World.BoardCount + 3 Then
            mnuBoards(i).Caption = "Add new Board"
        ElseIf i = World.BoardCount + 4 Then
            mnuBoards(i).Caption = "Delete this Board"
        End If
    Next i
    mnuBoards(0).Visible = False ' This is the divider - make it invisible
    
    ' RECENT FILES menu
    mnuFileRecent(0).Visible = True
    For i = 1 To mnuFileRecent.UBound
        Unload mnuFileRecent(i)
    Next i
    For i = 1 To 8
        If RecentFiles(i - 1) <> "" Then
            Load mnuFileRecent(i)
            If InStr(RecentFiles(i - 1), "\") > 0 Then
                mnuFileRecent(i).Caption = Mid$(RecentFiles(i - 1), InStrRev(RecentFiles(i - 1), "\") + 1)
            Else
                mnuFileRecent(i).Caption = RecentFiles(i - 1)
            End If
            mnuFileRecent(i).Caption = "[&" + CStr(i) + "] " + mnuFileRecent(i).Caption
        End If
    Next i
    If mnuFileRecent.UBound > 0 Then
        mnuFileRecent(0).Visible = False
        mnuFileRecent2.Visible = True
    Else
        mnuFileRecent2.Visible = False
    End If
End Sub

Private Sub picMain_Paint()
    DrawMethod
End Sub

Private Sub picMain_Resize()
    picMain.ScaleWidth = 480
    picMain.ScaleHeight = 400
End Sub

Private Sub DrawPoint()
    If bContinuousDrawing Then
        SetTileLine OldMousePosX, OldMousePosY, MousePosX, MousePosY
    Else
        SetTile MousePosX, MousePosY
        RefreshTile MousePosX, MousePosY
    End If
End Sub

Private Sub tmrCursor_Timer()
    RefreshCursor
    If bDrawPoint Then
        DrawPoint
        bDrawPoint = False
    End If
    OldMouseX = MouseX
    OldMouseY = MouseY
End Sub

Private Sub tmrMain_Timer()
    Dim x As Long
    Dim y As Long
    Dim z As Long
    Dim lHighlightColor As Long
    If bAnimateStuff Then
        AnimCount = (AnimCount + 1) Mod 120
    End If
    AnimCountBlink = (AnimCountBlink + 1) Mod 3
    For z = 0 To World.StatCount(CurrentBoard) 'refresh all onscreen objects
        World.GetStatLocation CurrentBoard, z, x, y
        DrawChar (x - 1) - HScroll.Value, (y - 1) - VScroll.Value
        RefreshTile x - 1, y - 1
        If lblObjCount.Value = 1 Then
            If World.ObjectOOP(CurrentBoard, z) <> "" Then
                lHighlightColor = &H44CC& 'vbYellow
            Else
                lHighlightColor = &HAAFF&
            End If
            picMain.Line ((((x - 1) - HScroll.Value) * 8), (((y - 1) - VScroll.Value) * 16))-((((x - 1) - HScroll.Value) * 8) + 7, (((y - 1) - VScroll.Value) * 16) + 15), lHighlightColor, B
        End If
    Next z
    If AnimCountBlink = 0 Then 'every 3 ticks, process blinks
        SetBlinkStatus xBlink
        xBlink = Not xBlink
        RenderBlinksOnlyDC iUseScale
    End If
    If CtrlMode = 1 Then
        DrawMethod
    End If
    RefreshPreview
    
    ForegroundWindow = GetForegroundWindow
    lblObjCount.Caption = CStr(xMaxObjects - World.StatCount(CurrentBoard))
    picMain.Refresh
End Sub

Sub DrawMethod()
    If iUseScale = 1 Then
        RenderDC
    Else
        RenderDCScale iUseScale
    End If
End Sub

Public Sub FillStats(x As Long, y As Long)
    Dim z As Long
    Dim a As Byte
    Dim b As Byte
    z = World.ObjectAt(CurrentBoard, x, y)
    If z >= 0 Then
        With MainStatBuffer
            World.GetObjectInfo1 CurrentBoard, z, .xXStep, .xYStep, .xCycle, .xP1, .xP2, .xP3
            World.GetObjectInfo2 CurrentBoard, z, .xFollow, .xLeader, b, b, .xPointer, .xInstruction, .xLength
            .xOOP = World.ObjectOOP(CurrentBoard, z)
            .xUseStats = True
        End With
    Else
        MainStatBuffer.xUseStats = False
    End If
End Sub

Public Function SelectedElement() As Long
    SelectedElement = xSelectedElement
End Function

Public Sub AddToBuffer(xID As Long, xCol As Long, xChar As Long)
    Dim x As Long
    For x = 9 To 1 Step -1 'rotate all buffer items back
        xBufferElements(x) = xBufferElements(x - 1)
        xBufferColors(x) = xBufferColors(x - 1)
        xBufferStats(x) = xBufferStats(x - 1)
        xBufferChars(x) = xBufferChars(x - 1)
    Next x
    xBufferElements(0) = xID
    xBufferColors(0) = xCol
    xBufferStats(0) = MainStatBuffer
    xBufferChars(0) = xChar
    picture5_MouseDown 1, 0, 0, 0 'set new item as current
    RedrawBuffer
End Sub

Public Sub RedrawBuffer()
    Dim x As Long
    For x = 0 To 9
        If xBufferElements(x) < E_BlueText Or xBufferElements(x) > E_WhiteText Then
            SetChar1 x, 0, xBufferChars(x), xBufferColors(x) + 0, Picture5.hdc
        Else
            SetChar1 x, 0, xBufferChars(x), DefaultColor(xBufferElements(x) + 0), Picture5.hdc
        End If
    Next x
End Sub

Public Sub RedrawBuffer2()
    Dim x As Long
    'Picture4.Cls
    For x = 0 To 9
        SetChar1 x, 0, DefaultChar(xBufferElements2(x) + 0, 14, 3), SelectedColor, Picture4.hdc
    Next x
End Sub

Public Function SelectedChar() As Long
    SelectedChar = xChar
End Function

Public Function SelectedColor() As Byte
    SelectedColor = xForcedCol
End Function

Public Sub SetElementByName(eName As String)
    Dim x As Long
    Combo1.ListIndex = -1
    Combo1.ListIndex = ElementFromName(eName)
End Sub

Public Function ElementFromName(eName As String) As Integer
    Dim x As Integer
    ElementFromName = -1
    For x = 0 To Combo1.ListCount - 1
        If UCase(eName) = UCase(Mid(Combo1.List(x), 6)) Then
            ElementFromName = x
            Exit For
        End If
    Next x
End Function

Public Function ElementNumFromName(eName As String) As Integer
    Dim x As Integer
    For x = 0 To Combo1.ListCount - 1
        If UCase(eName) = UCase(Mid(Combo1.List(x), 6)) Then
            ElementNumFromName = Val(Mid$(Combo1.List(x), 2, 2))
            Exit For
        End If
    Next x
End Function

Public Sub SetElement(eElem As Long, Optional useDefaultStat As Boolean = False)
    Dim x As Long
    For x = 0 To Combo1.ListCount - 1
        If eElem = Val(Mid(Combo1.List(x), 2, 2)) And Left(Combo1.List(x), 1) = "[" Then
            bAutoSetElement = True
            Combo1.ListIndex = -1
            Combo1.ListIndex = x + 0
            xSelectedElement = eElem
            Exit For
        End If
    Next x
End Sub

Private Sub chkDefaultColor_Click()
    Picture1_MouseDown 0, 32, 0, 0, 0
End Sub

Private Sub chkStats_Click()
    MainStatBuffer.xUseStats = (chkStats.Value = 1)
End Sub

Private Sub Combo1_Click()
    If Not bChangingElement Then
        ChangeElement
    End If
End Sub

Sub ChangeElement()
    bChangingElement = True
    Dim x As Long
    xSelectedElement = 0
    If Combo1.ListIndex > -1 Then
        If Left(Combo1.List(Combo1.ListIndex), 1) = "[" Then
            xSelectedElement = Val(Mid(Combo1.List(Combo1.ListIndex), 2, 2))
        End If
        If Not bAutoSetElement Then
            If DefaultStats(xSelectedElement) Then
                chkStats.Value = 1
            Else
                chkStats.Value = 0
            End If
        End If
        With MainStatBuffer
            .xCycle = DefaultCycle(xSelectedElement)
            .xFollow = 0
            .xInstruction = 0
            .xLeader = 0
            .xLength = 0
            .xOOP = ""
            .xP1 = DefaultP1(xSelectedElement)
            .xP2 = DefaultP2(xSelectedElement)
            .xP3 = DefaultP3(xSelectedElement)
            .xPointer = 0
            .xUseStats = (chkStats.Value = 1)
            .xXStep = DefaultXStep(xSelectedElement)
            .xYStep = DefaultYStep(xSelectedElement)
        End With
        bAutoSetElement = False
        'If Not bIsSetByProgram Then
        bSetStats = True ' False
        'End If
    End If
    bChangingElement = False
End Sub
Public Sub SetToolColor(i As Long, c As Long)
    Picture1(i).BackColor = c
End Sub

Public Sub ShowToolColors()
    Picture1(xSelectedBCol).Line (0.1, 0.1)-(0.8, 0.8), , B
    Picture1(xSelectedFCol).Circle (0.45, 0.45), 0.25
    lblColor.Caption = "col: 0x" + Hex(xSelectedBCol) + Hex(xSelectedFCol) + " / " + CStr((xSelectedBCol * 16) + xSelectedFCol)
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
    frmEdit.SimulateKey KeyCode, Shift
End Sub

'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    frmEdit.SimulateKey KeyCode, Shift
'End Sub

Public Sub NextColorF()
    xSelectedFCol = (xSelectedFCol + 1) Mod 16
    ClearColors
    ShowToolColors
End Sub

Public Sub NextColorB()
    xSelectedBCol = (xSelectedBCol + 1) Mod 16
    ClearColors
    ShowToolColors
End Sub

Private Sub ClearColors()
    Dim a As Long
    For a = 0 To 15
        Picture1(a).Cls
    Next a
    ShowToolColors
End Sub

Private Sub mnuAbout_Click()
    ShowAboutWindow
End Sub

Private Sub mnuOptimizeEmpties_Click()
    World.ChangeAllTiles E_Empty, 256, 256, 0
End Sub

Private Sub mnuOptionsGraphics2X_Click()
    If mnuOptionsGraphics2X.Checked = False Then
        SetEditScale 2
        mnuOptionsGraphics2X.Checked = True
    Else
        SetEditScale 1
        mnuOptionsGraphics2X.Checked = False
    End If
    RefreshBoard
End Sub

Private Sub mnuOptionsGrid_Click()
    If mnuOptionsGrid.Checked = False Then
        mnuOptionsGrid.Checked = True
    Else
        mnuOptionsGrid.Checked = False
    End If
End Sub

Private Sub Picture1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    frmEdit.SimulateKey KeyCode, Shift
End Sub

Private Sub Picture1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim a As Long
    If (Button And 1) = 1 Then
        xSelectedFCol = Index
    ElseIf (Button And 2) = 2 Then
        xSelectedBCol = Index
    End If
    ClearColors
    'ShowToolColors
End Sub

Public Sub RefreshPreview()
    Dim DCol As Long
    xChar = DefaultChar(xSelectedElement, xSelectedFCol + (xSelectedBCol * 16))
    If chkDefaultColor.Value = 1 Then 'Or (xSelectedElement >= E_BlueText And xSelectedElement <= E_WhiteText) Then
        xForcedCol = DefaultColor(xSelectedElement, xSelectedFCol + (xSelectedBCol * 16))
        DCol = xForcedCol
    Else
        xForcedCol = xSelectedFCol + (xSelectedBCol * 16)
        If xSelectedElement >= E_BlueText And xSelectedElement <= E_WhiteText Then
            DCol = DefaultColor(xSelectedElement, 0, 0)
        Else
            DCol = xForcedCol
        End If
    End If
    If chkBlinks.Value = 0 Then
        xForcedCol = (xForcedCol And &H7F)
    End If
    SetChar1 0, 0, xChar, DCol + 0, Picture3.hdc, 2
    SetChar1 3, 0, xChar, DCol + 0, Picture3.hdc
    SetChar1 3, 1, xChar, DCol + 0, Picture3.hdc
    SetChar1 4, 0, xChar, DCol + 0, Picture3.hdc
    SetChar1 4, 1, xChar, DCol + 0, Picture3.hdc
    RedrawBuffer2
End Sub

Private Sub picture5_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    x = Int(x)
    If Button = 1 Then
        SetElement xBufferElements(x) + 0
        xSelectedFCol = xBufferColors(x) And 15
        xSelectedBCol = (xBufferColors(x) And 240) \ 16
        MainStatBuffer = xBufferStats(x)
        If MainStatBuffer.xUseStats = True Then
            chkStats.Value = 1
        Else
            chkStats.Value = 0
        End If
        RefreshPreview
        ClearColors
    ElseIf Button = 2 Then
        AddToBuffer xSelectedElement, xForcedCol, xChar
    End If
End Sub

Private Sub Picture4_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    x = Int(x)
    chkStats.Value = 0
    SetElement xBufferElements2(x) + 0
    RefreshPreview
End Sub

Private Sub picture5_Paint()
    RedrawBuffer
End Sub

Private Sub Picture4_Paint()
    RedrawBuffer2
End Sub

Private Sub VScroll_Change()
    RefreshBoard
End Sub

Private Sub SetupScrollers()
    If World.IsSuperZZT Then
        HScroll.Enabled = True
        VScroll.Enabled = True
        HScroll.Value = 18
        VScroll.Value = 28
        xBoardWidth = 96
        xBoardHeight = 80
    Else
        HScroll.Enabled = False
        VScroll.Enabled = False
        HScroll.Value = 0
        VScroll.Value = 0
        xBoardHeight = 25
        xBoardWidth = 60
    End If
End Sub

Private Sub VScroll_Scroll()
    VScroll_Change
End Sub

Private Sub GetSelection(bCut As Boolean)
    Dim x As Long
    Dim y As Long
    Dim z As Long
    Dim a As Long
    Dim b As Long
    If Not (SelectionX1 > -1 And SelectionX2 > -1 And SelectionY1 > -1 And SelectionY2 > -1) Then
        'MsgBox "Make a selection first..."
        Exit Sub
    End If
    xCopySizeX = (SelectionX2 - SelectionX1) + 1
    xCopySizeY = (SelectionY2 - SelectionY1) + 1
    z = ((SelectionX2 - SelectionX1) + 1) * ((SelectionY2 - SelectionY1) + 1)
    ReDim xCopyBuffer(1 To z) As xtiSelectionType
    ReDim xCopyDispBuffer(1 To z) As xtiANSIPair
    ReDim xCopyObjects(0) As xtiBufferStats
    z = 1 'bufferposition
    a = 0 'objectcount
    xCopyObjectCount = 0
    For y = SelectionY1 To SelectionY2
        For x = SelectionX1 To SelectionX2
            With xCopyBuffer(z)
                .xColor = World.BoardCol(CurrentBoard, x, y)
                .xID = World.BoardID(CurrentBoard, x, y)
                b = World.ObjectAt(CurrentBoard, x, y)
                If b > -1 Then
                    a = a + 1
                    ReDim Preserve xCopyObjects(0 To a) As xtiBufferStats
                    With xCopyObjects(a)
                        World.GetObjectInfo1 CurrentBoard, b, .xXStep, .xYStep, .xCycle, .xP1, .xP2, .xP3
                        World.GetObjectInfo2 CurrentBoard, b, .xFollow, .xLeader, .xUnderID, .xUnderColor, .xPointer, .xInstruction, .xLength
                        .xOOP = World.ObjectOOP(CurrentBoard, b)
                        .xIsFilled = True
                        .xID = xCopyBuffer(z).xID
                    End With
                    .xObjectRef = a
                    xCopyObjectCount = xCopyObjectCount + 1
                Else
                    .xObjectRef = 255
                End If
            End With
            With xCopyDispBuffer(z)
                .xChar = DefaultChar(xCopyBuffer(z).xID + 0, xCopyBuffer(z).xColor + 0)
                If xCopyBuffer(z).xID >= E_BlueText And xCopyBuffer(z).xID <= E_WhiteText Then
                    'properly display text
                    .xChar = World.BoardCol(CurrentBoard, x, y)
                    .xCol = DefaultColor(xCopyBuffer(z).xID + 0, xCopyBuffer(z).xColor + 0)
                ElseIf xCopyBuffer(z).xID = 0 Then
                    .xCol = 0 'empties are invisible
                Else
                    .xCol = World.BoardCol(CurrentBoard, x, y)
                End If
                
            End With
            z = z + 1
            If bCut = True Then
                SetTile x, y, E_Empty, 0
            End If
        Next x
    Next y
End Sub

Sub PutSelection()
    If xCopySizeX > 0 And xCopySizeY > 0 Then
        ModeType = MODE_PASTE
        DoMode
    End If
End Sub

Sub PasteSelection(x As Long, y As Long)
    Dim a As Long
    Dim b As Long
    Dim z As Long
    Dim c As Long
    z = 1
    For b = y To (y + xCopySizeY) - 1
        For a = x To (x + xCopySizeX) - 1
            With xCopyBuffer(z)
                If a < xBoardWidth And b < xBoardHeight Then
                    c = World.ObjectAt(CurrentBoard, a, b)
                    If c <> 0 Then ' don't erase player
                        World.EraseObjectAt CurrentBoard, a, b
                        If .xObjectRef < 255 Then
                            If World.NextFreeStat(CurrentBoard, a + 0, b + 0) > -1 Then
                                World.SetBoardID CurrentBoard, a, b, .xID + 0
                                World.SetBoardCol CurrentBoard, a, b, .xColor + 0
                                With xCopyObjects(.xObjectRef)
                                    c = World.CreateStat(CurrentBoard, a + 0, b + 0, xCopyBuffer(z).xID + 0, c, 0, 0, 0, "")
                                    If c >= 0 Then
                                        World.SetObjectInfo1 CurrentBoard, c, .xXStep, .xYStep, .xCycle, .xP1, .xP2, .xP3
                                        World.SetObjectInfo2 CurrentBoard, c, .xFollow, .xLeader, .xUnderID, .xUnderColor, .xPointer, .xInstruction, 0
                                        World.SetObjectOOP CurrentBoard, c, .xOOP
                                    End If
                                End With
                            End If
                        Else
                            World.SetBoardID CurrentBoard, a, b, .xID + 0
                            World.SetBoardCol CurrentBoard, a, b, .xColor + 0
                        End If
                    End If
                End If
            End With
            z = z + 1
        Next a
    Next b
End Sub

Sub FlipSelectionV()
    Dim x As Long
    Dim y As Long
    Dim z1 As Long
    Dim z2 As Long
    Dim t1 As xtiANSIPair
    Dim t2 As xtiSelectionType
    For y = 0 To (xCopySizeY \ 2) - 1
        For x = 0 To xCopySizeX - 1
            z1 = 1 + (x + (y * xCopySizeX))
            z2 = 1 + (x + (((xCopySizeY - y) - 1) * xCopySizeX))
            'swap 1 (element)
            t2 = xCopyBuffer(z1)
            xCopyBuffer(z1) = xCopyBuffer(z2)
            xCopyBuffer(z2) = t2
            'swap 2 (display)
            t1 = xCopyDispBuffer(z1)
            xCopyDispBuffer(z1) = xCopyDispBuffer(z2)
            xCopyDispBuffer(z2) = t1
        Next x
    Next y
    RefreshPaste
End Sub

Sub FlipSelectionH()
    Dim x As Long
    Dim y As Long
    Dim z1 As Long
    Dim z2 As Long
    Dim t1 As xtiANSIPair
    Dim t2 As xtiSelectionType
    For y = 0 To xCopySizeY - 1
        For x = 0 To (xCopySizeX \ 2) - 1
            z1 = 1 + (x + (y * xCopySizeX))
            z2 = 1 + (((xCopySizeX - x) - 1) + (y * xCopySizeX))
            'swap 1 (element)
            t2 = xCopyBuffer(z1)
            xCopyBuffer(z1) = xCopyBuffer(z2)
            xCopyBuffer(z2) = t2
            'swap 2 (display)
            t1 = xCopyDispBuffer(z1)
            xCopyDispBuffer(z1) = xCopyDispBuffer(z2)
            xCopyDispBuffer(z2) = t1
        Next x
    Next y
    RefreshPaste
End Sub

Private Sub vsRBU_Change()
    lblRBU = vsRBU
    xRBUCount = vsRBU
End Sub
