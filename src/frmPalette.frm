VERSION 5.00
Begin VB.Form frmPalette 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tools"
   ClientHeight    =   1440
   ClientLeft      =   150
   ClientTop       =   390
   ClientWidth     =   7665
   ControlBox      =   0   'False
   Icon            =   "frmPalette.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   7665
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
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
      Left            =   7080
      TabIndex        =   0
      Top             =   960
      Width           =   2055
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsGraphics 
         Caption         =   "Graphics"
         Begin VB.Menu mnuOptionsGraphics2X 
            Caption         =   "Use 2X mode (slower)"
         End
         Begin VB.Menu mnuOptionsGrid 
            Caption         =   "Show Grid"
         End
      End
      Begin VB.Menu mnuOptions1 
         Caption         =   "-"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuOptimize 
      Caption         =   "O&ptimize"
      Begin VB.Menu mnuOptimizeEmpties 
         Caption         =   "Change ALL empties to color 00 (world-wide)"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
      NegotiatePosition=   3  'Right
   End
End
Attribute VB_Name = "frmPalette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


