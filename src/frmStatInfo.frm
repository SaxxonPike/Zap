VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmStatInfo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Stat Element Information"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6135
   ControlBox      =   0   'False
   Icon            =   "frmStatInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frameTab 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   5775
      Index           =   1
      Left            =   240
      TabIndex        =   49
      Top             =   480
      Width           =   5655
      Begin VB.CommandButton cmdCommand5 
         Caption         =   "Stop Music"
         Height          =   255
         Left            =   4320
         TabIndex        =   55
         ToolTipText     =   "Stop music playback."
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton cmdCommand4 
         Caption         =   "Play Music"
         Height          =   255
         Left            =   4320
         TabIndex        =   56
         ToolTipText     =   "Play all music in this code. This will not process loops or jumps."
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdCommand3 
         Caption         =   "Insert"
         Height          =   255
         Left            =   3120
         TabIndex        =   54
         ToolTipText     =   "Insert code from a library into this object's code."
         Top             =   480
         Width           =   975
      End
      Begin VB.ComboBox cmbCode1 
         Height          =   315
         ItemData        =   "frmStatInfo.frx":038A
         Left            =   0
         List            =   "frmStatInfo.frx":038C
         Style           =   2  'Dropdown List
         TabIndex        =   53
         ToolTipText     =   "Select your library here."
         Top             =   120
         Width           =   2535
      End
      Begin VB.CommandButton cmdCommand2 
         Caption         =   "Load Library..."
         Height          =   255
         Left            =   2640
         TabIndex        =   52
         ToolTipText     =   "Load an object library from disk."
         Top             =   120
         Width           =   1455
      End
      Begin VB.ComboBox cmbCode2 
         Height          =   315
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   51
         ToolTipText     =   "Within the library, you can select code fragments here."
         Top             =   480
         Width           =   3015
      End
      Begin VB.TextBox txtCode1 
         BackColor       =   &H00800000&
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
         Height          =   4575
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   50
         Top             =   840
         Width           =   5415
      End
      Begin VB.Line Line1 
         X1              =   4200
         X2              =   4200
         Y1              =   720
         Y2              =   120
      End
      Begin VB.Label lblCode 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   0
         TabIndex        =   57
         Top             =   5520
         Width           =   5655
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1320
      TabIndex        =   48
      Top             =   6480
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   3720
      TabIndex        =   47
      Top             =   6480
      Width           =   2295
   End
   Begin VB.Frame frameTab 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   5775
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   5655
      Begin VB.CommandButton Command3 
         Caption         =   "Unlock Sliders"
         Height          =   375
         Left            =   840
         TabIndex        =   46
         Top             =   3960
         Width           =   1695
      End
      Begin VB.Frame Frame1 
         Caption         =   "Parameter 1:"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   0
         Left            =   0
         TabIndex        =   41
         Top             =   960
         Width           =   3375
         Begin VB.HScrollBar HScroll1 
            Height          =   255
            Index           =   0
            Left            =   120
            Max             =   255
            TabIndex        =   44
            Top             =   240
            Width           =   3135
         End
         Begin VB.PictureBox picChar 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   180
            ScaleHeight     =   16
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   200
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   540
            Visible         =   0   'False
            Width           =   3000
         End
         Begin VB.CheckBox chkStat 
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   42
            Top             =   360
            Visible         =   0   'False
            Width           =   3135
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "0"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   45
            Top             =   555
            Width           =   3135
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Parameter 2:"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   1
         Left            =   0
         TabIndex        =   37
         Top             =   1920
         Width           =   3375
         Begin VB.HScrollBar HScroll1 
            Height          =   255
            Index           =   1
            Left            =   120
            Max             =   255
            TabIndex        =   39
            Top             =   240
            Width           =   3135
         End
         Begin VB.CheckBox chkStat 
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   38
            Top             =   360
            Visible         =   0   'False
            Width           =   3135
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "0"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   40
            Top             =   555
            Width           =   3135
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Parameter 3:"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   2
         Left            =   0
         TabIndex        =   32
         Top             =   2880
         Width           =   3375
         Begin VB.HScrollBar HScroll1 
            Height          =   255
            Index           =   2
            Left            =   120
            Max             =   255
            TabIndex        =   35
            Top             =   240
            Width           =   3135
         End
         Begin VB.ComboBox cmbBoard 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   240
            Visible         =   0   'False
            Width           =   3135
         End
         Begin VB.CheckBox chkStat 
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   33
            Top             =   360
            Visible         =   0   'False
            Width           =   3135
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "0"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   36
            Top             =   555
            Width           =   3135
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Other Info"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   3480
         TabIndex        =   2
         Top             =   960
         Width           =   2175
         Begin VB.CommandButton Command2 
            Caption         =   "v"
            Height          =   255
            Index           =   2
            Left            =   1440
            TabIndex        =   5
            TabStop         =   0   'False
            ToolTipText     =   "Step-Down. For objects, this is the walk direction. For transporters, this is the teleport direction."
            Top             =   1620
            Width           =   495
         End
         Begin VB.CommandButton Command2 
            Caption         =   ">"
            Height          =   375
            Index           =   1
            Left            =   1800
            TabIndex        =   7
            TabStop         =   0   'False
            ToolTipText     =   "Step-Right. For objects, this is the walk direction. For transporters, this is the teleport direction."
            Top             =   1260
            Width           =   255
         End
         Begin VB.CommandButton Command2 
            Caption         =   "X"
            Height          =   375
            Index           =   4
            Left            =   1560
            TabIndex        =   8
            TabStop         =   0   'False
            ToolTipText     =   "Step-None. This resets the step-positions."
            Top             =   1260
            Width           =   255
         End
         Begin VB.CommandButton Command2 
            Caption         =   "<"
            Height          =   375
            Index           =   0
            Left            =   1320
            TabIndex        =   9
            TabStop         =   0   'False
            ToolTipText     =   "Step-Left. For objects, this is the walk direction. For transporters, this is the teleport direction."
            Top             =   1260
            Width           =   255
         End
         Begin VB.CommandButton Command2 
            Caption         =   "^"
            Height          =   255
            Index           =   3
            Left            =   1440
            TabIndex        =   19
            TabStop         =   0   'False
            ToolTipText     =   "Step-Up. For objects, this is the walk direction. For transporters, this is the teleport direction."
            Top             =   1020
            Width           =   495
         End
         Begin VB.TextBox txtInfo 
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
            TabIndex        =   18
            Text            =   "1"
            ToolTipText     =   "Cycle. This determines how quickly an object is able to execute movement."
            Top             =   960
            Width           =   495
         End
         Begin VB.TextBox txtInfo 
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
            TabIndex        =   17
            Text            =   "0"
            ToolTipText     =   "X-Step. This has many uses. You can set a specific step-direction using the arrows on the right side."
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txtInfo 
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
            Left            =   1560
            TabIndex        =   16
            Text            =   "0"
            ToolTipText     =   "Y-Step. This has many uses. You can set a specific step-direction using the arrows on the right side."
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txtInfo 
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
            Left            =   720
            TabIndex        =   15
            Text            =   "0"
            ToolTipText     =   "Follow. When unsure, it's best to leave this at zero."
            Top             =   1320
            Width           =   495
         End
         Begin VB.TextBox txtInfo 
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
            Left            =   720
            TabIndex        =   14
            Text            =   "0"
            ToolTipText     =   "Leader. When unsure, it's best to leave this at zero."
            Top             =   1680
            Width           =   495
         End
         Begin VB.TextBox txtInfo 
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
            Left            =   720
            TabIndex        =   13
            ToolTipText     =   "Under ID. This determines what is under the stat object. Leave it blank to leave it unchanged."
            Top             =   2040
            Width           =   495
         End
         Begin VB.TextBox txtInfo 
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
            Left            =   1560
            TabIndex        =   12
            ToolTipText     =   "Under Color. This determines what color is under the stat object. Leave it blank to leave it unchanged."
            Top             =   2040
            Width           =   495
         End
         Begin VB.TextBox txtInfo 
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
            Left            =   720
            TabIndex        =   11
            Text            =   "0"
            ToolTipText     =   "Pointer. For internal ZZT use. This should be left at zero."
            Top             =   2400
            Width           =   1335
         End
         Begin VB.TextBox txtInfo 
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
            Left            =   720
            TabIndex        =   10
            Text            =   "0"
            ToolTipText     =   "Current OOP instruction. This determines where an object is currently executing its code."
            Top             =   2760
            Width           =   1335
         End
         Begin VB.ComboBox cmbBinds 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmStatInfo.frx":038E
            Left            =   720
            List            =   "frmStatInfo.frx":0395
            Style           =   2  'Dropdown List
            TabIndex        =   6
            ToolTipText     =   "Binds. This allows you to make an object copy another object's code - and save space while doing so."
            Top             =   3120
            Width           =   1335
         End
         Begin VB.TextBox txtInfo 
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
            Left            =   720
            TabIndex        =   4
            Text            =   "0"
            ToolTipText     =   "X-Step. This has many uses. You can set a specific step-direction using the arrows on the right side."
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtInfo 
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
            Left            =   1560
            TabIndex        =   3
            Text            =   "0"
            ToolTipText     =   "X-Step. This has many uses. You can set a specific step-direction using the arrows on the right side."
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Cycle"
            Height          =   255
            Index           =   0
            Left            =   90
            TabIndex        =   31
            Top             =   990
            Width           =   600
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Step X"
            Height          =   255
            Index           =   1
            Left            =   90
            TabIndex        =   30
            Top             =   630
            Width           =   600
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Follow"
            Height          =   255
            Index           =   3
            Left            =   90
            TabIndex        =   29
            Top             =   1350
            Width           =   600
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Leader"
            Height          =   255
            Index           =   4
            Left            =   90
            TabIndex        =   28
            Top             =   1710
            Width           =   600
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "UnderID"
            Height          =   255
            Index           =   5
            Left            =   90
            TabIndex        =   27
            Top             =   2070
            Width           =   600
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Col"
            Height          =   255
            Index           =   6
            Left            =   1200
            TabIndex        =   26
            Top             =   2070
            Width           =   330
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Pointer"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   25
            Top             =   2445
            Width           =   570
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "oop inst."
            Height          =   255
            Index           =   8
            Left            =   30
            TabIndex        =   24
            Top             =   2790
            Width           =   660
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Binds"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   23
            Top             =   3180
            Width           =   570
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "X"
            Height          =   255
            Index           =   9
            Left            =   90
            TabIndex        =   22
            Top             =   270
            Width           =   600
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Y"
            Height          =   255
            Index           =   10
            Left            =   1200
            TabIndex        =   21
            Top             =   270
            Width           =   330
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "SY"
            Height          =   255
            Index           =   11
            Left            =   1200
            TabIndex        =   20
            Top             =   630
            Width           =   330
         End
      End
   End
   Begin ComctlLib.TabStrip TabStrip 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   11033
      TabWidthStyle   =   2
      ShowTips        =   0   'False
      TabFixedWidth   =   2646
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Stats"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Code"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmStatInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyStatBuffer As xtiBufferStats
Dim StatNumber As Long
Dim StatColor As Long
Dim ObjectNum As Long
Dim uC As Byte
Dim uI As Byte
Dim bEdited As Boolean
Dim bCancelEdit As Boolean
Const ZZMZZLfilter = "All Openable Files|*.ZZM;*.ZZL;*.TXT|ZZM Music Library (*.ZZM)|*.ZZM|ZZL Object Library (*.ZZL)|*.ZZL|Text File (*.TXT)|*.TXT|All Files|*.*"

Public Sub SetObjectNumber(z As Long)
    Dim a As Long
    Dim x As Long
    Dim y As Long
    Dim d(0 To 2) As String
    Dim e As String
    Dim f As String
    Dim b As Long
    Dim c As Byte
    StatNumber = z
    ObjectNum = z
    If z = -1 Then
        MyStatBuffer = MainStatBuffer
    Else
        With MyStatBuffer
            World.GetObjectInfo1 CurrentBoard, z, .xXStep, .xYStep, .xCycle, .xP1, .xP2, .xP3
            World.GetObjectInfo2 CurrentBoard, z, .xFollow, .xLeader, uC, uI, .xPointer, .xInstruction, .xLength
            .xOOP = World.ObjectOOP(CurrentBoard, z)
        End With
        World.GetStatLocation CurrentBoard, z, x, y
        x = x - 1
        y = y - 1
        b = World.BoardID(CurrentBoard, x, y)
        StatColor = (World.BoardCol(CurrentBoard, x, y) And 127)
        If World.ObjectName(CurrentBoard, z) <> "" Then
            Me.Caption = """" + World.ObjectName(CurrentBoard, z) + """"
        End If
        Me.Caption = Me.Caption + " #" + CStr(z) + " (" + CStr(x) + "," + CStr(y) + ")"
        txtCode1.Left = ((frameTab(1).Width - txtCode1.Width) / 2)
        If World.BoardID(CurrentBoard, x, y) = E_Scroll Then
            lblCode = "Scroll"
        Else
            lblCode = "Object"
        End If
        lblCode = lblCode + " #" + CStr(z) + " at (" + CStr(x) + "," + CStr(y) + ")"
        txtCode1 = World.ObjectOOP(CurrentBoard, ObjectNum)
    End If
    d(0) = DefaultP1Name(b)
    d(1) = DefaultP2Name(b)
    d(2) = DefaultP3Name(b)
    For a = 0 To 2
        If d(a) <> "" Then
            e = Left$(d(a), InStr(d(a), ",") - 1)
            If InStr(e, "-") > 0 Then
                f = Mid$(e, InStr(e, "-") + 1)
                e = Left$(e, InStr(e, "-") - 1)
                HScroll1(a).Min = Val(e)
                HScroll1(a).Max = Val(f)
                If e = 0 And f = 1 Then 'checkbox
                    HScroll1(a).Visible = False
                    chkStat(a).Visible = True
                    chkStat(a).Caption = Mid$(d(a), InStr(d(a), "(") + 1)
                    chkStat(a).Caption = Left$(chkStat(a).Caption, Len(chkStat(a).Caption) - 1)
                    d(a) = Left$(d(a), InStr(d(a), "(") - 1)
                    Label1(a).Visible = False
                End If
            Else
                HScroll1(a).Enabled = False
            End If
            
            d(a) = Mid$(d(a), InStr(d(a), ",") + 1)
            Frame1(a).Caption = Frame1(a).Caption + " " + d(a)
            Frame1(a).Enabled = True
        End If
    Next a
    Select Case b
        Case E_Object
            picChar.Visible = True
            PopulateObjects
        Case E_Passage
            HScroll1(2).Visible = False
            Frame1(2).Enabled = True
            cmbBoard.Visible = True
            For a = 0 To World.BoardCount
                cmbBoard.AddItem CStr(a) + ": " + World.BoardName(a)
            Next a
            If MyStatBuffer.xP3 < cmbBoard.ListCount Then
                cmbBoard.ListIndex = MyStatBuffer.xP3
            Else
                cmbBoard.ListIndex = 0
            End If
    End Select
    With MyStatBuffer
        txtInfo(0) = .xCycle
        txtInfo(1) = .xXStep
        txtInfo(2) = .xYStep
        txtInfo(3) = .xFollow
        txtInfo(4) = .xLeader
        'txtInfo(5) = .xUnderID
        'txtInfo(6) = .xUnderColor
        txtInfo(7) = .xPointer
        txtInfo(8) = .xInstruction
        If .xP1 > HScroll1(0).Max Or .xP2 > HScroll1(1).Max Or .xP3 > HScroll1(2).Max Then
            'auto-unlock sliders for large values
            Command3_Click
        End If
        HScroll1(0).Value = .xP1
        HScroll1(1).Value = .xP2
        HScroll1(2).Value = .xP3
    End With
End Sub

Private Sub chkStat_Click(Index As Integer)
    HScroll1(Index).Value = chkStat(Index).Value
    bEdited = True
End Sub



Private Sub cmbBinds_Change()
    bEdited = True
End Sub

Private Sub cmbBoard_Change()
    bEdited = True
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click(Index As Integer)
    Select Case Index
        Case 0
            txtInfo(1) = -1
            txtInfo(2) = 0
        Case 1
            txtInfo(1) = 1
            txtInfo(2) = 0
        Case 2
            txtInfo(1) = 0
            txtInfo(2) = 1
        Case 3
            txtInfo(1) = 0
            txtInfo(2) = -1
        Case 4
            txtInfo(1) = 0
            txtInfo(2) = 0
    End Select
End Sub

Private Sub Command3_Click()
    Frame1(0).Enabled = True
    Frame1(1).Enabled = True
    Frame1(2).Enabled = True
    HScroll1(0).Enabled = True
    HScroll1(1).Enabled = True
    HScroll1(2).Enabled = True
    HScroll1(0).Max = 255
    HScroll1(1).Max = 255
    HScroll1(2).Max = 255
    HScroll1(0).Min = 0
    HScroll1(1).Min = 0
    HScroll1(2).Min = 0
    If chkStat(0).Visible = True Then HScroll1(0).Visible = True: Label1(0).Visible = True
    If chkStat(1).Visible = True Then HScroll1(1).Visible = True: Label1(1).Visible = True
    If chkStat(2).Visible = True Then HScroll1(2).Visible = True: Label1(2).Visible = True
    chkStat(0).Visible = False
    chkStat(1).Visible = False
    chkStat(2).Visible = False
End Sub

Private Sub Command4_Click()
    frmCode.Visible = True
    frmCode.SetObjectNumber StatNumber
    Unload Me
End Sub

Private Sub Command5_Click()
    Sound.StopPlaying
    If bEdited Then
        If MsgBox("Discard changes?", vbYesNo, "Confirmation") = vbYes Then
            bCancelEdit = True
            Unload Me
        End If
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
    'frmEdit.Enabled = False
    'frmPalette.Enabled = False
    DisableForms
    cmbBinds.ListIndex = 0
    If World.IsSuperZZT Then
        'because superZZT uses 30-wide
        txtCode1.Width = 3975
        'txtCode1.Left = 720
    Else
        txtCode1.Width = 5415
        'txtCode1.Left = 0
    End If
    TabStrip.TabIndex = 0
    RefreshTabs
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim x As Long
    Dim y As Long
    Dim z As Long
    Dim a As Long
    Dim c As Byte
    
    If MyStatBuffer.xLength > 0 And cmbBinds.ListIndex > 0 Then
        If MsgBox("This object has code. If you bind it to another object, the code is lost." + vbCrLf + "Continue?", vbYesNo, "Confirmation") = vbNo Then
            Cancel = 1
            Exit Sub
        End If
    End If
    
    If Val(txtInfo(0)) < 1 Or Val(txtInfo(0)) > 32767 Then txtInfo(0) = 1
    If Val(txtInfo(1)) < -32768 Or Val(txtInfo(1)) > 32767 Then txtInfo(1) = 0
    If Val(txtInfo(2)) < -32768 Or Val(txtInfo(2)) > 32767 Then txtInfo(2) = 0
    If Val(txtInfo(3)) < -32768 Or Val(txtInfo(3)) > 32767 Then txtInfo(3) = 0
    If Val(txtInfo(4)) < -32768 Or Val(txtInfo(4)) > 32767 Then txtInfo(4) = 0
    If Val(txtInfo(5)) < 0 Or Val(txtInfo(5)) > 53 Then txtInfo(5) = 0
    If Val(txtInfo(6)) < 0 Or Val(txtInfo(6)) > 255 Then txtInfo(6) = 0
    If Val(txtInfo(8)) < 0 Or Val(txtInfo(8)) > 32767 Then txtInfo(8) = 0
    
    If bEdited = True And bCancelEdit = False Then
        With MyStatBuffer
            .xUseStats = True
            .xCycle = Val(txtInfo(0))
            .xXStep = Val(txtInfo(1))
            .xYStep = Val(txtInfo(2))
            .xFollow = Val(txtInfo(3))
            .xLeader = Val(txtInfo(4))
            '.xUnderID = Val(txtInfo(5))
            '.xUnderColor = Val(txtInfo(6))
            .xPointer = Val(txtInfo(7))
            .xInstruction = Val(txtInfo(8))
            .xIsFilled = True
            If cmbBinds.ListIndex > 0 Then
                .xLength = -(cmbBinds.ItemData(cmbBinds.ListIndex))
            End If
            
            If Frame1(0).Enabled = True Then .xP1 = HScroll1(0).Value
            If Frame1(1).Enabled = True Then .xP2 = HScroll1(1).Value
            If Frame1(2).Enabled = True Then .xP3 = HScroll1(2).Value
            If cmbBoard.Visible = True Then .xP3 = cmbBoard.ListIndex
            
            If txtInfo(5).Text <> "" Then
                uI = CByte(txtInfo(5).Text)
            End If
            If txtInfo(6).Text <> "" Then
                uC = CByte(txtInfo(6).Text)
            End If
            
            World.SetObjectInfo1 CurrentBoard, StatNumber, .xXStep, .xYStep, .xCycle, .xP1, .xP2, .xP3
            World.SetObjectInfo2 CurrentBoard, StatNumber, .xFollow, .xLeader, uC, uI, .xPointer, .xInstruction, .xLength
            
        End With
    
        World.SetObjectOOP CurrentBoard, ObjectNum + 0, txtCode1.Text
        MainStatBuffer.xOOP = txtCode1.Text
    End If
    
    World.GetStatLocation CurrentBoard, StatNumber, x, y
    x = x - 1
    y = y - 1
    EnableForms
    MainStatBuffer = MyStatBuffer
    z = World.BoardID(CurrentBoard, x, y)
    a = World.BoardCol(CurrentBoard, x, y)
    If StatNumber >= 0 Then
        frmEdit.chkStats = 1
    End If
    frmEdit.AddToBuffer z, a, World.ObjectCharAt(CurrentBoard, x, y, 1)
    bEdited = False
    bCancelEdit = False
End Sub

Private Sub HScroll1_Change(Index As Integer)
    Label1(Index) = HScroll1(Index).Value
    RefreshPreview
    bEdited = True
End Sub

Private Sub HScroll1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub

Private Sub HScroll1_Scroll(Index As Integer)
    Label1(Index) = HScroll1(Index).Value
    RefreshPreview
End Sub

Private Sub picChar_Paint()
    RefreshPreview
End Sub

Private Sub RefreshPreview()
    Dim x As Long
    Dim y As Long
    If picChar.Visible = True Then
        picChar.ToolTipText = "Char #" + CStr(HScroll1(0).Value) + " (" + Hex(HScroll1(0).Value) + "h)"
        y = 0
        For x = HScroll1(0).Value - 12 To HScroll1(0).Value + 12
            If x >= 0 And x <= 255 Then
                If x = HScroll1(0).Value Then
                    SetChar1 y, 0, x + 0, &H1E, picChar.hdc
                Else
                    SetChar1 y, 0, x + 0, &H16, picChar.hdc
                End If
            Else
                SetChar1 y, 0, 0, 0, picChar.hdc
            End If
            y = y + 1
        Next x
    End If
End Sub

Private Sub PopulateObjects()
    Dim x As Long
    Dim y As String
    For x = 1 To World.StatCount(CurrentBoard)
        y = World.ObjectName(CurrentBoard, x)
        If y <> "" And StatNumber <> x Then
            cmbBinds.AddItem CStr(x) + ":" + y
            cmbBinds.ItemData(cmbBinds.ListCount - 1) = x
            If MyStatBuffer.xLength < 0 Then
                If -(MyStatBuffer.xLength) = x Then
                    cmbBinds.ListIndex = cmbBinds.ListCount - 1
                End If
            End If
        End If
    Next x
    cmbBinds.Enabled = True
End Sub

Private Sub TabStrip_Click()
    RefreshTabs
End Sub

Sub RefreshLibraries()
    Dim ll() As String
    ReDim ll(0 To 0) As String
    Dim x As Long
    ObjectLibrary.GetLibList ll()
    cmbCode1.Clear
    For x = 1 To UBound(ll)
        cmbCode1.AddItem ll(x)
    Next x
End Sub


Private Sub RefreshTabs()
    Dim x As Long
    For x = frameTab.LBound To frameTab.UBound
        frameTab(x).Visible = (TabStrip.SelectedItem.Index = (x + 1))
    Next x
End Sub

Private Sub txtCode1_Change()
    bEdited = True
End Sub

Private Sub txtCode1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
    bEdited = True
    bCancelEdit = False
End Sub

Private Sub txtInfo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
    bEdited = True
End Sub

Private Sub cmdCommand3_Click()
    If cmbCode2.ListIndex > -1 Then
        txtCode1.Text = txtCode1.Text + ObjectLibrary.EntryData(cmbCode2.ItemData(cmbCode2.ListIndex))
    End If
End Sub

Private Sub cmdCommand4_Click()
    Dim x As Long
    Dim y As String
    Dim z As String
    Sound.StopPlaying
    y = UCase$(txtCode1.Text) + vbCrLf
    Do While InStr(y, Chr$(10)) > 0
        z = Left$(y, InStr(y, Chr$(13)) - 1)
        If InStr(z, "#PLAY ") > 0 Then
            If Left$(z, 1) = "/" Or Left$(z, 1) = "#" Then
                Sound.AddToBuffer Mid$(z, InStr(z, "#PLAY ") + 6)
            End If
        End If
        y = Mid$(y, InStr(y, Chr$(10)) + 1)
    Loop
End Sub

Private Sub cmdCommand4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub

Private Sub cmdCommand5_Click()
    Sound.StopPlaying
End Sub

Private Sub cmdCommand5_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub

Private Sub cmdCommand6_Click()
    Sound.StopPlaying
    If MsgBox("Discard changes?", vbYesNo, "Confirmation") = vbYes Then
        bCancelEdit = True
        Unload Me
    End If
End Sub

Private Sub cmdCommand6_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub

Private Sub cmbCode1_Click()
    Dim bl() As String
    Dim x As Long
    ReDim bl(0 To 0) As String
    cmbCode2.Clear
    ObjectLibrary.GetEntryList bl(), cmbCode1.List(cmbCode1.ListIndex)
    For x = 1 To UBound(bl)
        If InStr(bl(x), Chr$(0)) > 0 Then
            cmbCode2.AddItem Left$(bl(x), InStr(bl(x), Chr$(0)) - 1)
            cmbCode2.ItemData(cmbCode2.ListCount - 1) = Val(Mid$(bl(x), InStr(bl(x), Chr$(0)) + 1))
        Else
            cmbCode2.AddItem bl(x)
        End If
    Next x
End Sub

Private Sub cmbCode1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub

Private Sub cmbCode2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub

Private Sub cmdCommand1_Click()
    Sound.StopPlaying
    Unload Me
End Sub

Private Sub cmdCommand1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub

Private Sub cmdCommand2_Click()
    Dim NewFN As String
    CD_ShowOpen_Save Me.hWnd, OFN_EXPLORER Or OFN_FILEMUSTEXIST, NewFN, "", "", "Load Code Library", ZZMZZLfilter, "", True
    If NewFN <> "" Then
        ObjectLibrary.OpenLibrary NewFN
        RefreshLibraries
    End If
End Sub


