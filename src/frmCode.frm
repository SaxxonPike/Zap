VERSION 5.00
Begin VB.Form frmCode 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Code"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5415
   ControlBox      =   0   'False
   Icon            =   "frmCode.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   5415
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCommand6 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Playback"
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Import"
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   5175
      Begin VB.ComboBox cmbCode2 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Within the library, you can select code fragments here."
         Top             =   600
         Width           =   3975
      End
      Begin VB.CommandButton cmdCommand2 
         Caption         =   "Load Library..."
         Height          =   255
         Left            =   3360
         TabIndex        =   2
         ToolTipText     =   "Load an object library from disk."
         Top             =   240
         Width           =   1695
      End
      Begin VB.ComboBox cmbCode1 
         Height          =   315
         ItemData        =   "frmCode.frx":038A
         Left            =   120
         List            =   "frmCode.frx":038C
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Select your library here."
         Top             =   240
         Width           =   3135
      End
      Begin VB.CommandButton cmdCommand3 
         Caption         =   "Insert"
         Height          =   255
         Left            =   4200
         TabIndex        =   4
         ToolTipText     =   "Insert code from a library into this object's code."
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdCommand1 
      Caption         =   "OK"
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   6240
      Width           =   1215
   End
   Begin VB.TextBox txtCode1 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
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
      TabIndex        =   0
      Top             =   1200
      Width           =   5415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   5880
      Width           =   2775
   End
End
Attribute VB_Name = "frmCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ObjectNum As Long
Dim bEdited As Boolean
Dim bCancelEdit As Boolean
Const ZZMZZLfilter = "All Openable Files|*.ZZM;*.ZZL;*.TXT|ZZM Music Library (*.ZZM)|*.ZZM|ZZL Object Library (*.ZZL)|*.ZZL|Text File (*.TXT)|*.TXT|All Files|*.*"

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

Private Sub Form_Load()
    'frmEdit.Enabled = False
    'frmPalette.Enabled = False
    DisableForms
    RefreshLibraries
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim x As Long
    Dim y As Long
    Dim z As Long
    Dim a As Long
    If bEdited = True And bCancelEdit = False Then
        World.SetObjectOOP CurrentBoard, ObjectNum, txtCode1.Text
        World.GetStatLocation CurrentBoard, ObjectNum, x, y
        MainStatBuffer.xOOP = txtCode1.Text
        x = x - 1
        y = y - 1
        z = World.BoardID(CurrentBoard, x, y)
        a = World.BoardCol(CurrentBoard, x, y)
        If ObjectNum >= 0 Then
            frmEdit.chkStats = 1
        End If
        frmEdit.AddToBuffer z, a, World.ObjectCharAt(CurrentBoard, x, y, 1)
    End If
    bEdited = False
    bCancelEdit = False
    EnableForms
End Sub

Public Sub SetObjectNumber(z As Long)
    Dim x As Long
    Dim y As Long
    Dim a As Long
    If World.ObjectLength(CurrentBoard, z) < 0 Then
        Select Case MsgBox("This object has been bound to another object. Would you like to duplicate its code?" + vbCrLf + "Click YES to duplicate. Click NO to create new code. Click CANCEL to keep the bind.", vbYesNoCancel, "Confirmation")
            Case vbYes
                World.SetObjectOOP CurrentBoard, z, World.ObjectOOP(CurrentBoard, -(World.ObjectLength(CurrentBoard, z)))
            Case vbNo
                World.SetObjectOOP CurrentBoard, z, ""
            Case vbCancel
                Unload Me
                Exit Sub
        End Select
    End If
    ObjectNum = z
    World.GetStatLocation CurrentBoard, z, x, y
    x = x - 1
    y = y - 1
    If World.BoardID(CurrentBoard, x, y) = E_Scroll Then
        lblCode = "Scroll"
    Else
        lblCode = "Object"
    End If
    lblCode = lblCode + " #" + CStr(z) + " at (" + CStr(x) + "," + CStr(y) + ")"
    txtCode1 = World.ObjectOOP(CurrentBoard, z)
    bEdited = True
End Sub

Private Sub txtCode1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub
