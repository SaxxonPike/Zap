VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsZZTLibrary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' **************************************************************
'  clsZZTLibrary
'
'  ZZT object/music library support
'  SaxxonPike 2oo7-2oo8
' **************************************************************

Private Type xtiLibData
    xLibrary As String
    xTitle As String
    xInfo As String
    xData As String
End Type

Dim xLibData() As xtiLibData
Dim xEntryCount As Long
Dim xFileOffset As Long

Public Function OpenLibrary(lname As String) As String
    Dim lfile As String
    Dim ext As String
    Dim f As Long
    Dim s As String
    Dim t As String
    Dim n As String
    Dim i As String
    Dim entries(1000) As String
    Dim enames(1000) As String
    Dim ec As Long
    Dim ol As Long
    Dim st As Long
    Dim st2 As Double
    xFileOffset = 1
    If InStr(lname, "\") > 0 Then
        lfile = Mid$(lname, InStrRev(lname, "\") + 1)
    Else
        lfile = lname
    End If
    If InStr(lfile, ".") > 0 Then
        ext = UCase$(Mid$(lfile, InStrRev(lfile, ".") + 1))
    Else
        ext = "TXT"
    End If
    f = FreeFile
    If Dir(lname) = "" Then
        Exit Function
    End If
    Open lname For Binary As #f
    
    Select Case ext
        Case "TXT" '*** text file ***
            Do While Not EOF2(f)
                GetNextLine f, s
                t = t + s + vbCrLf
            Loop
            AddEntry lfile, lfile, "", t
        Case "ZZM" '*** zzm music library ***
            ol = 0
            Do While Not EOF2(f)
                GetNextLine f, s
                s = Trim$(s)
                If Right$(s, 1) = ";" Then
                    s = UCase$(s)
                End If
                If Left$(s, 14) = "; $SONG TITLE " Then
                    n = Mid$(s, 15)
                    ol = Val(Left$(n, InStr(n, " ") - 1))
                    n = Mid$(n, InStr(n, " ") + 1)
                    enames(ol - 1) = n
                ElseIf Left$(s, 12) = "; $SONG ENDS" Then
                    entries(ol - 1) = t
                    t = ""
                    n = ""
                    ol = 0
                    st2 = 0
                ElseIf Left$(s, 8) = "; $SONG " Then
                    n = Mid$(s, 9)
                    t = "#cycle 1" + vbCrLf
                    ol = Val(n)
                    If ol > ec Then
                        ec = ol
                    End If
                    If enames(ol - 1) = "" Then
                        enames(ol - 1) = "Untitled song " + n
                    End If
                ElseIf Left$(s, 1) <> ";" Then
                    If ol > 0 Then
                        t = t + "#play " + s + vbCrLf
                        st2 = st2 + GetSubNotes(s)
                        Do While st2 > 42
                            st2 = st2 - 42
                            t = t + "/i/i/i/i/i/i/i/i/i/i/i/i/i/i/i/i/i/i/i/i/i" + vbCrLf
                        Loop
                        For st = 1 To (st2 \ 2)
                            t = t + "/i"
                        Next st
                        If (st2 \ 2) > 0 Then
                            t = t + vbCrLf
                            st2 = st2 - ((st2 \ 2) * 2)
                        End If
                    End If
                End If
            Loop
        Case "ZZL" '*** zzl object library ***
            GetNextLine f, s 'library name, we won't use it
            Do While Not EOF2(f) 'bypass comment section
                GetNextLine f, s
                If Left$(s, 1) <> "*" Then
                    Exit Do
                End If
            Loop
            st = 0
            Do While Not EOF2(f)
                If ol = 0 And st = 2 Then
                    st = 0
                    entries(ec) = t
                    enames(ec) = n
                    t = ""
                    i = ""
                    ec = ec + 1
                    n = "Untitled Object " + CStr(ec)
                End If
                If st = 0 Then
                    'object name
                    If s <> "" Then
                        n = s
                        Debug.Print n
                    End If
                    st = st + 1
                ElseIf st = 1 Then
                    'object info
                    i = s
                    ol = Val(Left$(i, InStr(i, ",") - 1))
                    st = st + 1
                ElseIf st = 2 Then
                    t = t + s + vbCrLf
                    ol = ol - 1
                End If
                GetNextLine f, s
            Loop
    End Select
    For ol = 0 To ec - 1
        AddEntry lfile, enames(ol), "", entries(ol)
    Next ol
    
    Close #f
        
    'Loop
    OpenLibrary = lfile
End Function

Public Sub GetEntryList(retEntries() As String, xLib As String)
    Dim x As Long
    Dim ec As Long
    For x = 1 To xEntryCount
        If UCase$(xLibData(x).xLibrary) = UCase$(xLib) Then
            ec = ec + 1
            ReDim Preserve retEntries(0 To ec) As String
            retEntries(ec) = xLibData(x).xTitle + Chr$(0) + CStr(x)
        End If
    Next x
    
End Sub

Public Sub GetLibList(retLibs() As String)
    Dim x As Long
    Dim ec As Long
    For x = 1 To xEntryCount
        If xLibData(x).xLibrary <> xLibData(x - 1).xLibrary Then
            ec = ec + 1
            ReDim Preserve retLibs(0 To ec) As String
            retLibs(UBound(retLibs)) = xLibData(x).xLibrary
        End If
    Next x
End Sub

Private Sub AddEntry(xLib As String, xTi As String, xInf As String, xDa As String)
    Dim x As Long
    xEntryCount = xEntryCount + 1
    ReDim Preserve xLibData(0 To xEntryCount) As xtiLibData
    With xLibData(UBound(xLibData))
        .xLibrary = xLib
        .xData = xDa
        .xInfo = xInf
        .xTitle = xTi
    End With
End Sub

Public Function EntryData(xEntryNumber As Long)
    EntryData = xLibData(xEntryNumber).xData
End Function

Private Sub Class_Initialize()
    ReDim xLibData(0 To 0) As xtiLibData
End Sub

Private Sub GetNextLine(f As Long, ByRef s As String)
    Dim b As Byte
    s = ""
    Do While xFileOffset < LOF(f)
        Get #f, xFileOffset, b
        xFileOffset = xFileOffset + 1
        If b <> 13 And b <> 10 Then
            s = s + Chr$(b)
        Else
            Exit Do
        End If
    Loop
    If xFileOffset < LOF(f) Then
        Get #f, xFileOffset, b
        If b = 13 Or b = 10 Then
            xFileOffset = xFileOffset + 1
        End If
    End If
End Sub

Private Function EOF2(f As Long) As Boolean
    If xFileOffset >= LOF(f) Then
        EOF2 = True
    End If
End Function

Private Function GetSubNotes(iString As String) As Double
    Dim x As Long
    Dim cl As Double
    Dim tr As Long
    cl = 1 'default to 32nd
    For x = 1 To Len(iString)
        Select Case LCase$(Mid$(iString, x, 1))
            Case "a", "b", "c", "d", "e", "f", "g", "x", "0", "1", "2", "4", "5", "6", "7", "8", "9"
                If tr > 3 Then
                    tr = tr - 1
                End If
                If tr = 0 Then
                    GetSubNotes = GetSubNotes + cl
                End If
            Case "w"
                cl = 32
            Case "h"
                cl = 16
            Case "q"
                cl = 8
            Case "i"
                cl = 4
            Case "s"
                cl = 2
            Case "t"
                cl = 1
            Case "."
                cl = cl * 1.5
            Case "3"
                tr = 3
        End Select
    Next x
End Function
