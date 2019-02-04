VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsZZTHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' helpfile resource
' saxxonpike 2oo8

Public Function LoadHelp(sHelpFile As String, bRet() As Byte) As Boolean
    Dim ResNum As Long
    Select Case UCase$(sHelpFile)
        'these help files should be included in the executable as
        'application resources
        Case "ABOUT":       ResNum = 200
        Case "ASCII":       ResNum = 201
        Case "CHEAT":       ResNum = 202
        Case "CREATURE":    ResNum = 203
        Case "CREDITS":     ResNum = 204
        Case "EDITOR":      ResNum = 205
        Case "GAME":        ResNum = 206
        Case "INFO":        ResNum = 207
        Case "ITEM":        ResNum = 208
        Case "KINDS2":      ResNum = 209
        Case "KINDS3":      ResNum = 210
        Case "KINDS":       ResNum = 211
        Case "LANG":        ResNum = 212
        Case "LANGREF":     ResNum = 213
        Case "LANGTUT":     ResNum = 214
        Case "TERRAIN":     ResNum = 215
        Case "WRKTBL":      ResNum = 216
        Case Else
            'look for an external help file
    End Select
    LoadHelp = True
    bRet() = LoadResData(ResNum, "HELP")
End Function
