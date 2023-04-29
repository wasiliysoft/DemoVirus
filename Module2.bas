Attribute VB_Name = "Module2"
Option Explicit

Public gConfig As AppConfigType

Public Type AppConfigType
    userCode As Integer
    adminCode As Integer
    isCreateEicar As Integer
    isWaiteAdmin As Integer
    isLogging As Integer
    logPath As String
End Type

Sub load_Config()
    Dim s As String: s = App.Path + "\config.txt"
    
    With gConfig
        .userCode = 111
        .adminCode = 222
        .isWaiteAdmin = 1
        .isCreateEicar = 0
        .isLogging = 1
        .logPath = Environ("USERPROFILE") & "\"
    End With
    
    Dim fh As Long: fh = FreeFile
    Dim fLen As Long
    On Error Resume Next
        fLen = FileLen(s)
    On Error GoTo 0
    
    If fLen > 0 Then
        Open s For Input Access Read As #fh
            Seek #fh, 1
            Line Input #fh, s
            gConfig.userCode = Val(s)
            Line Input #fh, s
            gConfig.adminCode = Val(s)
            Line Input #fh, s
            gConfig.isCreateEicar = Val(s)
            Line Input #fh, s
            gConfig.isWaiteAdmin = Val(s)
            Line Input #fh, s
            gConfig.isLogging = Val(s)
            Line Input #fh, s
            gConfig.logPath = s
        Close #fh
    End If
End Sub

