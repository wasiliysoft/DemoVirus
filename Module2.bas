Attribute VB_Name = "Module2"
Option Explicit
Public Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long
Public gConfig As AppConfigType

Public Type AppConfigType
    pwd As String ' Пароль
End Type

Sub load_Config()
    Dim s As String: s = App.Path + "\pwd.txt"
    Dim pwd As String: pwd = "123"
    
    Dim fh As Long: fh = FreeFile
    Dim fLen As Long
    On Error Resume Next
        fLen = FileLen(s)
    On Error GoTo 0
    
    If fLen > 0 Then
        Open s For Input Access Read As fh
            Seek #fh, 1
            Line Input #fh, pwd
        Close #fh
    End If
    
    gConfig.pwd = pwd
End Sub

