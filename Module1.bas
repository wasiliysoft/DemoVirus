Attribute VB_Name = "Module1"
Option Explicit

Global Const eicar = "X5O!P%@AP[4\PZX54(P^)7CC)7}$EICAR-STANDARD-ANTIVIRUS-TEST-FILE!$H+H*"


Sub createEicar()
    On Error GoTo errorHander
    Dim fh As Long: fh = FreeFile
    Dim s As String
    Dim i As Integer
    For i = 1 To 3
        s = Environ("TEMP") & "\eicar_by_exe_test_" & i & ".com"
        Open s For Output As #fh
            Print #fh, eicar
        Close #fh
    Next
    log ("Создан тестовый файл угрозы")
    Exit Sub
errorHander:
    log ("createEicar ERROR" & vbTab & Err.Number & vbTab & Err.Description)
End Sub


Public Sub log(text As String)
    Dim s As String: s = Now() & vbTab & text
       
    If (gConfig.isLogEnabled = False) Then
        Debug.Print s
        Exit Sub
    End If
    Dim sPath       As String
    sPath = gConfig.logPath
    Dim fh          As Long: fh = FreeFile
    
    On Error GoTo errorHandler
        Open sPath & Environ("USERNAME") & "_" & Environ("COMPUTERNAME") & ".log" For Append As fh
        Print #fh, s
        Close #fh
    On Error GoTo 0
    Exit Sub
    
errorHandler:
    Debug.Print Err.Number
    Select Case Err.Number
    Case 75: sPath = Environ("USERPROFILE") & "\": Resume
    Case 76: sPath = Environ("USERPROFILE") & "\": Resume
    Case Else
        Debug.Print Now(), "error", Err.Number, Err.Description
        Exit Sub
    End Select
End Sub
