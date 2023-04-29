Attribute VB_Name = "Module1"
Option Explicit

Global Const eicar = "X5O!P%@AP[4\PZX54(P^)7CC)7}$EICAR-STANDARD-ANTIVIRUS-TEST-FILE!$H+H*"


Sub createEicar()
    On Error GoTo errorHander
    Dim fh As Long: fh = FreeFile
    Dim s As String
    Dim i As Integer
    For i = 1 To 3
        s = Environ("TEMP") & "\eicar_test_" & i & ".com"
        Open s For Output As #fh
            Print #fh, eicar
        Close #fh
    Next
errorHander:
    log ("createEicar ERROR" & vbTab & Err.Number & vbTab & Err.Description)
End Sub


Sub log(ByVal text As String)
    Debug.Print text
    ' TODO
    
End Sub
