VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000080&
   BorderStyle     =   0  'None
   Caption         =   "Киоск"
   ClientHeight    =   7275
   ClientLeft      =   0
   ClientTop       =   105
   ClientWidth     =   11115
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   11115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000C000&
      Height          =   3735
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   10455
   End
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   840
      Top             =   6120
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   360
      Top             =   3600
   End
   Begin VB.CommandButton btnUnlock 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "РАЗБЛОКИРОВАТЬ"
      Default         =   -1  'True
      Height          =   615
      Left            =   8280
      MaskColor       =   &H000000C0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
   Begin VB.TextBox txtPass 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   360
      MaxLength       =   16
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   480
      Width           =   6855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000080&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   5280
      Width           =   10575
   End
   Begin VB.Label LabelLang 
      BackColor       =   &H00000080&
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7320
      TabIndex        =   4
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000080&
      Caption         =   "Пароль разблокировки рабочего стола"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000080&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   1080
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private counter As Integer

Private Sub btnLogoff_Click()
    If (logOn(Trim(CStr(txtPass.Text)))) Then
        onCorrectPass
        Shell "LOGOFF"
    Else
        onIncorrectPass
    End If
End Sub

Private Sub btnReboot_Click()
    If (logOn(Trim(CStr(txtPass.Text)))) Then
        onCorrectPass
        Shell "shutdown -r -t 0 -f"
    Else
        onIncorrectPass
    End If
End Sub

Private Sub Form_Load()
    MsgBox "ПОЖАЛУЙСТА НЕ ЗАКРЫВАЙТЕ ЭТО ОКНО ПОКА НЕ СОХРАНИТЕ ВСЕ ФАЙЛЫ!", vbExclamation
    load_Config
    updateLangIndicator
    
End Sub

Private Sub btnUnlock_Click()
    If (logOn(Trim(CStr(txtPass.Text)))) Then
        onCorrectPass
        Unload Me
    Else
        onIncorrectPass
    End If
End Sub
Private Sub onCorrectPass()
        Label2.Caption = ""
        txtPass.Text = ""
End Sub
Private Sub onIncorrectPass()
        Label2.Caption = "Неправильный пароль"
        txtPass.SetFocus
        SendKeys "{Home}+{End}"
End Sub
Private Sub updateLangIndicator()
    LabelLang.Caption = IIf(GetKeyboardLayout(0) = 67699721, "EN", "RU")
End Sub

Private Sub Timer1_Timer()
    counter = counter + 1

    Label1.Caption = "Не выключайте компьютер, выполняется шифрование: " & Format(counter / 100, "0.00") & "%"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    updateLangIndicator
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    updateLangIndicator
End Sub

Function logOn(ByVal pass As String) As Boolean
    logOn = False
    pass = Trim(pass)

    If (Trim(gConfig.pwd) = pass) Then
        logOn = True
    End If
End Function


Private Sub Timer2_Timer()
       List1.AddItem "Шифрование C:\User\" & Environ("username") & "\Рабочий стол\" & randomStr & "       OK"
       List1.ListIndex = List1.ListCount - 1
End Sub

Function randomStr()
    Dim r As Integer
    Dim r2 As Integer
   
    r = Rnd(15) * 20
   
    
    Dim i As Integer
    For i = 1 To r
        r2 = Rnd(15) * 60
       
        
        randomStr = randomStr & Chr(190 + r2)
    Next i

    Dim r3 As Integer
    r3 = Rnd * 2
    
    If (r3 = 2) Then
        randomStr = randomStr & ".docx"
    Else
        randomStr = randomStr & ".xlsx"
    End If
End Function
















