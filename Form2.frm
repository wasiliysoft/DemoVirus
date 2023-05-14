VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00008000&
   BorderStyle     =   0  'None
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12690
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   12690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "Form2.frx":0000
      Top             =   2040
      Width           =   12135
   End
   Begin VB.Timer Timer1 
      Left            =   360
      Top             =   6600
   End
   Begin VB.CommandButton unlockBtn 
      BackColor       =   &H0000C000&
      Caption         =   "ОК"
      Default         =   -1  'True
      Height          =   495
      Left            =   4800
      TabIndex        =   1
      Top             =   720
      Width           =   1215
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
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   720
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00008000&
      Caption         =   "Код отмены (Администратор)"
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
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   5055
   End
   Begin VB.Label Label4 
      BackColor       =   &H00008000&
      Caption         =   "Полезные советы:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackColor       =   &H00008000&
      Caption         =   "Дождитесь Администратора ИБ, он введет свой код отмены."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   5880
      Width           =   12135
   End
   Begin VB.Label Label2 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   4335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    log ("Показ формы Администратора ИБ")
    
    Dim s As String: s = App.Path + "\notes.txt"
    
    Dim fh As Long: fh = FreeFile
    Dim fLen As Long
    On Error Resume Next
        fLen = FileLen(s)
    On Error GoTo 0
    
    If fLen > 0 Then
        Text2.text = ""
        Open s For Input Access Read As #fh
            Do Until (EOF(fh))
                Line Input #fh, s
                Text2.text = Text2.text & s & vbNewLine
            Loop
        Close #fh
    End If
End Sub

Private Sub Form_Terminate()
    log ("Закрыта форма Администратора ИБ")
End Sub



Private Sub txtPass_KeyDown(KeyCode As Integer, Shift As Integer)
    Label2.Caption = ""
End Sub

Private Sub unlockBtn_Click()
    If (Val(gConfig.adminCode) = Val(txtPass.text)) Then
        log ("Введен код отмены Администратора ИБ")
        onCorrectPass
        Unload Me
    Else
        log ("Неправильный ввод кода отмены Администратора ИБ")
        onIncorrectPass
    End If
End Sub
Private Sub onCorrectPass()
    Label2.Caption = ""
    txtPass.text = ""
End Sub
Private Sub onIncorrectPass()
    Label2.Caption = "Неправильный код!"
    txtPass.text = ""
    txtPass.SetFocus
End Sub

