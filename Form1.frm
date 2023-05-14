VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000080&
   BorderStyle     =   0  'None
   Caption         =   "ТР"
   ClientHeight    =   9225
   ClientLeft      =   0
   ClientTop       =   105
   ClientWidth     =   11820
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   11820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer3 
      Interval        =   30000
      Left            =   3360
      Top             =   8400
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000C000&
      Height          =   3735
      Left            =   360
      TabIndex        =   5
      Top             =   1920
      Width           =   10695
   End
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   2400
      Top             =   8400
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   1680
      Top             =   8400
   End
   Begin VB.CommandButton btnUnlock 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "ОСТАНОВИТЬ"
      Default         =   -1  'True
      Height          =   495
      Left            =   8520
      MaskColor       =   &H000000C0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
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
      Top             =   840
      Width           =   7695
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000080&
      Caption         =   $"Form1.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   1095
      Left            =   360
      TabIndex        =   6
      Top             =   6480
      Visible         =   0   'False
      Width           =   11055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000080&
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
      Left            =   360
      TabIndex        =   4
      Top             =   5760
      Width           =   10335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000080&
      Caption         =   "Код отмены"
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
      Left            =   360
      TabIndex        =   3
      Top             =   360
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000080&
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
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   7695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private counter As Integer

Private Sub Form_Initialize()
    MsgBox "ПОЖАЛУЙСТА НЕ ЗАКРЫВАЙТЕ ЭТО ОКНО ПОКА НЕ СОХРАНИТЕ ВСЕ ФАЙЛЫ!", vbExclamation
End Sub

Private Sub Form_Load()
    load_Config
    log ("Показ формы пользователя")
    log ("Режим ожидания Администратора ИБ" & vbTab & gConfig.isWaiteAdmin)
    If (gConfig.isCreateEicar) Then createEicar
End Sub

Private Sub btnUnlock_Click()
    If (Val(gConfig.userCode) = Val(txtPass.text)) Then
        log ("Введен код отмены пользователя")
        onCorrectPass
        Form2.Show
        
        Unload Me
        log ("Закрыта форма пользователя")
    Else
        log ("Неправильный ввод кода отмены пользователя")
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

Private Sub Timer1_Timer()
    counter = counter + 1
    Label1.Caption = "Не выключайте компьютер, выполняется шифрование: " & Format(counter / 100, "0.00") & "%"

    If (counter > 20) Then
        Label4.Visible = counter Mod 3
    End If
End Sub

Private Sub Timer2_Timer()
       List1.AddItem "Шифрование C:\User\" & Environ("username") & "\Рабочий стол\" & randomStr & ".............OK"
       If (List1.ListCount > 50) Then List1.RemoveItem (0)
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

Private Sub Timer3_Timer()
    log ("Таймер 30с ОК")
End Sub

Private Sub txtPass_KeyDown(KeyCode As Integer, Shift As Integer)
  Label2.Caption = ""
End Sub
