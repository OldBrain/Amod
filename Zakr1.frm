VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Zakr1 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Закрытие расчетного периода"
   ClientHeight    =   1860
   ClientLeft      =   168
   ClientTop       =   552
   ClientWidth     =   4920
   ControlBox      =   0   'False
   FillColor       =   &H0000C0C0&
   ForeColor       =   &H0080FFFF&
   Icon            =   "Zakr1.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   4920
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   4695
      _ExtentX        =   8276
      _ExtentY        =   656
      _Version        =   393216
      Appearance      =   1
      Min             =   1e-4
      Max             =   1100
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Отмена"
      Height          =   615
      Left            =   1560
      TabIndex        =   1
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Если Вы хотите закрыть период расчета то нажмите <Shift>/<Ctrl>+<F12>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4455
   End
   Begin VB.Menu Меню 
      Caption         =   "Меню"
      Begin VB.Menu Закрыть 
         Caption         =   "Закрыть период"
         Shortcut        =   +^{F12}
      End
      Begin VB.Menu Выход 
         Caption         =   "Выход"
      End
   End
End
Attribute VB_Name = "Zakr1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
Unload Me
MainMenu.Enabled = True
End Sub

Private Sub Form_Load()
MainMenu.Enabled = False
Zakr1.ProgressBar1.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
MainMenu.Enabled = True
End Sub

Private Sub Закрыть_Click()
'MsgBox ("Закрываем")
MainForm.Zakritie

End Sub
