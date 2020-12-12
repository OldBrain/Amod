VERSION 5.00
Begin VB.Form Zakr 
   BackColor       =   &H0000FFFF&
   Caption         =   "Закрытие расчетного периода"
   ClientHeight    =   5100
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6975
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Отмена"
      Height          =   615
      Left            =   2040
      TabIndex        =   0
      Top             =   4080
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Закрыть"
      Height          =   615
      Left            =   2040
      TabIndex        =   3
      Top             =   3240
      Width           =   3015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Закрыть расчетный период?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   2520
      Width           =   5535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"Zakr.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6735
   End
End
Attribute VB_Name = "Zakr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
Zakr1.Show
End Sub

Private Sub Command2_Click()
Unload Me
MainMenu.Enabled = True
End Sub

Private Sub Form_Load()
MainMenu.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
MainMenu.Enabled = True
End Sub
