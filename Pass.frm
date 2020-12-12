VERSION 5.00
Begin VB.Form Pass 
   BackColor       =   &H000000FF&
   Caption         =   "Пароль"
   ClientHeight    =   3165
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   4680
   LinkTopic       =   "Form8"
   ScaleHeight     =   3165
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Далее"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   2520
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Пароль"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "!!Использование этого раздела программы может привести к необратимым изменениям в базе данных.!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
   Begin VB.Menu a 
      Caption         =   "1"
      HelpContextID   =   1
      Index           =   1
      Begin VB.Menu s 
         Caption         =   "2"
         HelpContextID   =   2
         Index           =   2
         Shortcut        =   +{DEL}
      End
   End
End
Attribute VB_Name = "Pass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
MenuNastr1.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
MainMenu.Enabled = True
End Sub

Private Sub s_Click(Index As Integer)
Command1_Click
End Sub
