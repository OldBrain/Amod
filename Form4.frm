VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Меню"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   FillColor       =   &H00FFC0C0&
   ForeColor       =   &H80000001&
   LinkTopic       =   " "
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Универсальный отчет"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   4335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Назад"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   2760
      Width           =   4335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Резерв"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   2340
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Возмещение льгот"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form4.Hide
Form11.Show
End Sub

Private Sub Command2_Click()
Form4.Hide
Form111.Show
End Sub

Private Sub Command3_Click()
Form4.Hide
Form1.Show
End Sub

Private Sub Command4_Click()
Form11.FG.Refresh
Form4.Hide
Form6.Show
End Sub

