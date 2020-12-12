VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Справка"
   ClientHeight    =   5130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7365
   LinkTopic       =   "Form2"
   ScaleHeight     =   5130
   ScaleWidth      =   7365
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Закрыть"
      Height          =   495
      Left            =   2640
      TabIndex        =   0
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000018&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   6975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Андрей Бугоров  2004 г.8-905-360-4006"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   4800
      Width           =   7335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Label2_Click()
Form2.Hide
End Sub
