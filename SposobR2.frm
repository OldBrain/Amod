VERSION 5.00
Begin VB.Form SposobR2 
   Caption         =   "Способ расчета"
   ClientHeight    =   1605
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4470
   ControlBox      =   0   'False
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Отмена"
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Без учета исправлений"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "С учетом исправлений"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Расчет"
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
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "SposobR2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload SposobR
Расчет2.Isprav = 1 'С учетом исправлений
Расчет2.Show
End Sub

Private Sub Command2_Click()
Unload SposobR
Расчет2.Isprav = 2 'Без учета исправлений
Расчет2.Show
End Sub

Private Sub Command3_Click()
Doc.Enabled = True
Unload Me
End Sub

Private Sub Form_Load()
Doc.Enabled = False
End Sub
