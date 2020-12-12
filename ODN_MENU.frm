VERSION 5.00
Begin VB.Form ODN_MENU 
   Caption         =   "Табличные документы"
   ClientHeight    =   2208
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   4332
   LinkTopic       =   "Form5"
   ScaleHeight     =   2208
   ScaleWidth      =   4332
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Реестр документов"
      Height          =   492
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   4092
   End
   Begin VB.CommandButton Command2 
      Caption         =   "В Ы Х О Д"
      Height          =   492
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   4092
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ОДН"
      Height          =   492
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4092
   End
End
Attribute VB_Name = "ODN_MENU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Odn.Show
Unload ODN_MENU
End Sub

Private Sub Command2_Click()
MainMenu.Visible = True

Unload Me
End Sub

Private Sub Command3_Click()
ReestrTablDoc.Show

Unload Me
End Sub
