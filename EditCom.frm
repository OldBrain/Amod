VERSION 5.00
Begin VB.Form EditCom 
   ClientHeight    =   1725
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   7350
   ControlBox      =   0   'False
   LinkTopic       =   "Form7"
   Moveable        =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   6975
   End
End
Attribute VB_Name = "EditCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Doc.FG.TextMatrix(Doc.FG.Row, 11) = Text1
Unload Me
End Sub

Private Sub Form_Load()
Text1 = Doc.FG.TextMatrix(Doc.FG.Row, 11)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Doc.Enabled = True
End Sub
