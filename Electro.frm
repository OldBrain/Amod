VERSION 5.00
Begin VB.Form Electro 
   ClientHeight    =   2325
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form8"
   ScaleHeight     =   2325
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Отмена"
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Сохранить"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   4575
   End
   Begin VB.Label Label1 
      Caption         =   "Тариф на электроэнергию для"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "Electro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
'Kvart.Q = "SELECT KLS_PODR.NAIM_KLS, Adding.Electro FROM KLS_PODR INNER JOIN (Adding INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer) ON KLS_PODR.КОД = MainOccupant.Dom WHERE (((KLS_PODR.NAIM_KLS)=" + Filter.FG.TextMatrix(Filter.FG.Row, 5) + ") AND ((Adding.Electro)=" + Text1 + "))"
Kvart.Q = "UPDATE (KLS_PODR INNER JOIN MainOccupant ON KLS_PODR.КОД = MainOccupant.Dom) INNER JOIN Adding ON MainOccupant.Numer = Adding.KodKv SET Adding.Electro = " + Text1 + " WHERE (((KLS_PODR.NAIM_KLS)=" + Chr(34) + Filter.FG.TextMatrix(Filter.FG.Row, 5) + Chr(34) + "))"
Kvart.T = Text1
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Label2 = Kvart.Combo2
Text1 = 0
End Sub
