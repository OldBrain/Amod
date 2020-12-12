VERSION 5.00
Begin VB.Form RepObor 
   BorderStyle     =   0  'None
   ClientHeight    =   5112
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   426
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   120
      TabIndex        =   14
      Text            =   "Все"
      Top             =   4560
      Width           =   1815
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2520
      TabIndex        =   12
      Text            =   "Все"
      Top             =   3840
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   3840
      Width           =   1815
   End
   Begin VB.OptionButton Option7 
      Caption         =   "Все данные "
      ForeColor       =   &H80000002&
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   2640
      Width           =   1575
   End
   Begin VB.OptionButton Option6 
      Caption         =   "По домам с сумами к возмещению"
      Height          =   615
      Left            =   2040
      TabIndex        =   6
      Top             =   600
      Width           =   2175
   End
   Begin VB.OptionButton Option5 
      Caption         =   "По домам с площадью и количеством прописанных"
      Height          =   615
      Left            =   2040
      TabIndex        =   5
      Top             =   1200
      Width           =   2175
   End
   Begin VB.OptionButton Option4 
      Caption         =   "По домам с кол-вом прописанных"
      Height          =   615
      Left            =   2040
      TabIndex        =   4
      Top             =   1920
      Width           =   2295
   End
   Begin VB.OptionButton Option3 
      Caption         =   "По домам с площадью"
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Caption         =   "Ok"
      Height          =   375
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   2295
   End
   Begin VB.OptionButton Option2 
      Caption         =   "По домам "
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Свернутая"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Благоустроенные"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Подразделение"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   11
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Приватизация"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Image imgTitleHelp 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1049
         SubFormatType   =   0
      EndProperty
      Height          =   195
      Left            =   0
      Picture         =   "RepObor.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   195
   End
   Begin VB.Image imgTitleRight 
      Height          =   360
      Left            =   2640
      Picture         =   "RepObor.frx":024A
      Top             =   0
      Width           =   228
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   """Квартплата + "" параметры отчета"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   120
      TabIndex        =   8
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   3690
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   2040
      Picture         =   "RepObor.frx":0994
      Stretch         =   -1  'True
      ToolTipText     =   "Двойной щелчек мышы развернет форму во весь экран или вернет в исходное состояние"
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleLeft 
      Height          =   360
      Left            =   480
      Picture         =   "RepObor.frx":10DE
      Top             =   0
      Width           =   228
   End
   Begin VB.Line Line5 
      X1              =   296
      X2              =   296
      Y1              =   32
      Y2              =   224
   End
   Begin VB.Line Line4 
      X1              =   8
      X2              =   296
      Y1              =   32
      Y2              =   32
   End
   Begin VB.Line Line3 
      X1              =   8
      X2              =   8
      Y1              =   32
      Y2              =   224
   End
   Begin VB.Line Line2 
      X1              =   8
      X2              =   296
      Y1              =   224
      Y2              =   224
   End
   Begin VB.Line Line1 
      X1              =   128
      X2              =   128
      Y1              =   32
      Y2              =   224
   End
End
Attribute VB_Name = "RepObor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If Me.Option1 Then Reports.Svern = "Свернутая"
If Me.Option2 Then Reports.Svern = "Дома2"

If Me.Option3 Then Reports.Svern = "Дома3"
If Me.Option4 Then Reports.Svern = "Дома4"
If Me.Option5 Then Reports.Svern = "Дома5"
If Me.Option6 Then Reports.Svern = "Дома6"

If Me.Option7 Then Reports.Svern = "Все"

Reports.Прив = Me.Combo1.Text
Reports.Благ = Me.Combo3.Text
Reports.Подр = Me.Combo2.Text
Unload Me
End Sub



Private Sub Form_Load()




Me.Option1.BackColor = RGB(207, 207, 207)
Me.Option2.BackColor = RGB(207, 207, 207)
Me.Option3.BackColor = RGB(207, 207, 207)
Me.Option4.BackColor = RGB(207, 207, 207)
Me.Option5.BackColor = RGB(207, 207, 207)
Me.Option6.BackColor = RGB(207, 207, 207)
Me.Option7.BackColor = RGB(207, 207, 207)
Me.Command1.BackColor = RGB(207, 207, 211)

Me.Combo1.Text = "Все"
Me.Combo1.AddItem ("Все")
Me.Combo1.AddItem ("Да")
Me.Combo1.AddItem ("Нет")
MakeWindow Me, False

Me.Combo2.AddItem ("Все")
Me.Combo2.AddItem ("Подр.№1")
Me.Combo2.AddItem ("Подр.№2")
Me.Combo2.AddItem ("Подр.№3")
Me.Combo2.AddItem ("Подр.№4")
Me.Combo2.AddItem ("Подр.№5")
Me.Combo2.AddItem ("Подр.№6")
Me.Combo2.AddItem ("Подр.№7")
Me.Combo2.AddItem ("Подр.№8")
Me.Combo2.AddItem ("Подр.№9")
Me.Combo2.AddItem ("Подр.№10")

Me.Combo3.AddItem ("Все")
Me.Combo3.AddItem ("Благоустр.")
Me.Combo3.AddItem ("Неблагоустр.")

Combo1.Enabled = False
Combo2.Enabled = False
Combo3.Enabled = False

End Sub

Private Sub Option1_Click()
Combo1.Enabled = False
Combo2.Enabled = False
Combo3.Enabled = False
End Sub

Private Sub Option2_Click()
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True

End Sub

Private Sub Option3_Click()
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True

End Sub

Private Sub Option4_Click()
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True

End Sub

Private Sub Option5_Click()
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True

End Sub

Private Sub Option6_Click()
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True

End Sub

Private Sub Option7_Click()
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True

End Sub
