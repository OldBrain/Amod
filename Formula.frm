VERSION 5.00
Begin VB.Form Formula 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3735
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   9960
   ControlBox      =   0   'False
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   249
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   664
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Text            =   "SummaI"
      Top             =   2160
      Width           =   9735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ввод"
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Text            =   "SummaI"
      Top             =   960
      Width           =   9735
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
      Height          =   225
      Left            =   0
      Picture         =   "Formula.frx":0000
      Top             =   0
      Width           =   225
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Окно ввода формул"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   0
      TabIndex        =   5
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   9810
   End
   Begin VB.Image imgTitleLeft 
      Height          =   450
      Left            =   4800
      Picture         =   "Formula.frx":0436
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleRight 
      Height          =   450
      Left            =   5160
      Picture         =   "Formula.frx":0B80
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   4560
      Picture         =   "Formula.frx":12CA
      Stretch         =   -1  'True
      ToolTipText     =   "Двойной щелчек мышы развернет форму во весь экран или вернет в исходное состояние"
      Top             =   0
      Width           =   285
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Формула без учета льгот"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   9735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Формула с учетом льгот"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   9735
   End
End
Attribute VB_Name = "Formula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Public Fr As String
'Public FrBl As String
Private KodN As String
'Dim Mconn As ADODB.Connection





Private Sub Command1_Click()


'Nachisleniy.Show
Fr = Trim(Formula.Text1)
FrBl = Trim(Formula.Text2)
'Set Mconn = ADODB.Connection
'MsgBox (FrBl)
On Error GoTo Er
SF1 = "UPDATE Adding_Err SET Adding_Err.SummaI =" + Fr
'SF1 = "UPDATE Adding_Err SET Adding_Err.SummaI = 0"
Mconn.Execute (SF1)

Mconn.Execute ("UPDATE Adding_Err SET Adding_Err.SummaI = " + FrBl)
'MsgBox "Добавил " + FrBl
Nachisleniy.FI = Trim(Fr)
Nachisleniy.FIBl = Trim(FrBl)

'mconn.Execute ("UPDATE nachisleniy SET nachisleniy.Formula = " + Chr(34) + Fr + Chr(34) + " , nachisleniy.FormulaB = " + Chr(34) + FrBl + Chr(34) + " WHERE (((nachisleniy.Kod)=" + KodN + "))")


Nachisleniy.FG1.TextMatrix(Nachisleniy.FG1.Row, 10) = Trim(Formula.Text2)
Nachisleniy.FG1.TextMatrix(Nachisleniy.FG1.Row, 5) = Fr

Unload Me
'Nachisleniy.Show

'Nachisleniy.Enabled = True

'Nachisleniy.FG1.TextMatrix(Nachisleniy.FG1.Row, 10) = Nachisleniy.FIBl

Exit Sub
Er:
MsgBox ("Ошибка в формуле" + Err.Description)
'FG1.TextMatrix(FG1.Row, 5) = Old

'Formula.Show
Formula.Text1 = Fr
Formula.Text2 = FrBl
Text1.SetFocus
'Nachisleniy.Enabled = False

End Sub

Private Sub Form_Load()
MakeWindow Me, True


'Set mconn = New ADODB.Connection
'mconn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Persist Security Info=True"
'mconn.Open "data/Kvartplata.mdb"
KodN = Nachisleniy.FG1.TextMatrix(Nachisleniy.FG1.Row, 1)
Text2 = Nachisleniy.FG1.TextMatrix(Nachisleniy.FG1.Row, 10)

End Sub

Private Sub Form_Unload(Cancel As Integer)
'Nachisleniy.FG1.TextMatrix(Nachisleniy.FG1.Row, 10) = Trim(Formula.Text2)
'Nachisleniy.FG1.TextMatrix(Nachisleniy.FG1.Row, 5) = Fr
'Nachisleniy.Enabled = True
End Sub

Private Sub imgTitleHelp_Click()
Command1_Click
End Sub

Private Sub imgTitleHelp_DblClick()
Command1_Click
End Sub

Private Sub Text1_LostFocus()
'Nachisleniy.FG1.TextMatrix(Nachisleniy.FG1.Row, Nachisleniy.FG1.Col) = Text1
End Sub

Private Sub Text2_LostFocus()
FrBl = Trim(Text2)
End Sub
