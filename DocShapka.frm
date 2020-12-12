VERSION 5.00
Begin VB.Form DocShapka 
   Caption         =   "Параметры документа"
   ClientHeight    =   2940
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   9048
   ControlBox      =   0   'False
   Icon            =   "DocShapka.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   9048
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   4
      Text            =   " "
      ToolTipText     =   "Коментарий к документу, любая текстовая информация"
      Top             =   1800
      Width           =   8895
   End
   Begin VB.TextBox Text1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd.MM.yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1049
         SubFormatType   =   3
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      ToolTipText     =   "Дата документа"
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Отмена"
      Height          =   375
      Left            =   7080
      TabIndex        =   6
      ToolTipText     =   "Отказ от создания документа"
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Далее"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "Создать документ"
      Top             =   2400
      Width           =   1815
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1320
      TabIndex        =   3
      Text            =   "Combo3"
      ToolTipText     =   "Адрес. Можно выбрать любой адрес, в последствии будет возможность проставлять адрес внутри документа по <F2>"
      Top             =   960
      Width           =   7695
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2280
      TabIndex        =   2
      Text            =   "Combo2"
      ToolTipText     =   "Начисление по умолчанию, если вабрать ""Любое начисление"". то в документе будет возможность ввода начисления вручную."
      Top             =   480
      Width           =   5415
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2880
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   2400
      Width           =   390
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Содержание (строка коментария)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   8775
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Дата документа:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Адрес:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Начисление / удержание:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "DocShapka"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim mconn As ADODB.Connection
Dim Combo_RS, Combo_rs1 As ADODB.Recordset
Private Sub Combo2_LostFocus()
If Trim(Combo2.Text) = "" Then Combo2.SetFocus
End Sub
Private Sub Combo3_LostFocus()
If Trim(Combo3.Text) = "" Then Combo3.SetFocus
End Sub

Private Sub Combo1_LostFocus()
If Trim(Combo1.Text) = "" Then Combo1.SetFocus
End Sub

Private Sub Combo1_Validate(Cancel As Boolean)
If Combo1.Text = "Выбери тип" Then Tip = ""
'MsgBox ("Вы не выбрали тип документа")
'Combo1.SetFocus

If Combo1.Text = "Начисление" Then Tip = "+"
If Combo1.Text = "Оплата или субсидия" Then Tip = "-"
End Sub

Private Sub Command1_Click()
ReestrDoc.Новый
'ReestrDoc.Hide
Unload Me
Unload ReestrDoc
ReestrDoc.Show
ReestrDoc.Enabled = True
End Sub

Private Sub Command2_Click()
ReestrDoc.Enabled = True
Unload Me
End Sub

Private Sub Form_Load()
Combo1.Text = "Выбери тип"
Combo1.AddItem "Начисление"
Combo1.AddItem "Оплата"
Combo1.AddItem "Субсидия"
Combo1.Text = "Резерв"
Combo1.Visible = False




'Set mconn = New ADODB.Connection
 ' mconn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.amd;Persist Security Info=True"
  'mconn.Open "data/kvartplata.amd"
    
Set Combo_RS = New ADODB.Recordset
Set Combo_RS.ActiveConnection = Mconn
 
Set Combo_rs1 = New ADODB.Recordset
Set Combo_rs1.ActiveConnection = Mconn
 
Combo_RS.CursorType = adOpenForwardOnly
Combo_RS.LockType = adLockBatchOptimistic
Combo_RS.Open "Nachisleniy"
Combo_rs1.Open "KLS_PODR"

Text1.Text = Date

' Заполняем Combo2 для начисления
'Set Combo2.DataSource = Combo_RS
Combo2.Text = "Любое начисление"
Cl = "Любое начисление"
Combo_RS.MoveFirst
Do While Not Combo_RS.EOF
Combo2.AddItem Cl
Cl = CStr(Combo_RS("Kod")) & "  " & Combo_RS("Naim")
'codN(Combo_RS("Kod")) = Combo_RS("Kod")
Combo_RS.MoveNext
Loop

' Заполняем Combo3 для адресов
'Set Combo2.DataSource = Combo_RS

Combo3.Text = ""
Cl = ""
Combo_rs1.MoveFirst

Do While Not Combo_rs1.EOF
If Trim(Cl) <> "" Then Combo3.AddItem Cl
Cl = CStr(Combo_rs1("Код")) & "  " & Combo_rs1("Naim_kls") & " дом № " & Combo_rs1("Num")
Combo_rs1.MoveNext
Loop
'Combo2.AddItem = cl




End Sub

