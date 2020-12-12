VERSION 5.00
Begin VB.Form Dni 
   Caption         =   "Form4"
   ClientHeight    =   2832
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   5172
   LinkTopic       =   "Form4"
   ScaleHeight     =   2832
   ScaleWidth      =   5172
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   3720
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   2040
      Width           =   1332
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Отмена"
      Height          =   372
      Left            =   3720
      TabIndex        =   4
      Top             =   2400
      Width           =   1452
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2280
      TabIndex        =   3
      Top             =   2400
      Width           =   1332
   End
   Begin VB.ComboBox Combo1 
      Height          =   288
      Left            =   0
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   2400
      Width           =   2172
   End
   Begin VB.Label Label3 
      Caption         =   "кол-во дней"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   2640
      TabIndex        =   6
      Top             =   2040
      Width           =   1092
   End
   Begin VB.Label Label2 
      Caption         =   "Выбери категорию расчета"
      Height          =   372
      Left            =   0
      TabIndex        =   2
      Top             =   2040
      Width           =   2172
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1920
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4980
   End
End
Attribute VB_Name = "Dni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Addrconn As ADODB.Recordset
Dim D As String



Private Sub Command1_Click()

Kt = Val(Me.Combo1.Text)
Dom = Filter.FG.TextMatrix(Filter.FG.Row, 10)



Mconn.Execute ("UPDATE Adding INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer SET Adding.DnF = " + Text1.Text + " WHERE (((Adding.KodKat)=" + Str(Kt) + ") AND ((MainOccupant.Dom)=" + Dom + "))")

MsgBox ("Количество расчетных дней по адресу " + Filter.FG.TextMatrix(Filter.FG.Row, 5) + " успешно проставлены! Пересчитайте лицевые счета.")
Filter.FG.Enabled = True
Unload Me
End Sub

Private Sub Command2_Click()

Filter.Enabled = True
Unload Me
End Sub



Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()
'MakeWindow Me, False
Me.KeyPreview = True

Filter.Enabled = False


Me.Caption = "Адрес " + Filter.FG.TextMatrix(Filter.FG.Row, 5)
Me.Label1.Caption = "Установить фактическое количество дней расчета для всех л/сч. по адресу " + Filter.FG.TextMatrix(Filter.FG.Row, 5) + " ?"
' open connection
'  Set mconn = New ADODB.Connection
 ' mconn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Persist Security Info=True"
 ' mconn.Open "data/Kvartplata.mdb"
  
  
  Set Addrconn = New ADODB.Recordset
Set Addrconn.ActiveConnection = Mconn
Addrconn.CursorType = adOpenStatic
Addrconn.LockType = adLockBatchOptimistic


'Addrconn.Open ("SELECT KLS_PODR.КОД, KLS_PODR.NAIM_KLS, KLS_PODR.Num, KLS_PODR.Tip, KLS_PODR.Tip_Naim FROM KLS_PODR ORDER BY KLS_PODR.NAIM_KLS")

Addrconn.Open ("SELECT Kategor.Код, Kategor.Name_Kategor FROM Kategor")

Addrconn.MoveFirst

Combo1.Text = Str(Addrconn("Код")) + "|" + Addrconn("Name_Kategor")

Do While Not Addrconn.EOF

Combo1.AddItem Str(Addrconn("Код")) + "|" + Addrconn("Name_Kategor")
Addrconn.MoveNext
Loop

Me.Text1.Text = MainForm.DnP
'SendKeys "{F4}"
End Sub


Private Sub Form_Unload(Cancel As Integer)
Filter.Enabled = True
Filter.FG.SetFocus
Filter.Command15.Visible = True
Filter.Command17.Visible = True
End Sub

