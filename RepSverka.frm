VERSION 5.00
Begin VB.Form RepSverka 
   BorderStyle     =   0  'None
   ClientHeight    =   4785
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5550
   ControlBox      =   0   'False
   LinkTopic       =   "Form8"
   ScaleHeight     =   319
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3000
      TabIndex        =   7
      Text            =   "Все"
      Top             =   2880
      Width           =   2295
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   6
      Text            =   "Все"
      Top             =   2880
      Width           =   2175
   End
   Begin VB.CommandButton BtnEnh1 
      BackColor       =   &H00BDC6BB&
      Caption         =   "Отмена"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4080
      Width           =   5295
   End
   Begin VB.CommandButton BtnEnh2 
      BackColor       =   &H00BDC6BB&
      Caption         =   "По лиц счетам"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3720
      Width           =   2655
   End
   Begin VB.CommandButton BtnEnh3 
      BackColor       =   &H00BDC6BB&
      Caption         =   "По видам расчета"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3720
      Width           =   2655
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   120
      TabIndex        =   2
      Text            =   "Combo2"
      Top             =   1920
      Width           =   5295
   End
   Begin VB.ComboBox Combo1 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   1080
      Width           =   5295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Приватизированные?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   9
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Благоустроеные?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Выбор параметров отчета"
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
      Left            =   360
      TabIndex        =   0
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   4050
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
      Left            =   5280
      Picture         =   "RepSverka.frx":0000
      Top             =   720
      Width           =   195
   End
   Begin VB.Image imgTitleLeft 
      Height          =   450
      Left            =   4800
      Picture         =   "RepSverka.frx":024A
      Top             =   120
      Width           =   285
   End
   Begin VB.Image imgTitleRight 
      Height          =   450
      Left            =   5160
      Picture         =   "RepSverka.frx":0994
      Top             =   120
      Width           =   285
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   4800
      Picture         =   "RepSverka.frx":10DE
      Stretch         =   -1  'True
      ToolTipText     =   "Двойной щелчек мышы развернет форму во весь экран или вернет в исходное состояние"
      Top             =   480
      Width           =   285
   End
End
Attribute VB_Name = "RepSverka"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Addrconn As ADODB.Recordset
'Dim mconn As ADODB.Connection

Private Sub BtnEnh1_Click()
MainMenu.Enabled = True
Unload Me
End Sub

Private Sub BtnEnh2_1_Click()

End Sub

Private Sub BtnEnh2_Click()
' По лиц.счетам
Dim sq As String
Dim Sort As String
Dim fil As Integer
fil = Val(Replace(Combo1.Text, " ", "_", 1))
filA = Val(Replace(Combo2.Text, " ", "_", 1))


If Combo1.Text = "Все дома" And Combo2.Text = "Все начисления" Then StrU = ""
If Combo1.Text <> "Все дома" And Combo2.Text = "Все начисления" Then StrU = "WHERE (((KLS_PODR.КОД)=" + Str(fil) + "))"
If Combo1.Text <> "Все дома" And Combo2.Text <> "Все начисления" Then StrU = "WHERE (((KLS_PODR.КОД)=" + Str(fil) + " And (Adding.KodN=" + Str(filA) + ")))"
If Combo1.Text = "Все дома" And Combo2.Text <> "Все начисления" Then StrU = "WHERE (((Adding.KodN)=" + Str(filA) + "))"


If Combo1.Text = "Все дома" And Combo2.Text = "Все начисления" And Combo3.Text <> "Все" And Combo4.Text = "Все" Then StrU = "WHERE (((KLS_PODR.Благ)='" + Combo3.Text + "'))"
If Combo1.Text = "Все дома" And Combo2.Text = "Все начисления" And Combo3.Text = "Все" And Combo4.Text <> "Все" Then StrU = "WHERE (((MainOccupant.priv)='" + Combo4.Text + "'))"
If Combo1.Text = "Все дома" And Combo2.Text = "Все начисления" And Combo3.Text <> "Все" And Combo4.Text <> "Все" Then StrU = "WHERE (((KLS_PODR.Благ)='" + Combo3.Text + "') AND ((MainOccupant.Priv)='" + Combo4.Text + "'))"

If Combo1.Text <> "Все дома" And Combo3.Text <> "Все" Then
MsgBox "Неверно заданы параметры фильтра." + vbNewLine + " Если Вы хотите собрать oтчет по неблагоустроеным домам, то надо указать <Все дома>"
Exit Sub
End If


If Combo1.Text <> "Все дома" And Combo2.Text <> "Все начисления" And Combo3.Text = "Все" And Combo4.Text <> "Все" Then StrU = "WHERE (((Adding.KodN)=" + Str(filA) + ") AND ((MainOccupant.Priv)='" + Combo4.Text + "') AND ((KLS_PODR.КОД)=" + Str(fil) + "))"
If Combo1.Text <> "Все дома" And Combo2.Text = "Все начисления" And Combo3.Text = "Все" And Combo4.Text <> "Все" Then StrU = "WHERE (((MainOccupant.Priv)='" + Combo4.Text + "') AND ((KLS_PODR.КОД)=" + Str(fil) + "))"
If Combo1.Text = "Все дома" And Combo2.Text <> "Все начисления" And Combo3.Text = "Все" And Combo4.Text <> "Все" Then StrU = "WHERE (((MainOccupant.Priv)='" + Combo4.Text + "') AND ((Adding.KodN)=" + Str(filA) + "))"



Analizlgot.Titl = "Оборотная ведомость за " + MonthName(Month(MainForm.DR), False) + " " + Str(Year(MainForm.DR)) + " г., по адресу:" + Combo1.Text
Analizlgot.G = 9
'sq = "SELECT KLS_PODR.NAIM_KLS as Адрес,  MainOccupant.KV_NUM as Кв,MainOccupant.FAM as Фамилия, MainOccupant.IM as Имя, MainOccupant.OT as Отчество,Adding.KodN as Код,Adding.NameN as Начисление,   Adding.SummaI FROM Adding INNER JOIN (MainOccupant INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.КОД) ON Adding.KodKv = MainOccupant.Numer " + StrU
sq = "SELECT KLS_PODR.NAIM_KLS as Адрес,  MainOccupant.KV_NUM as Кв,MainOccupant.FAM as Фамилия, MainOccupant.IM as Имя, MainOccupant.OT as Отчество,Adding.KodN as Код,Adding.NameN as Начисление,   IIf([Adding]![Tip]='+',[SummaI],0) AS Начислено, IIf([Adding]![Tip]='-',[SummaI],0) AS Оплачено, IIf([Adding]![Tip]='s',[SummaI],0) AS Субсидии FROM Adding INNER JOIN (MainOccupant INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.КОД) ON Adding.KodKv = MainOccupant.Numer " + StrU
Analizlgot.G = 11
Analizlgot.StrSQL = sq
Analizlgot.Show
Analizlgot.FG1.AutoResize = True
Unload Me
Analizlgot.FG1.Subtotal flexSTSum, 1, 8, , RGB(150, 250, 200), vbBlack, True, "И ТОГО ПО ДОМУ"
Analizlgot.FG1.Subtotal flexSTSum, 1, 9, , RGB(150, 250, 200), vbBlack, True, "И ТОГО ПО ДОМУ"
Analizlgot.FG1.Subtotal flexSTSum, 1, 10, , RGB(150, 250, 200), vbBlack, True, "И ТОГО ПО ДОМУ"
End Sub

Private Sub BtnEnh3_Click()
'По видам расчета

Dim sq As String
Dim Sort As String
Dim StrU As String
Dim fil As Integer
Dim filA As Integer

fil = Val(Replace(Combo1.Text, " ", "_", 1))
filA = Val(Replace(Combo2.Text, " ", "_", 1))

If Combo1.Text = "Все дома" And Combo2.Text = "Все начисления" And Combo3.Text = "Все" And Combo4.Text = "Все" Then StrU = ""
If Combo1.Text <> "Все дома" And Combo2.Text = "Все начисления" And Combo3.Text = "Все" And Combo4.Text = "Все" Then StrU = "WHERE (((KLS_PODR.КОД)=" + Str(fil) + "))"
If Combo1.Text <> "Все дома" And Combo2.Text <> "Все начисления" And Combo3.Text = "Все" And Combo4.Text = "Все" Then StrU = "WHERE (((KLS_PODR.КОД)=" + Str(fil) + " And (Adding.KodN=" + Str(filA) + ")))"
If Combo1.Text = "Все дома" And Combo2.Text <> "Все начисления" And Combo3.Text = "Все" And Combo4.Text = "Все" Then StrU = "WHERE (((Adding.KodN)=" + Str(filA) + "))"

If Combo1.Text = "Все дома" And Combo2.Text = "Все начисления" And Combo3.Text <> "Все" And Combo4.Text = "Все" Then StrU = "WHERE (((KLS_PODR.Благ)='" + Combo3.Text + "'))"
If Combo1.Text = "Все дома" And Combo2.Text = "Все начисления" And Combo3.Text = "Все" And Combo4.Text <> "Все" Then StrU = "WHERE (((MainOccupant.priv)='" + Combo4.Text + "'))"
If Combo1.Text = "Все дома" And Combo2.Text = "Все начисления" And Combo3.Text <> "Все" And Combo4.Text <> "Все" Then StrU = "WHERE (((KLS_PODR.Благ)='" + Combo3.Text + "') AND ((MainOccupant.Priv)='" + Combo4.Text + "'))"

If Combo1.Text <> "Все дома" And Combo3.Text <> "Все" Then
MsgBox "Неверно заданы параметры фильтра." + vbNewLine + " Если Вы хотите собрать oтчет по неблагоустроеным домам, то надо указать <Все дома>"
Exit Sub
End If


If Combo1.Text <> "Все дома" And Combo2.Text <> "Все начисления" And Combo3.Text = "Все" And Combo4.Text <> "Все" Then StrU = "WHERE (((Adding.KodN)=" + Str(filA) + ") AND ((MainOccupant.Priv)='" + Combo4.Text + "') AND ((KLS_PODR.КОД)=" + Str(fil) + "))"
If Combo1.Text <> "Все дома" And Combo2.Text = "Все начисления" And Combo3.Text = "Все" And Combo4.Text <> "Все" Then StrU = "WHERE (((MainOccupant.Priv)='" + Combo4.Text + "') AND ((KLS_PODR.КОД)=" + Str(fil) + "))"
If Combo1.Text = "Все дома" And Combo2.Text <> "Все начисления" And Combo3.Text = "Все" And Combo4.Text <> "Все" Then StrU = "WHERE (((MainOccupant.Priv)='" + Combo4.Text + "') AND ((Adding.KodN)=" + Str(filA) + "))"


Analizlgot.Titl = "Оборотная ведомость за " + MonthName(Month(MainForm.DR), False) + " " + Str(Year(MainForm.DR)) + " г., по адресу:" + Combo1.Text
Analizlgot.G = 9
sq = "SELECT KLS_PODR.NAIM_KLS as Адрес, Adding.KodN as Код, Adding.NameN as Начисление, MainOccupant.KV_NUM as Кв, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, Adding.SummaI FROM Adding INNER JOIN (MainOccupant INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.КОД) ON Adding.KodKv = MainOccupant.Numer " + StrU
Analizlgot.G = 9
Analizlgot.StrSQL = sq
Analizlgot.Show
Analizlgot.FG1.AutoResize = True
Unload Me
Analizlgot.FG1.Subtotal flexSTSum, 1, 8, , RGB(150, 250, 200), vbBlack, True, "И ТОГО ПО ДОМУ"



End Sub


Private Sub Form_Load()
MakeWindow Me, True

  'Set mconn = New ADODB.Connection
  'mconn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Persist Security Info=True"
  'mconn.Open "data/Kvartplata.mdb"






Set Addrconn = New ADODB.Recordset
Set Addrconn.ActiveConnection = Mconn
Addrconn.CursorType = adOpenStatic
Addrconn.LockType = adLockBatchOptimistic


Set Nconn = New ADODB.Recordset
Set Nconn.ActiveConnection = Mconn
Nconn.CursorType = adOpenStatic
Nconn.LockType = adLockBatchOptimistic



'AddrConn.Open ("KLS_PODR")
Addrconn.Open ("SELECT KLS_PODR.КОД, KLS_PODR.NAIM_KLS, KLS_PODR.Num, KLS_PODR.Tip, KLS_PODR.Tip_Naim, KLS_PODR.Подразделение, KLS_PODR.Благ From KLS_PODR ORDER BY KLS_PODR.NAIM_KLS")

Combo1.Text = "Все дома"

'Для комбобокса адресов
Addrconn.MoveFirst
Combo1.AddItem "Все дома"
Do While Not Addrconn.EOF
If Addrconn("КОД") <> -1 Then
Combo1.AddItem Trim(Str(Addrconn("КОД"))) + " " + Addrconn("NAIM_KLS") + " дом № " + Addrconn("Num")
End If
Addrconn.MoveNext
Loop

'Для комбобокса начислений
Nconn.Open ("SELECT nachisleniy.Kod, nachisleniy.Naim From Nachisleniy ORDER BY nachisleniy.Kod DESC")
Combo2.Text = "Все начисления"
Nconn.MoveFirst
Combo2.AddItem "Все начисления"
Do While Not Nconn.EOF
Combo2.AddItem Trim(Str(Nconn("kod"))) + " " + Nconn("NAIM")
Nconn.MoveNext
Loop
Nconn.Close
Addrconn.Close

Combo3.AddItem "Все"
Combo3.AddItem "Благоустр."
Combo3.AddItem "Неблагоустр."


Combo4.AddItem "Все"
Combo4.AddItem "Да"
Combo4.AddItem "Нет"

End Sub


