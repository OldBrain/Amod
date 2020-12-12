VERSION 5.00
Begin VB.Form Reports 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   6756
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7452
   ControlBox      =   0   'False
   Icon            =   "Reports.frx":0000
   LinkTopic       =   "Form7"
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   563
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   621
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command19 
      Caption         =   "Печать квитанций"
      Height          =   492
      Left            =   3600
      TabIndex        =   22
      Top             =   5280
      Width           =   3732
   End
   Begin VB.CommandButton Command18 
      Caption         =   "статистика по счетчикам"
      Height          =   492
      Left            =   3960
      TabIndex        =   20
      Top             =   3840
      Width           =   3372
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Command17"
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   1560
      Width           =   3615
   End
   Begin KvPay.xpcmdbutton xpcmdbutton1 
      Height          =   492
      Left            =   120
      TabIndex        =   18
      Top             =   5280
      Width           =   3372
      _ExtentX        =   5948
      _ExtentY        =   868
      Caption         =   "Выгрузка данных"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Планирование"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4320
      Width           =   7170
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3360
      Width           =   3375
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Должники по месяцам"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3840
      Width           =   3615
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Выход"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   116
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6240
      Width           =   7170
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Статистика"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4800
      Width           =   7170
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Отчеты в администрацию, анализ расчетов."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2880
      Width           =   3375
   End
   Begin VB.CommandButton BtnEnh2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Оборотная ведомость"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "По начислениям"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   623
      Width           =   3375
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Должники"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1080
      Width           =   3615
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "По подразделениям"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1080
      Width           =   3375
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ведомость начислений"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1920
      Width           =   3615
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ведомость оплаты/субсидий"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1530
      Width           =   3375
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Универсальный отчет"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2400
      Width           =   3615
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Свод по домам"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1980
      Width           =   3375
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ведомость поступления оплаты"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3360
      Width           =   3615
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Список лиц.счетов"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2430
      Width           =   3375
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Сверка"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2880
      Width           =   3615
   End
   Begin KvPay.xpcmdbutton cm2 
      Height          =   492
      Left            =   120
      TabIndex        =   21
      Top             =   5760
      Width           =   7212
      _ExtentX        =   12721
      _ExtentY        =   868
      Caption         =   "Счетчики-Анализ данных "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "АРМ ""Квартплата + "" Отчеты и Анализ"
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
      Left            =   240
      TabIndex        =   0
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   4170
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
      Height          =   156
      Left            =   6360
      Picture         =   "Reports.frx":030A
      Top             =   120
      Width           =   156
   End
   Begin VB.Image imgTitleLeft 
      Height          =   360
      Left            =   3120
      Picture         =   "Reports.frx":0554
      Top             =   240
      Width           =   228
   End
   Begin VB.Image imgTitleRight 
      Height          =   360
      Left            =   600
      Picture         =   "Reports.frx":0C9E
      Top             =   240
      Width           =   228
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   2520
      Picture         =   "Reports.frx":13E8
      Stretch         =   -1  'True
      ToolTipText     =   "Двойной щелчек мышы развернет форму во весь экран или вернет в исходное состояние"
      Top             =   240
      Width           =   285
   End
End
Attribute VB_Name = "Reports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Temp
Dim flgResize As Boolean
Dim OldCursorPos As PointAPI
Dim NewCursorPos As PointAPI
Dim st As ADODB.Recordset
Dim D As String
Public Svern As String
Public Прив As String
Public Благ As String
Public Подр As String

Public sq As String

Private Sub BtnEnh1_Click()

End Sub

Private Sub cm2_Click()
sc_analiz.Show
Unload Reports
End Sub

Private Sub Command10_Click()
'
Analizlgot.Titl = "Список лицевых счетов " + MainMenu.Command13.Caption
'+ D + " " + Str(Year(MainForm.DR))

Analizlgot.G = 11
sq = "SELECT KLS_PODR.NAIM_KLS AS Адрес, MainOccupant.kv_num AS Кв, MainOccupant.OLDNUM AS [N лиц сч], MainOccupant.BanKN AS [N лиц сч банк], MainOccupant.FAM AS Фамилия, MainOccupant.IM AS Имя, MainOccupant.OT AS Отчество, MainOccupant.COMSPACE AS Площадь, MainOccupant.NLODGERF AS Прописано FROM MainOccupant LEFT JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.КОД ORDER BY KLS_PODR.NAIM_KLS, MainOccupant.kv_num"
'Analizlgot.Об 2

Analizlgot.fg1.OutlineBar = flexOutlineBarComplete
Analizlgot.fg1.Subtotal flexSTSum, 0, 8, , RGB(150, 200, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 0, 9, , RGB(150, 250, 200), vbBlack, True


Analizlgot.fg1.Subtotal flexSTSum, 1, 8, , RGB(250, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 9, , RGB(250, 250, 200), vbBlack, True

Unload Me
Analizlgot.Show
End Sub

Private Sub BtnEnh14_Click()
'Оплата по датам
Analizlgot.Titl = "Список лицевых счетов " + MainMenu.Command13.Caption
'+ D + " " + Str(Year(MainForm.DR))

Analizlgot.G = 12
sq = "SELECT KLS_PODR.NAIM_KLS AS Адрес, MainOccupant.kv_num AS Кв, MainOccupant.OLDNUM AS [N лиц сч], MainOccupant.BanKN AS [N лиц сч банк], MainOccupant.FAM AS Фамилия, MainOccupant.IM AS Имя, MainOccupant.OT AS Отчество, MainOccupant.COMSPACE AS Площадь, MainOccupant.NLODGERF AS Прописано, Lgota.Numer AS [Код льг], Lgota.NAME_KLS AS Льгота FROM (MainOccupant LEFT JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.КОД) LEFT JOIN Lgota ON MainOccupant.Numer = Lgota.NomNum ORDER BY KLS_PODR.NAIM_KLS, MainOccupant.kv_num"
Analizlgot.Об 0
Unload Me
Analizlgot.Show
End Sub

Private Sub BtnEnh2_Click()


RepObor.Show 1



If Подр = "Все" And Прив = "Все" And Благ = "Все" Then FI = ""
If Подр <> "Все" And Прив = "Все" And Благ = "Все" Then FI = " WHERE (KLS_PODR.Подразделение) = '" + Подр + "'"
If Подр = "Все" And Прив <> "Все" And Благ = "Все" Then FI = " WHERE (MainOccupant.Priv) = '" + Прив + "'"
If Подр = "Все" And Прив = "Все" And Благ <> "Все" Then FI = " WHERE (KLS_PODR.благ) = '" + Благ + "'"

If Подр = "Все" And Прив <> "Все" And Благ <> "Все" Then FI = " WHERE (KLS_PODR.благ) = '" + Благ + "' and ((MainOccupant.Priv)='" + Прив + "')"
If Подр <> "Все" And Прив = "Все" And Благ <> "Все" Then FI = " WHERE (KLS_PODR.благ) = '" + Благ + "' and ((KLS_PODR.Подразделение)='" + Подр + "')"
If Подр <> "Все" And Прив <> "Все" And Благ = "Все" Then FI = " WHERE (MainOccupant.Priv) = '" + Прив + "' and ((KLS_PODR.Подразделение)='" + Подр + "')"

If Подр <> "Все" And Прив <> "Все" And Благ <> "Все" Then FI = " WHERE (((MainOccupant.Priv)='" + Прив + "') AND ((KLS_PODR.Благ)='" + Благ + "') AND ((KLS_PODR.Подразделение)='" + Подр + "'))"

'"WHERE (((MainOccupant.Priv)="ДА") AND ((KLS_PODR.Благ)="Благоустр.") AND ((KLS_PODR.Подразделение)="Подр.№2")) GROUP BY Adding.NameKat, KLS_PODR.NAIM_KLS ORDER BY Adding.NameKat"

If Svern = "Свернутая" Then



Dim rsProv As ADODB.Recordset

Set rsProv = New ADODB.Recordset
rsProv.Open ("SELECT Adding.KodKv,Adding.key FROM Adding LEFT JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer WHERE (((MainOccupant.Numer) Is Null))"), Mconn, adOpenKeyset, adLockPessimistic
If rsProv.RecordCount > 0 Then
rsProv.MoveFirst
Do While Not rsProv.EOF
Mconn.Execute ("DELETE Adding.KodKv From Adding WHERE (((Adding.key)=" + Str(rsProv("key")) + "))")
rsProv.MoveNext
Loop
End If
'If Arhiv = True Then
Analizlgot.Titl = "Оборотная ведомость за " + MainMenu.Command13.Caption
'+ " " + Str(Year(MainForm.DR))
'Else
'Analizlgot.Titl = "Оборотная ведомость за " + MonthName(Month(MainForm.DR), False) + " " + Str(Year(MainForm.DR))
'End If
Analizlgot.Vid = "ОбВд"

Analizlgot.G = 8
'sq = "SELECT Adding.NameKat as [Категория начисления], KLS_PODR.NAIM_KLS as Адрес, IIf([Adding]![Kol]=0,0,round([Adding]![SaldoN]/[Adding]![Kol],2)) as [Сальдо на начало], IIf([Adding]![Tip]=" & Chr(34) & "+" & Chr(34) & ",round([Adding]![SummaI],2),0) AS Начислено, IIf([Adding]![Tip]=" & Chr(34) & "-" & Chr(34) & ",round([Adding]![SummaI],2),0) AS Оплачено, IIf([Adding]![Tip]=" & Chr(34) & "s" & Chr(34) & ",round([Adding]![SummaI],2),0) AS Субсидии, IIf([Adding]![Kol]=0,0,round([Adding]![SaldoK]/[Adding]![Kol],2)) as [Сальдо конечное] FROM Adding LEFT JOIN (KLS_PODR RIGHT JOIN MainOccupant ON KLS_PODR.КОД = MainOccupant.Dom) ON Adding.KodKv = MainOccupant.Numer ORDER BY Adding.NameKat"
'sq = "SELECT Adding.NameKat as [Категория начисления], ' ' as ' ',IIf([Adding]![Kol]=0,0,round([Adding]![SaldoN]/[Adding]![Kol],2)) as [Сальдо на начало], IIf([Adding]![Tip]=" & Chr(34) & "+" & Chr(34) & ",round([Adding]![SummaI],2),0) AS Начислено, IIf([Adding]![Tip]=" & Chr(34) & "-" & Chr(34) & ",round([Adding]![SummaI],2),0) AS Оплачено, IIf([Adding]![Tip]=" & Chr(34) & "s" & Chr(34) & ",round([Adding]![SummaI],2),0) AS Субсидии, IIf([Adding]![Kol]=0,0,round([Adding]![SaldoK]/[Adding]![Kol],2)) as [Сальдо конечное] FROM Adding LEFT JOIN (KLS_PODR RIGHT JOIN MainOccupant ON KLS_PODR.КОД = MainOccupant.Dom) ON Adding.KodKv = MainOccupant.Numer ORDER BY Adding.NameKat"
'sq = "SELECT Adding.NameKat as [Категория начисления], ' ' as _,IIf([Adding]![Kol]=0,0,[Adding]![SaldoN]/[Adding]![Kol]) as [Сальдо на начало], IIf([Adding]![Tip]=" & Chr(34) & "+" & Chr(34) & ",[Adding]![SummaI],0) AS Начислено, IIf([Adding]![Tip]=" & Chr(34) & "-" & Chr(34) & ",[Adding]![SummaI],0) AS Оплачено, IIf([Adding]![Tip]=" & Chr(34) & "s" & Chr(34) & ",[Adding]![SummaI],0) AS Субсидии, IIf([Adding]![Kol]=0,0,[Adding]![SaldoK]/[Adding]![Kol]) as [Сальдо конечное] FROM Adding LEFT JOIN (KLS_PODR RIGHT JOIN MainOccupant ON KLS_PODR.КОД = MainOccupant.Dom) ON Adding.KodKv = MainOccupant.Numer ORDER BY Adding.NameKat"

'sq = "SELECT ' ' AS _, Adding.NameKat AS [Категория начисления], Sum(IIf([Adding]![Kol]=0,0,[Adding]![SaldoN]/[Adding]![Kol])) AS [Сальдо на начало], Sum(IIf([Adding]![Tip]='+',[Adding]![SummaI],0)) AS Начислено, Sum(IIf([Adding]![Tip]='-',[Adding]![SummaI],0)) AS Оплачено, Sum(IIf([Adding]![Tip]='s',[Adding]![SummaI],0)) AS Субсидии, Sum(IIf([Adding]![Kol]=0,0,[Adding]![SaldoK]/[Adding]![Kol])) AS [Сальдо конечное] FROM Adding LEFT JOIN (KLS_PODR RIGHT JOIN MainOccupant ON KLS_PODR.КОД = MainOccupant.Dom) ON Adding.KodKv = MainOccupant.Numer GROUP BY Adding.NameKat, ' ' ORDER BY Adding.NameKat"

sq = "SELECT ' ' AS _, Adding.NameKat AS [Категория начисления], Round(Sum(IIf([Adding]![Kol]=0,0,[Adding]![SaldoN]/[Adding]![Kol])),2) AS [Сальдо на начало], Sum(IIf([Adding]![Tip]='+',[Adding]![SummaI],0)) AS Начислено, Sum(IIf([Adding]![Tip]='-',[Adding]![SummaI],0)) AS Оплачено, Sum(IIf([Adding]![Tip]='s',[Adding]![SummaI],0)) AS Субсидии, Round(Sum(IIf([Adding]![Kol]=0,0,[Adding]![SaldoK]/[Adding]![Kol])),2) AS [Сальдо конечное] FROM Adding LEFT JOIN (KLS_PODR RIGHT JOIN MainOccupant ON KLS_PODR.КОД = MainOccupant.Dom) ON Adding.KodKv = MainOccupant.Numer GROUP BY Adding.NameKat, ' ' ORDER BY Adding.NameKat"

Analizlgot.Об 1

End If
                'ПО ДОМАМ ВСЕ

If Svern = "Все" Then
Analizlgot.G = 14

sq = "SELECT Adding.NameKat AS [Категория начисления], KLS_PODR.NAIM_KLS, Sum([NROOM]/Adding!Kol) AS [Количество комнат], Sum([NLODGERF]/Adding!Kol) AS Прописано, Sum([NLODGER]/Adding!Kol) AS Проживает, Sum([COMSPACE]/Adding!Kol) AS [Общая площадь], Sum([HABSPACE]/Adding!Kol) AS [Полезная пложадь], Sum(IIf(Adding!Kol=0,0,Adding!SaldoN/Adding!Kol)) AS [Сальдо на начало], Sum(IIf(Adding!Tip='+',Adding!SummaI,0)) AS Начислено, Sum(IIf(Adding!Tip='-',Adding!SummaI,0)) AS Оплачено, Sum(IIf(Adding!Tip='s',Adding!SummaI,0)) AS Субсидии, Sum(IIf(Adding!Kol=0,0,Adding!SaldoK/Adding!Kol)) AS [Сальдо конечное], Sum(IIf(Adding!Tip='+',Adding!SummaBl-Adding!SummaI,0)) AS [К возмещению] FROM Adding LEFT JOIN (KLS_PODR RIGHT JOIN MainOccupant ON KLS_PODR.КОД = MainOccupant.Dom) ON Adding.KodKv = MainOccupant.Numer" + FI + " GROUP BY Adding.NameKat, KLS_PODR.NAIM_KLS ORDER BY Adding.NameKat"

Analizlgot.fg1.Subtotal flexSTSum, 1, 3, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 4, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 5, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 6, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 7, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 8, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 9, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 10, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 11, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 12, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 13, , RGB(150, 250, 200), vbBlack, True

End If

If Svern = "Дома5" Then
Analizlgot.G = 11

sq = "SELECT Adding.NameKat AS [Категория начисления], KLS_PODR.NAIM_KLS AS Адрес, Sum([NLODGERF]/Adding!Kol) AS Прописано, Sum([COMSPACE]/Adding!Kol) AS [Общая площадь], Sum(IIf(Adding!Kol=0,0,Adding!SaldoN/Adding!Kol)) AS [Сальдо на начало], Sum(IIf(Adding!Tip='+',Adding!SummaI,0)) AS Начислено, Sum(IIf(Adding!Tip='-',Adding!SummaI,0)) AS Оплачено, Sum(IIf(Adding!Tip='s',Adding!SummaI,0)) AS Субсидии, Sum(IIf(Adding!Kol=0,0,Adding!SaldoK/Adding!Kol)) AS [Сальдо конечное], Sum(IIf([Adding]![Tip]='+',[Adding]![SummaBl]-[Adding]![SummaI],0)) AS [К возмещению] FROM Adding LEFT JOIN (KLS_PODR RIGHT JOIN MainOccupant ON KLS_PODR.КОД = MainOccupant.Dom) ON Adding.KodKv = MainOccupant.Numer" + FI + " GROUP BY Adding.NameKat, KLS_PODR.NAIM_KLS ORDER BY Adding.NameKat"


Analizlgot.fg1.Subtotal flexSTSum, 1, 3, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 4, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 5, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 6, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 7, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 8, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 9, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 10, , RGB(150, 250, 200), vbBlack, True

End If

If Svern = "Дома6" Then
Analizlgot.G = 9

sq = "SELECT Adding.NameKat AS [Категория начисления], KLS_PODR.NAIM_KLS AS Адрес, Sum(IIf(Adding!Kol=0,0,Adding!SaldoN/Adding!Kol)) AS [Сальдо на начало], Sum(IIf(Adding!Tip='+',Adding!SummaI,0)) AS Начислено, Sum(IIf(Adding!Tip='-',Adding!SummaI,0)) AS Оплачено, Sum(IIf(Adding!Tip='s',Adding!SummaI,0)) AS Субсидии, Sum(IIf(Adding!Kol=0,0,Adding!SaldoK/Adding!Kol)) AS [Сальдо конечное], Sum(IIf([Adding]![Tip]='+',[Adding]![SummaBl]-[Adding]![SummaI],0)) AS [К возмещению] FROM Adding LEFT JOIN (KLS_PODR RIGHT JOIN MainOccupant ON KLS_PODR.КОД = MainOccupant.Dom) ON Adding.KodKv = MainOccupant.Numer" + FI + " GROUP BY Adding.NameKat, KLS_PODR.NAIM_KLS ORDER BY Adding.NameKat"

Analizlgot.fg1.Subtotal flexSTSum, 1, 3, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 4, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 5, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 6, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 7, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 8, , RGB(150, 250, 200), vbBlack, True

End If


If Svern = "Дома4" Then
Analizlgot.G = 11

sq = "SELECT Adding.NameKat AS [Категория начисления], KLS_PODR.NAIM_KLS, Sum([NLODGERF]/Adding!Kol) AS Прописано, Sum([NLODGER]/Adding!Kol) AS Проживает, Sum(IIf(Adding!Kol=0,0,Adding!SaldoN/Adding!Kol)) AS [Сальдо на начало], Sum(IIf(Adding!Tip='+',Adding!SummaI,0)) AS Начислено, Sum(IIf(Adding!Tip='-',Adding!SummaI,0)) AS Оплачено, Sum(IIf(Adding!Tip='s',Adding!SummaI,0)) AS Субсидии, Sum(IIf(Adding!Kol=0,0,Adding!SaldoK/Adding!Kol)) AS [Сальдо конечное], Sum(IIf(Adding!Tip='+',Adding!SummaBl-Adding!SummaI,0)) AS [К возмещению] FROM Adding LEFT JOIN (KLS_PODR RIGHT JOIN MainOccupant ON KLS_PODR.КОД = MainOccupant.Dom) ON Adding.KodKv = MainOccupant.Numer" + FI + " GROUP BY Adding.NameKat, KLS_PODR.NAIM_KLS ORDER BY Adding.NameKat"

Analizlgot.fg1.Subtotal flexSTSum, 1, 3, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 4, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 5, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 6, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 7, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 8, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 9, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 10, , RGB(150, 250, 200), vbBlack, True
End If

If Svern = "Дома2" Then
Analizlgot.G = 9

sq = "SELECT Adding.NameKat AS [Категория начисления], KLS_PODR.NAIM_KLS, Sum(IIf(Adding!Kol=0,0,Adding!SaldoN/Adding!Kol)) AS [Сальдо на начало], Sum(IIf(Adding!Tip='+',Adding!SummaI,0)) AS Начислено, Sum(IIf(Adding!Tip='-',Adding!SummaI,0)) AS Оплачено, Sum(IIf(Adding!Tip='s',Adding!SummaI,0)) AS Субсидии, Sum(IIf(Adding!Kol=0,0,Adding!SaldoK/Adding!Kol)) AS [Сальдо конечное], Sum(IIf(Adding!Tip='+',Adding!SummaBl-Adding!SummaI,0)) AS [К возмещению] FROM Adding LEFT JOIN (KLS_PODR RIGHT JOIN MainOccupant ON KLS_PODR.КОД = MainOccupant.Dom) ON Adding.KodKv = MainOccupant.Numer" + FI + " GROUP BY Adding.NameKat, KLS_PODR.NAIM_KLS ORDER BY Adding.NameKat"

Analizlgot.fg1.Subtotal flexSTSum, 1, 3, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 4, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 5, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 6, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 7, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 8, , RGB(150, 250, 200), vbBlack, True
End If

If Svern = "Дома3" Then
Analizlgot.G = 10

sq = "SELECT Adding.NameKat AS [Категория начисления], KLS_PODR.NAIM_KLS AS Адрес, Sum([COMSPACE]/Adding!Kol) AS [Общая площадь], Sum([HABSPACE]/Adding!Kol) AS [Полезная пложадь], Sum(IIf(Adding!Kol=0,0,Adding!SaldoN/Adding!Kol)) AS [Сальдо на начало], Sum(IIf(Adding!Tip='+',Adding!SummaI,0)) AS Начислено, Sum(IIf(Adding!Tip='-',Adding!SummaI,0)) AS Оплачено, Sum(IIf(Adding!Tip='s',Adding!SummaI,0)) AS Субсидии, Sum(IIf(Adding!Kol=0,0,Adding!SaldoK/Adding!Kol)) AS [Сальдо конечное], Sum(IIf(Adding!Tip='+',Adding!SummaBl-Adding!SummaI,0)) AS [К возмещению] FROM Adding LEFT JOIN (KLS_PODR RIGHT JOIN MainOccupant ON KLS_PODR.КОД = MainOccupant.Dom) ON Adding.KodKv = MainOccupant.Numer" + FI + " GROUP BY Adding.NameKat, KLS_PODR.NAIM_KLS ORDER BY Adding.NameKat"

Analizlgot.fg1.Subtotal flexSTSum, 1, 3, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 4, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 5, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 6, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 7, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 8, , RGB(150, 250, 200), vbBlack, True
Analizlgot.fg1.Subtotal flexSTSum, 1, 9, , RGB(150, 250, 200), vbBlack, True
End If

Unload Me
Analizlgot.Show

End Sub

Private Sub BtnEnh3_Click()
'Cell (MakeSQ)
'AnalizLgot.G = 10
'sq = "SELECT Adding.NameKat as [Категория начисления], KLS_PODR.NAIM_KLS as Адрес, MainOccupant.Fam, IIf([Adding]![Kol]=0,0,[Adding]![SaldoN]/[Adding]![Kol]) as [Сальдо на начало], IIf([Adding]![Tip]=" & Chr(34) & "+" & Chr(34) & ",[Adding]![SummaI],0) AS Начислено, IIf([Adding]![Tip]=" & Chr(34) & "-" & Chr(34) & ",[Adding]![SummaI],0) AS Оплачено, IIf([Adding]![Tip]=" & Chr(34) & "s" & Chr(34) & ",[Adding]![SummaI],0) AS Субсидии, IIf([Adding]![Kol]=0,0,[Adding]![SaldoK]/[Adding]![Kol]) as [Сальдо конечное] FROM Adding INNER JOIN (KLS_PODR INNER JOIN MainOccupant ON KLS_PODR.КОД = MainOccupant.Dom) ON Adding.KodKv = MainOccupant.Numer ORDER BY Adding.NameKat"
'Unload Me
'AnalizLgot.Show
'AnalizLgot.Об 3
End Sub

Private Sub BtnEnh4_Click()
Analizlgot.G = 8
sq = "Rep_n"
Unload Me
Analizlgot.Show
Analizlgot.Об 3
End Sub

Private Sub BtnEnh5_Click()
Analizlgot.G = 5
Unload Me

FilRep.Show

End Sub

Private Sub Command1_Click()
RepParam1.Show
End Sub

Private Sub Command11_Click()
'Сверка
RepSverka.Show
Unload Me
End Sub

Private Sub Command12_Click()
Unload Me
MainMenu.Enabled = True
MainMenu.Show
End Sub

Private Sub Command13_Click()
RepStat.Show
End Sub

Private Sub Command14_Click()
Analizlgot.G = 10
sq = "SELECT Round([Долг]/[Начислено],0) AS Месяцы, KLS_PODR.NAIM_KLS as Адрес, MainOccupant.kv_num as КВ, MainOccupant.FAM AS [Фамилия/Количество], MainOccupant.IM AS Имя, MainOccupant.OT AS Отчество, MainOccupant.COMSPACE AS Площадь, IIf([Adding]![SaldoK]>0,[Adding]![SaldoK]/[Adding]![Kol],0) AS Долг, IIf([Adding]![Tip]='+' And [Adding]![SummaI]>0,[Adding]![SummaI],1) AS Начислено FROM KLS_PODR INNER JOIN (Adding INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer) ON KLS_PODR.КОД = MainOccupant.Dom WHERE (((IIf([Adding]![Tip]='+' And [Adding]![SummaI]>0,[Adding]![SummaI],1))<>1))"
Unload Me
Analizlgot.Show
Analizlgot.Об 2
Analizlgot.fg1.Subtotal flexSTCount, 1, 4
Analizlgot.fg1.Subtotal flexSTSum, 1, 5, , RGB(150, 250, 200), vbBlack, False


End Sub

Private Sub Command15_Click()
Analizlgot.G = 10
sq = "SELECT KLS_PODR.КОД AS Код, KLS_PODR.NAIM_KLS AS Адрес, KLS_PODR.Num, Adding.NameKat, Sum(([Adding]![SaldoN]*1000/[Adding]![Kol])/1000) AS [Сальдо нач], Sum(IIf([Adding]![Tip]='+',[SummaI],0)) AS Начислено, Sum(IIf([Adding]![Tip]='s',[SummaI],0)) AS Оплачено, Sum(IIf([Adding]![Tip]='-',[SummaI],0)) AS Субсидии, Sum(([Adding]![SaldoK]*1000/[Adding]![Kol])/1000) AS [Сальдо кон] FROM KLS_PODR INNER JOIN (Adding INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer) ON KLS_PODR.КОД = MainOccupant.Dom GROUP BY KLS_PODR.КОД, KLS_PODR.NAIM_KLS, KLS_PODR.Num, Adding.NameKat ORDER BY KLS_PODR.NAIM_KLS"
Unload Me
Analizlgot.Show
Analizlgot.Об 1
End Sub

Private Sub Command16_Click()
Me.Enabled = False
RepPlan.Show
End Sub

Private Sub Command17_Click()
'Оплата по датам
Analizlgot.Titl = "Ведомость оплаты по датам за " + MainMenu.Command13.Caption
'+ D + " " + Str(Year(MainForm.DR))

Analizlgot.G = 9
'sq = "SELECT Adding.DataR as [Дата], Adding.NameKat as [Категория], Adding.NameN as [Наименование], Sum(Adding.SummaI) AS [Сумма] From Adding Where (((Adding.Tip) = '-')) GROUP BY Adding.DataR, Adding.NameKat, Adding.NameN Having (((Sum(Adding.SummaI)) <> 0)) ORDER BY Adding.DataR"
sq = "SELECT Adding.DataR AS Дата, Adding.NameKat AS Категория, Adding.NameN AS Платеж, Sum(Adding.SummaI) AS Сумма, nachisleniy.NDS AS [% НДС], nachisleniy.Komis AS [% Комиссии], Round(Sum([SummaI])*[% НДС]/(100+[% НДС]),2) AS [Сумма НДС], Round(Sum([SummaI])*[% Комиссии]/(100),2) AS [Сумма комиссии] FROM Adding INNER JOIN nachisleniy ON Adding.KodN = nachisleniy.Kod WHERE (((Adding.Tip)='-')) GROUP BY Adding.DataR, Adding.NameKat, Adding.NameN, nachisleniy.NDS, nachisleniy.Komis Having (((Sum(Adding.SummaI)) <> 0)) ORDER BY Adding.DataR"
Analizlgot.Об 3
Unload Me
Analizlgot.Show
End Sub

Private Sub Command18_Click()
Analizlgot.G = 8
sq = "Rep_Ch"
Unload Me
Analizlgot.Show
Analizlgot.Об 1
End Sub

Private Sub Command19_Click()
rep_kvit.Show

End Sub

Private Sub Command2_Click()
Analizlgot.G = 8
sq = "Rep_n"
Unload Me
Analizlgot.Show
Analizlgot.Об 3
End Sub

Private Sub Command3_Click()
Analizlgot.G = 5
Unload Me

FilRep.Show

End Sub

Private Sub Command4_Click()
Analizlgot.G = 8
sq = "Rep_nJAK5"
Unload Me
Analizlgot.Show
Analizlgot.Об 3
End Sub

Private Sub Command5_Click()
Analizlgot.G = 20
sq = "TRANSFORM Sum(OBR_CrostabN.SummaI) AS [Sum-SummaI] SELECT OBR_CrostabN.Адрес, Sum(OBR_CrostabN.SummaI) AS [Итог по дому] From OBR_CrostabN GROUP BY OBR_CrostabN.Адрес PIVOT OBR_CrostabN.NameN"
Unload Me
Analizlgot.Show
Analizlgot.Об 1

End Sub

Private Sub Command6_Click()
Analizlgot.G = 20
sq = "TRANSFORM Sum(OBR_CrostabU.SummaI) AS [Sum-SummaI]SELECT OBR_CrostabU.Адрес, Sum(OBR_CrostabU.SummaI) AS [Итого по дому] From OBR_CrostabU GROUP BY OBR_CrostabU.Адрес PIVOT OBR_CrostabU.NameN"
Unload Me
Analizlgot.Show
Analizlgot.Об 1
End Sub

Private Sub Command7_Click()
'Exit Sub
Form6.Show
End Sub

Private Sub Command8_Click()
Form8.Show
Unload Me
End Sub

Private Sub Command9_Click()
'Оплата по датам
Analizlgot.Titl = "Ведомость оплаты по датам за " + MainMenu.Command13.Caption
'+ D + " " + Str(Year(MainForm.DR))

Analizlgot.G = 9
'sq = "SELECT Adding.DataR as [Дата], Adding.NameKat as [Категория], Adding.NameN as [Наименование], Sum(Adding.SummaI) AS [Сумма] From Adding Where (((Adding.Tip) = '-')) GROUP BY Adding.DataR, Adding.NameKat, Adding.NameN Having (((Sum(Adding.SummaI)) <> 0)) ORDER BY Adding.DataR"
sq = "SELECT Adding.DataR AS Дата, Adding.NameKat AS Категория, Adding.NameN AS Платеж, Sum(Adding.SummaI) AS Сумма, nachisleniy.NDS AS [% НДС], nachisleniy.Komis AS [% Комиссии], Round(Sum([SummaI])*[% НДС]/(100+[% НДС]),2) AS [Сумма НДС], Round(Sum([SummaI])*[% Комиссии]/(100),2) AS [Сумма комиссии] FROM Adding INNER JOIN nachisleniy ON Adding.KodN = nachisleniy.Kod WHERE (((Adding.Tip)='-')) GROUP BY Adding.DataR, Adding.NameKat, Adding.NameN, nachisleniy.NDS, nachisleniy.Komis Having (((Sum(Adding.SummaI)) <> 0)) ORDER BY Adding.DataR"
Analizlgot.Об 3
Unload Me
Analizlgot.Show
End Sub

Private Sub Form_Load()
sq = ""
MakeWindow Me, True

Set st = New ADODB.Recordset
st.Open ("Settings"), Mconn
D = MonthName(Month(st("TekData")), False)
st.Close

End Sub


Sub MakeSQ(sq)
Dim Tbl As String

Tbl = "J_ALL"
sq = "SELECT "

' ************ 1
If Form6.Check1.Value Then
            sq = sq + Tbl + ".[Наименование_льготы], "
        Else
            sq = sq
        End If
' ************ 2
If Form6.Check2.Value Then
            sq = sq + Tbl + ".NAIM_KLS, "
        Else
            sq = sq
        End If
        
' ************ 3
If Form6.Check3.Value Then
            sq = sq + Tbl + ".[ФИО], "
        Else
            sq = sq
        End If

' ************ 4
If Form6.Check4.Value Then
            sq = sq + Tbl + ".[ЖилаяПЛ], "
        Else
            sq = sq
        End If

' ************ 5
If Form6.Check5.Value Then
            sq = sq + Tbl + ".[Площадь], "
        Else
            sq = sq
       End If
       
  ' ************ 6
If Form6.Check6.Value Then
            sq = sq + Tbl + ".[Проживает], "
        Else
            sq = sq
       End If
       
' ************ 7
If Form6.Check7.Value Then
            sq = sq + Tbl + ".[Прописано], "
        Else
            sq = sq
       End If
       
 ' ************ 8
If Form6.Check8.Value Then
            sq = sq + Tbl + ".[Начисление], "
        Else
            sq = sq
       End If
       
 ' ************ 9
If Form6.Check9.Value Then
            sq = sq + Tbl + ".[Счет_затрат], "
        Else
            sq = sq
       End If
       
  ' ************ 18
If Form6.Check18.Value Then
            sq = sq + Tbl + ".[Сумма], "
        Else
            sq = sq
       End If
       
       ' ************ 10
If Form6.Check10.Value Then
            sq = sq + Tbl + ".[Кухня], "
        Else
            sq = sq
       End If
       
 ' ************ 11
If Form6.Check11.Value Then
            sq = sq + Tbl + ".[Ванная], "
        Else
            sq = sq
       End If
       
       
 ' ************ 12
If Form6.Check12.Value Then
            sq = sq + Tbl + ".[Коридор], "
        Else
            sq = sq
       End If
       
 ' ************ 13
If Form6.Check13.Value Then
            sq = sq + Tbl + ".[Туалет], "
        Else
            sq = sq
       End If
       
       
' ************ 14
If Form6.Check14.Value Then
            sq = sq + Tbl + ".[Балкон], "
        Else
            sq = sq
       End If
       
       
 ' ************ 15
If Form6.Check15.Value Then
            sq = sq + Tbl + ".[Этаж], "
        Else
            sq = sq
       End If
       
' ************ 16
If Form6.Check16.Value Then
            sq = sq + Tbl + ".[ДатаПрописки], "
        Else
            sq = sq
       End If
       
' ************ 17
If Form6.Check17.Value Then
            sq = sq + Tbl + ".[Прописка], "
        Else
            sq = sq
       End If
   
'*********************************************
        sq = Left(sq, Len(sq) - 2) + " "
sq = sq + "FROM " + Tbl

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Pod
End Sub

Private Sub imgTitleHelp_Click()
Dim AboutBox As New AboutBox
With AboutBox
    .Title = " Расчет и анализ коммунальных платежей населения"
    .Version = "Версия: " + Str(App.Major) + "." + Str(App.Minor) + "." + Str(App.Revision)
    .Company = "Квартплата +  (C) Copyright, 2005, Астрахань"
    .Copyright = " Бугоров Андрей Владимирович"
    .Description = "Комплексная автоматизация бухучета"
    .License = "Связь с автором E-Mail:bestonline@list.ru телефоны: +79881733600"
    .hWndOwner = Me.hwnd
    'Set .Icon = Me.Icon
    .AboutBox
End With

End Sub

Private Sub xpcmdbutton1_Click()
'If Trim(InputBox("Введите код")) <> Trim("1967") Then Exit Sub

The_end.Show
Unload Reports
End Sub

Private Sub xpcmdbutton2_Click()

End Sub
