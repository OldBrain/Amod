VERSION 5.00
Begin VB.Form RepPlan 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2985
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4440
   ControlBox      =   0   'False
   Icon            =   "RepPlan.frx":0000
   LinkTopic       =   "Статистические отчеты"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   199
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   296
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Свернутая с НДС и комиссией "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   0
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1950
      Width           =   4455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Выход"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2400
      Width           =   4455
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Свернутая"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   0
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1500
      Width           =   4455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Плановые поступления (Развернутая по домам)"
      Height          =   450
      Left            =   0
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1050
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Плановые поступления (Развернутая)"
      Height          =   450
      Left            =   0
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Запрос Количествопроп"
      Top             =   600
      Width           =   4455
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "АРМ ""Квартплата + "" Планирование"""
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
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   4170
   End
   Begin VB.Image imgTitleLeft 
      Height          =   450
      Left            =   360
      Picture         =   "RepPlan.frx":030A
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   600
      Picture         =   "RepPlan.frx":0A54
      Stretch         =   -1  'True
      ToolTipText     =   "Двойной щелчек мышы развернет форму во весь экран или вернет в исходное состояние"
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleRight 
      Height          =   450
      Left            =   840
      Picture         =   "RepPlan.frx":119E
      Top             =   0
      Width           =   285
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
      Height          =   240
      Left            =   0
      Picture         =   "RepPlan.frx":18E8
      ToolTipText     =   "Закрыть"
      Top             =   0
      Width           =   240
   End
End
Attribute VB_Name = "RepPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Reports
Unload RepStat
Analizlgot.G = 10
Analizlgot.Titl = "Анализ планируемых поступлений средств за " + D + " " + Str(Year(MainForm.DR)) + " (Развернутая)"
Reports.sq = "SELECT IIf([Adding]![Tip]='+','Начисления',IIf([Adding]![Tip]='-','Оплата','Субсидии')) AS Вид, Adding.NameKat as Категория, IIf([Adding]![Tip]='+',[Propis],0) AS Прописано, IIf([Adding]![Tip]='+',[Tarif],0) AS Тариф, Count(Adding.KodKv) AS [Количество л/сч], [Прописано]*[Количество л/сч] AS [ВСЕГО Прописано], Sum(IIf([Tip]='+',[Adding]![SummaBl],0)) AS [Без льгот], Sum(Adding.SummaI) AS [С учетом льгот], Sum(IIf([Tip]='+',[Adding]![SummaBl]-[Adding]![SummaI],0)) AS [К возмещению] From Adding GROUP BY IIf([Adding]![Tip]='+','Начисления',IIf([Adding]![Tip]='-','Оплата','Субсидии')), Adding.NameKat, IIf([Adding]![Tip]='+',[Propis],0), IIf([Adding]![Tip]='+',[Tarif],0)"
'Analizlgot.Об 2

Analizlgot.FG1.OutlineBar = flexOutlineBarComplete


Analizlgot.FG1.Subtotal flexSTSum, 0, 6, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 7, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 9, , RGB(150, 250, 200), vbBlack, True

Analizlgot.FG1.Subtotal flexSTSum, 1, 8, , RGB(150, 250, 250), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 1, 5, , RGB(150, 250, 250), vbBlack, True
Unload Me
Analizlgot.Show
End Sub


Private Sub Command2_Click()
Unload Reports
Unload RepStat
Analizlgot.G = 12
Analizlgot.Titl = "Анализ планируемых поступлений средств за " + D + " " + Str(Year(MainForm.DR)) + " (Развернутая по домам)"
Reports.sq = "SELECT Adding.NameKat AS Категория, IIf([Adding]![Tip]='+','Начисления',IIf([Adding]![Tip]='-','Оплата','Субсидии')) AS Вид, KLS_PODR.NAIM_KLS as Адрес, IIf([Adding]![Tip]='+',[Propis],0) AS Прописано, IIf([Adding]![Tip]='+',[Tarif],0) AS Тариф, Count(Adding.KodKv) AS [Количество л/сч], [Прописано]*[Количество л/сч] AS [ВСЕГО Прописано], Sum(IIf([Adding]![Tip]='+',[Adding]![SummaBl],0)) AS [Без льгот], Sum(Adding.SummaI) AS [С учетом льгот], Sum(IIf([Adding]![Tip]='+',[Adding]![SummaBl]-[Adding]![SummaI],0)) AS [К возмещению] FROM KLS_PODR INNER JOIN (Adding INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer) ON KLS_PODR.КОД = MainOccupant.Dom GROUP BY Adding.NameKat, IIf([Adding]![Tip]='+','Начисления',IIf([Adding]![Tip]='-','Оплата','Субсидии')), KLS_PODR.NAIM_KLS, IIf([Adding]![Tip]='+',[Propis],0), IIf([Adding]![Tip]='+',[Tarif],0)"
'Analizlgot.Об 4

Analizlgot.FG1.OutlineBar = flexOutlineBarComplete

Analizlgot.FG1.Subtotal flexSTSum, 0, 8, , RGB(150, 250, 250), vbBlack, True

Analizlgot.FG1.Subtotal flexSTSum, 0, 10, , RGB(150, 250, 250), vbBlack, True

'Analizlgot.FG1.Subtotal flexSTSum, 1, 9, , RGB(150, 250, 200), vbBlack, True

Analizlgot.FG1.Subtotal flexSTSum, 2, 9, , RGB(240, 240, 230), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 3, 9, , RGB(250, 250, 200), vbBlack, True


Analizlgot.FG1.Subtotal flexSTSum, 1, 7, , RGB(250, 250, 250), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 3, 7, , RGB(250, 250, 200), vbBlack, True

Analizlgot.FG1.Subtotal flexSTSum, 1, 6, , RGB(250, 250, 250), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 3, 6, , RGB(250, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 3, 10, , RGB(250, 250, 200), vbBlack, True
Unload Me
Analizlgot.Show
End Sub

Private Sub Command3_Click()
Unload Me
Reports.Enabled = True
End Sub

Private Sub Command4_Click()

Unload Reports
Unload RepStat
Analizlgot.G = 11
Analizlgot.Titl = "Анализ планируемых поступлений средств за " + D + " " + Str(Year(MainForm.DR)) + " (Свернутая с налогами и комиссией)"
Reports.sq = "SELECT IIf([Adding]![Tip]='+','Начисления',IIf([Adding]![Tip]='-','Оплата','Субсидии')) AS Вид, Adding.NameKat as Категория, IIf([Adding]![Tip]='+',[Tarif],0) AS Тариф, Count(Adding.KodKv) AS [Количество л/сч], Sum(Adding.Propis) AS [Прописано], Sum(IIf([Adding]![Tip]='+',[Adding]![SummaBl],0)) AS [Без льгот], Sum(Adding.SummaI) AS [Sum-SummaI], Sum(IIf([Adding]![Tip]='+',[Adding]![SummaBl]-[Adding]![SummaI],0)) AS [К возмещению], Sum([Adding]![SummaI]*[nachisleniy]![NDS]/(100+[nachisleniy]![NDS])) AS [В т ч НДС], Sum([Adding]![SummaI]*[nachisleniy]![Komis]/100) AS Комиссия FROM Adding INNER JOIN nachisleniy ON Adding.KodN = nachisleniy.Kod GROUP BY IIf([Adding]![Tip]='+','Начисления',IIf([Adding]![Tip]='-','Оплата','Субсидии')), Adding.NameKat, IIf([Adding]![Tip]='+',[Tarif],0)"
'Analizlgot.Об 1

Analizlgot.FG1.OutlineBar = flexOutlineBarComplete

Analizlgot.FG1.Subtotal flexSTSum, 0, 8, , RGB(150, 200, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 1, 8, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 2, 7, , RGB(150, 250, 250), vbBlack, True

Analizlgot.FG1.Subtotal flexSTSum, 2, 5, , RGB(150, 250, 250), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 2, 4, , RGB(150, 250, 250), vbBlack, True

Analizlgot.FG1.Subtotal flexSTSum, 0, 6, , RGB(150, 200, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 1, 6, , RGB(150, 250, 200), vbBlack, True

Analizlgot.FG1.Subtotal flexSTSum, 1, 9, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 1, 10, , RGB(150, 250, 200), vbBlack, True
Unload Me
Analizlgot.Show

End Sub

Private Sub Command7_Click()
Unload Reports
Unload RepStat
Analizlgot.G = 9
Analizlgot.Titl = "Анализ планируемых поступлений средств за " + D + " " + Str(Year(MainForm.DR)) + " (Свернутая)"
Reports.sq = "SELECT Adding.NameKat AS Категория, IIf([Adding]![Tip]='+','Начисления',IIf([Adding]![Tip]='-','Оплата','Субсидии')) AS Вид, IIf([Adding]![Tip]='+',[Tarif],0) AS Тариф, Count(Adding.KodKv) AS [Количество л/сч], Sum(Adding.Propis) AS Прописано, Sum(IIf([Adding]![Tip]='+',[Adding]![SummaBl],0)) AS [Без льгот], Sum(Adding.SummaI) AS [С учетом льгот], Sum(IIf([Adding]![Tip]='+',[Adding]![SummaBl]-[Adding]![SummaI],0)) AS [К возмещению] From Adding GROUP BY Adding.NameKat, IIf([Adding]![Tip]='+','Начисления',IIf([Adding]![Tip]='-','Оплата','Субсидии')), IIf([Adding]![Tip]='+',[Tarif],0)"
'Analizlgot.Об 1

Analizlgot.FG1.OutlineBar = flexOutlineBarComplete



'Analizlgot.FG1.Subtotal flexSTNone, 1, 3, , RGB(150, 250, 200), vbBlack, True
'Analizlgot.FG1.Subtotal flexSTNone, 0, 4, , RGB(150, 250, 200), vbBlack, True

'Analizlgot.FG1.Subtotal
Analizlgot.FG1.Subtotal flexSTSum, 0, 8, , RGB(150, 200, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 1, 8, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 2, 7, , RGB(150, 250, 250), vbBlack, True

Analizlgot.FG1.Subtotal flexSTSum, 2, 5, , RGB(150, 250, 250), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 2, 4, , RGB(150, 250, 250), vbBlack, True

Analizlgot.FG1.Subtotal flexSTSum, 0, 6, , RGB(150, 200, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 1, 6, , RGB(150, 250, 200), vbBlack, True
Unload Me
Analizlgot.Show
End Sub

Private Sub Form_Load()
Reports.Enabled = False
MakeWindow Me, True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Reports.Enabled = True

End Sub
