VERSION 5.00
Begin VB.Form RepParam1 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7740
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7155
   ControlBox      =   0   'False
   Icon            =   "RepParam1.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   516
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   477
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command8 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Отчет в администрацию развернутый"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4920
      Width           =   6975
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Список л/сч. с возможными ошибками расчета льгот и лиготируемых площадей, 'Вывоз мусора'. "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Просто обратить внимание на эти счета"
      Top             =   6600
      Width           =   6975
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Список л/сч. с возможными ошибками расчета льгот и лиготируемых площадей, 'Квартплата'. "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Просто обратить внимание на эти счета"
      Top             =   5760
      Width           =   6975
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Отчет в администрацию свернутый"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4080
      Width           =   6975
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "По видам льгот развернутая лиготируемая площадь"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3240
      Width           =   3495
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "По видам льгот свернутая лиготируемая площадь"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3240
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "По лиц.счетам"
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
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2760
      Width           =   3480
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
      Left            =   1320
      Sorted          =   -1  'True
      TabIndex        =   6
      Text            =   "Все"
      Top             =   1920
      Width           =   5775
   End
   Begin VB.ComboBox Combo2 
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
      Left            =   2520
      TabIndex        =   4
      Text            =   "Все"
      Top             =   1320
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Развернутый"
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
      TabIndex        =   1
      Top             =   2760
      Width           =   3495
   End
   Begin VB.ComboBox Combo1 
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
      Left            =   2040
      TabIndex        =   0
      Text            =   "Все"
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "В отчет не включаются ""вручную"" исправленные суммы "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2400
      Width           =   6975
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "АРМ ""Квартплата + "" Отчеты и Анализ"
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
      Left            =   1440
      TabIndex        =   9
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
      Height          =   240
      Left            =   0
      Picture         =   "RepParam1.frx":030A
      ToolTipText     =   "Закрыть"
      Top             =   0
      Width           =   240
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   960
      Picture         =   "RepParam1.frx":084C
      Stretch         =   -1  'True
      ToolTipText     =   "Двойной щелчек мышы развернет форму во весь экран или вернет в исходное состояние"
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleRight 
      Height          =   450
      Left            =   360
      Picture         =   "RepParam1.frx":0F96
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleLeft 
      Height          =   450
      Left            =   2040
      Picture         =   "RepParam1.frx":16E0
      Top             =   0
      Width           =   285
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Адрес:"
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
      TabIndex        =   5
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Категория расчета:"
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
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Расчет от:"
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
      Top             =   720
      Width           =   1695
   End
End
Attribute VB_Name = "RepParam1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Категория As String
Dim Подразделение As String
Dim Адрес As String

Private Sub Combo1_Validate(Cancel As Boolean)
Подразделение = Trim(Combo1.Text)

End Sub
Private Sub Combo2_Validate(Cancel As Boolean)
Категория = Combo2.Text
End Sub

Private Sub Combo3_Click()
If Combo3.Text <> "Все" Then
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Command8.Enabled = False

Else
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Command8.Enabled = True
End If
End Sub

Private Sub Combo3_Validate(Cancel As Boolean)
Адрес = Combo3.Text


End Sub

Private Sub Command1_Click()

If Combo2.Text = "Все" Then
MsgBox "Выбери категорию"
Combo2.SetFocus
Label2.ForeColor = vbRed
Label2.FontBold = True
Exit Sub
End If

If Combo1.Text = "Количества жильцов" And Адрес = "Все" Then

Analizlgot.Titl = "РАСЧЕТ" + vbNewLine + "   на возмещение разницы в тарифах по жилищно-коммунальным услугам льготным категориям граждан" + vbNewLine + "  за " + MonthName(Month(MainForm.DR), False) + " " + Str(Year(MainForm.DR))
Analizlgot.G = 18

Reports.sq = "SELECT KLS_PODR.NAIM_KLS AS Адрес, MainOccupant.kv_num AS Кв, Adding.KodKv AS №, MainOccupant.FAM AS Фамилия, MainOccupant.IM AS Имя, MainOccupant.OT AS Отчество, Adding.Propis AS Прописано, Adding.Tarif AS Тариф, Adding.SummaI AS Начислено, Adding.SummaBl AS [Без льгот], [Без льгот]-[Начислено] AS [К возмещению], tmp_lgota.NAME_KLS AS Наименование, Sum([tmp_lgota]![Prim1]) AS [Количество лиг жильцов], tmp_lgota.Procent AS [Процент льгот], tmp_lgota.Use, [Adding]![Tarif]*[Количество лиг жильцов]*[tmp_lgota]![Procent]/100 AS [К воз-ию], Count(tmp_lgota.UniKOd) AS [Кол-во льгот], Adding.ispr FROM KLS_PODR INNER JOIN (MainOccupant INNER JOIN (Adding INNER JOIN tmp_lgota ON Adding.Key = tmp_lgota.UniKOd) ON MainOccupant.Numer = Adding.KodKv) ON KLS_PODR.КОД = MainOccupant.Dom"
Reports.sq = Reports.sq + " GROUP BY KLS_PODR.NAIM_KLS, MainOccupant.kv_num, Adding.KodKv, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.SummaBl, tmp_lgota.NAME_KLS, tmp_lgota.Procent, tmp_lgota.Use, Adding.ispr, Adding.NameKat, tmp_lgota.Prim HAVING (((Adding.ispr)=0) AND ((Adding.NameKat)='" + Combo2.Text + "') AND ((tmp_lgota.Prim)=1)) ORDER BY KLS_PODR.NAIM_KLS, Adding.KodKv"


'Reports.sq = "SELECT KLS_PODR.NAIM_KLS AS Адрес, MainOccupant.kv_num AS Кв, Adding.KodKv AS №, MainOccupant.FAM AS Фамилия, MainOccupant.IM AS Имя, MainOccupant.OT AS Отчество, Adding.ObPl AS [Общ пл], Adding.Propis AS Прописано, Adding.Tarif AS Тариф, Adding.SummaI AS Начислено, Adding.SummaBl AS [Без льгот], Sum([Без льгот]-[Начислено]) AS [К возмещению], tmp_lgota.NAME_KLS AS Наименование, tmp_lgota.PloLG AS [Лиг пло], tmp_lgota.Procent AS [Процент льгот], [Adding]![Tarif]*[tmp_lgota]![PloLG]*[tmp_lgota]![Procent]/100 AS [К воз-ию], Count(tmp_lgota.UniKOd) AS [Кол-во льгот] FROM KLS_PODR INNER JOIN (MainOccupant INNER JOIN (Adding INNER JOIN tmp_lgota ON Adding.Key = tmp_lgota.UniKOd) ON MainOccupant.Numer = Adding.KodKv) ON KLS_PODR.КОД = MainOccupant.Dom"
'Reports.sq = Reports.sq + " GROUP BY KLS_PODR.NAIM_KLS, MainOccupant.kv_num, Adding.KodKv, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.SummaBl, tmp_lgota.NAME_KLS, tmp_lgota.PloLG, tmp_lgota.Procent, [Adding]![Tarif]*[tmp_lgota]![PloLG]*[tmp_lgota]![Procent]/100, Adding.NameKat, tmp_lgota.Prim HAVING (((Adding.NameKat)=" + Chr(34) + Combo2.Text + Chr(34) + ") AND ((tmp_lgota.Prim)=1)) ORDER BY KLS_PODR.NAIM_KLS"


Analizlgot.FG1.Subtotal flexSTSum, 0, 13, , RGB(150, 150, 200), vbBlack, True, "Всего"

Analizlgot.FG1.Subtotal flexSTSum, 0, 16, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 17, , RGB(150, 250, 200), vbBlack, True



Analizlgot.FG1.Subtotal flexSTSum, 1, 13, , RGB(150, 250, 200), vbBlack, True, "И ТОГО ПО ДОМУ"

Analizlgot.FG1.Subtotal flexSTSum, 1, 16, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 1, 17, , RGB(150, 250, 200), vbBlack, True




End If


If Combo1.Text = "Количества жильцов" And Адрес <> "Все" Then

Analizlgot.Titl = "РАСЧЕТ" + vbNewLine + "   на возмещение разницы в тарифах по жилищно-коммунальным услугам льготным категориям граждан" + vbNewLine + "  за " + MonthName(Month(MainForm.DR), False) + " " + Str(Year(MainForm.DR))
Analizlgot.G = 18

Reports.sq = "SELECT KLS_PODR.NAIM_KLS AS Адрес, MainOccupant.kv_num AS Кв, Adding.KodKv AS №, MainOccupant.FAM AS Фамилия, MainOccupant.IM AS Имя, MainOccupant.OT AS Отчество, Adding.Propis AS Прописано, Adding.Tarif AS Тариф, Adding.SummaI AS Начислено, Adding.SummaBl AS [Без льгот], [Без льгот]-[Начислено] AS [К возмещению], tmp_lgota.NAME_KLS AS Наименование, Sum([tmp_lgota]![Prim1]) AS [Количество лиг жильцов], tmp_lgota.Procent AS [Процент льгот], tmp_lgota.Use, [Adding]![Tarif]*[Количество лиг жильцов]*[tmp_lgota]![Procent]/100 AS [К воз-ию], Count(tmp_lgota.UniKOd) AS [Кол-во льгот], Adding.ispr FROM KLS_PODR INNER JOIN (MainOccupant INNER JOIN (Adding INNER JOIN tmp_lgota ON Adding.Key = tmp_lgota.UniKOd) ON MainOccupant.Numer = Adding.KodKv) ON KLS_PODR.КОД = MainOccupant.Dom"
Reports.sq = Reports.sq + " GROUP BY KLS_PODR.NAIM_KLS, MainOccupant.kv_num, Adding.KodKv, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.SummaBl, tmp_lgota.NAME_KLS, tmp_lgota.Procent, tmp_lgota.Use, Adding.ispr, Adding.NameKat, tmp_lgota.Prim HAVING (((KLS_PODR.NAIM_KLS)='" + Адрес + "') AND ((Adding.ispr)=0) AND ((Adding.NameKat)='" + Combo2.Text + "') AND ((tmp_lgota.Prim)=1)) ORDER BY KLS_PODR.NAIM_KLS, Adding.KodKv"



'Reports.sq = "SELECT KLS_PODR.NAIM_KLS AS Адрес, MainOccupant.kv_num AS Кв, Adding.KodKv AS №, MainOccupant.FAM AS Фамилия, MainOccupant.IM AS Имя, MainOccupant.OT AS Отчество, Adding.ObPl AS [Общ пл], Adding.Propis AS Прописано, Adding.Tarif AS Тариф, Adding.SummaI AS Начислено, Adding.SummaBl AS [Без льгот], Sum([Без льгот]-[Начислено]) AS [К возмещению], tmp_lgota.NAME_KLS AS Наименование, tmp_lgota.PloLG AS [Лиг пло], tmp_lgota.Procent AS [Процент льгот], [Adding]![Tarif]*[tmp_lgota]![PloLG]*[tmp_lgota]![Procent]/100 AS [К воз-ию], Count(tmp_lgota.UniKOd) AS [Кол-во льгот] FROM KLS_PODR INNER JOIN (MainOccupant INNER JOIN (Adding INNER JOIN tmp_lgota ON Adding.Key = tmp_lgota.UniKOd) ON MainOccupant.Numer = Adding.KodKv) ON KLS_PODR.КОД = MainOccupant.Dom"
'Reports.sq = Reports.sq + " GROUP BY KLS_PODR.NAIM_KLS, MainOccupant.kv_num, Adding.KodKv, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.SummaBl, tmp_lgota.NAME_KLS, tmp_lgota.PloLG, tmp_lgota.Procent, [Adding]![Tarif]*[tmp_lgota]![PloLG]*[tmp_lgota]![Procent]/100, Adding.NameKat, tmp_lgota.Prim HAVING (((Adding.NameKat)=" + Chr(34) + Combo2.Text + Chr(34) + ") AND ((tmp_lgota.Prim)=1)) ORDER BY KLS_PODR.NAIM_KLS"


'Analizlgot.FG1.Subtotal flexSTSum, 0, 13, , RGB(150, 250, 200), vbBlack, True, "Всего"

Analizlgot.FG1.Subtotal flexSTSum, 0, 16, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 17, , RGB(150, 250, 200), vbBlack, True



Analizlgot.FG1.Subtotal flexSTSum, 1, 13, , RGB(150, 250, 200), vbBlack, True, "И ТОГО ПО ДОМУ"

Analizlgot.FG1.Subtotal flexSTSum, 1, 16, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 1, 17, , RGB(150, 250, 200), vbBlack, True




End If




If Combo1.Text = "Площади" And Адрес = "Все" Then
Analizlgot.Titl = "РАСЧЕТ" + vbNewLine + "   на возмещение разницы в тарифах по жилищно-коммунальным услугам льготным категориям граждан" + vbNewLine + "  за " + MonthName(Month(MainForm.DR), False) + " " + Str(Year(MainForm.DR))
Analizlgot.G = 18
Reports.sq = "SELECT KLS_PODR.NAIM_KLS AS Адрес, MainOccupant.kv_num AS Кв, Adding.KodKv AS №, MainOccupant.FAM AS Фамилия, MainOccupant.IM AS Имя, MainOccupant.OT AS Отчество, Adding.ObPl AS [Общ пл], Adding.Propis AS Прописано, Adding.Tarif AS Тариф, Adding.SummaI AS Начислено, Adding.SummaBl AS [Без льгот], [Без льгот]-[Начислено] AS [К возмещению], tmp_lgota.NAME_KLS AS Наименование, Sum(tmp_lgota.PloLG) AS [Лиг пло], tmp_lgota.Procent AS [Процент льгот], Adding!Tarif*[Лиг пло]*tmp_lgota!Procent/100 AS [К воз-ию], Count(tmp_lgota.UniKOd) AS [Кол-во льгот], Adding.ispr FROM KLS_PODR INNER JOIN (MainOccupant INNER JOIN (Adding INNER JOIN tmp_lgota ON Adding.Key=tmp_lgota.UniKOd) ON MainOccupant.Numer=Adding.KodKv) ON KLS_PODR.КОД=MainOccupant.Dom"
Reports.sq = Reports.sq + " GROUP BY KLS_PODR.NAIM_KLS, MainOccupant.kv_num, Adding.KodKv, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.SummaBl, tmp_lgota.NAME_KLS, tmp_lgota.Procent, Adding.ispr, Adding.NameKat, tmp_lgota.Prim HAVING (((Adding.ispr)=0) AND ((Adding.NameKat)=" + Chr(34) + Combo2.Text + Chr(34) + ") AND ((tmp_lgota.Prim)=1)) ORDER BY KLS_PODR.NAIM_KLS, Adding.KodKv"



Analizlgot.FG1.Subtotal flexSTSum, 0, 14, , RGB(150, 250, 200), vbBlack, True, "Всего"
'Analizlgot.FG1.Subtotal flexSTSum, 0, 15, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 16, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 17, , RGB(150, 250, 200), vbBlack, True



Analizlgot.FG1.Subtotal flexSTSum, 1, 14, , RGB(150, 250, 200), vbBlack, True, "И ТОГО ПО ДОМУ"
'Analizlgot.FG1.Subtotal flexSTSum, 1, 15, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 1, 16, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 1, 17, , RGB(150, 250, 200), vbBlack, True




End If

If Combo1.Text = "Площади" And Адрес <> "Все" Then

Analizlgot.Titl = "РАСЧЕТ" + vbNewLine + "   на возмещение разницы в тарифах по жилищно-коммунальным услугам льготным категориям граждан" + vbNewLine + "  за " + MonthName(Month(MainForm.DR), False) + " " + Str(Year(MainForm.DR))
Analizlgot.G = 18
Reports.sq = "SELECT KLS_PODR.NAIM_KLS AS Адрес, MainOccupant.kv_num AS Кв, Adding.KodKv AS №, MainOccupant.FAM AS Фамилия, MainOccupant.IM AS Имя, MainOccupant.OT AS Отчество, Adding.ObPl AS [Общ пл], Adding.Propis AS Прописано, Adding.Tarif AS Тариф, Adding.SummaI AS Начислено, Adding.SummaBl AS [Без льгот], [Без льгот]-[Начислено] AS [К возмещению], tmp_lgota.NAME_KLS AS Наименование, Sum(tmp_lgota.PloLG) AS [Лиг пло], tmp_lgota.Procent AS [Процент льгот], Adding!Tarif*[Лиг пло]*tmp_lgota!Procent/100 AS [К воз-ию], Count(tmp_lgota.UniKOd) AS [Кол-во льгот] FROM KLS_PODR INNER JOIN (MainOccupant INNER JOIN (Adding INNER JOIN tmp_lgota ON Adding.Key = tmp_lgota.UniKOd) ON MainOccupant.Numer = Adding.KodKv) ON KLS_PODR.КОД = MainOccupant.Dom GROUP BY KLS_PODR.NAIM_KLS, MainOccupant.kv_num, Adding.KodKv, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.SummaBl, tmp_lgota.NAME_KLS, tmp_lgota.Procent, Adding.ispr, Adding.NameKat, tmp_lgota.Prim"


Reports.sq = Reports.sq + " HAVING (((KLS_PODR.NAIM_KLS)='" + Адрес + "') AND ((Adding.ispr)=0) AND ((Adding.NameKat)=" + Chr(34) + Combo2.Text + Chr(34) + ") AND ((tmp_lgota.Prim)=1)) ORDER BY KLS_PODR.NAIM_KLS, Adding.KodKv"


Analizlgot.FG1.Subtotal flexSTSum, 1, 14, , RGB(150, 250, 200), vbBlack, True, "И ТОГО ПО ДОМУ"
'Analizlgot.FG1.Subtotal flexSTSum, 1, 15, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 1, 16, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 1, 17, , RGB(150, 250, 200), vbBlack, t

End If



'MsgBox Reports.sq
'Analizlgot.Об 2
Analizlgot.Show

Unload Me
'Unload RepLgota
Unload Reports
Exit Sub
'Unload RepLgota
Unload Me
End Sub

Private Sub Command2_Click()
If Combo2.Text = "Все" Then
MsgBox "Выбери категорию"
Combo2.SetFocus
Label2.ForeColor = vbRed
Label2.FontBold = True
Exit Sub
End If

If Combo1.Text = "Площади" And Адрес = "Все" Then
Analizlgot.Titl = "РАСЧЕТ" + vbNewLine + "   на возмещение разницы в тарифах по жилищно-коммунальным услугам льготным категориям граждан" + vbNewLine + "  за " + MonthName(Month(MainForm.DR), False) + " " + Str(Year(MainForm.DR))
Analizlgot.G = 14
Reports.sq = "SELECT KLS_PODR.NAIM_KLS AS Адрес, KLS_PODR.Tip_Naim, MainOccupant.kv_num AS Кв, Adding.KodKv AS №, MainOccupant.FAM AS Фамилия, MainOccupant.IM AS Имя, MainOccupant.OT AS Отчество, Adding.ObPl AS [Общ пл], Adding.Propis AS Прописано, Adding.Tarif AS Тариф, Adding.SummaI AS Начислено, Adding.SummaBl AS [Без льгот], [Adding]![SummaBl]-[Adding]![SummaI] AS [К возмещению], Adding.ispr FROM KLS_PODR INNER JOIN (MainOccupant INNER JOIN Adding ON MainOccupant.Numer = Adding.KodKv) ON KLS_PODR.КОД = MainOccupant.Dom GROUP BY KLS_PODR.NAIM_KLS, KLS_PODR.Tip_Naim, MainOccupant.kv_num, Adding.KodKv, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.SummaBl, Adding.ispr, Adding.NameKat Having ((([Adding]![SummaBl] - [Adding]![SummaI]) <> 0) And ((Adding.ispr) = 0) And ((Adding.NameKat) = " + Chr(34) + Combo2.Text + Chr(34) + ")) ORDER BY KLS_PODR.NAIM_KLS, Adding.KodKv"

'Reports.sq = Reports.sq + ""

Analizlgot.FG1.Subtotal flexSTSum, 0, 13, , RGB(150, 250, 200), vbBlack, True, "ВСЕГО"
Analizlgot.FG1.Subtotal flexSTSum, 0, 8, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 9, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 11, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 12, , RGB(150, 250, 200), vbBlack, True

Analizlgot.FG1.Subtotal flexSTSum, 1, 8, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 1, 9, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 1, 11, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 1, 12, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 1, 13, , RGB(150, 250, 200), vbBlack, True


End If

If Combo1.Text = "Площади" And Адрес <> "Все" Then

Analizlgot.Titl = "РАСЧЕТ" + vbNewLine + "   на возмещение разницы в тарифах по жилищно-коммунальным услугам льготным категориям граждан" + vbNewLine + "  за " + MonthName(Month(MainForm.DR), False) + " " + Str(Year(MainForm.DR))
Analizlgot.G = 14
Reports.sq = "SELECT KLS_PODR.NAIM_KLS AS Адрес, KLS_PODR.Tip_Naim, MainOccupant.kv_num AS Кв, Adding.KodKv AS №, MainOccupant.FAM AS Фамилия, MainOccupant.IM AS Имя, MainOccupant.OT AS Отчество, Adding.ObPl AS [Общ пл], Adding.Propis AS Прописано, Adding.Tarif AS Тариф, Adding.SummaI AS Начислено, Adding.SummaBl AS [Без льгот], [Без льгот]-[Начислено] AS [К возмещению], Adding.ispr FROM KLS_PODR INNER JOIN (MainOccupant INNER JOIN (Adding INNER JOIN tmp_lgota ON Adding.Key = tmp_lgota.UniKOd) ON MainOccupant.Numer = Adding.KodKv) ON KLS_PODR.КОД = MainOccupant.Dom GROUP BY KLS_PODR.NAIM_KLS, KLS_PODR.Tip_Naim, MainOccupant.kv_num, Adding.KodKv, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, Adding.ObPl, Adding.Propis, Adding.Tarif, Adding.SummaI, Adding.SummaBl, Adding.ispr, Adding.NameKat, tmp_lgota.Prim HAVING (((KLS_PODR.NAIM_KLS)='" + Адрес + "') AND ((Adding.ispr)=0) AND ((Adding.NameKat)='" + Combo2.Text + "') AND ((tmp_lgota.Prim)=1)) ORDER BY KLS_PODR.NAIM_KLS, Adding.KodKv"
'Reports.sq = Reports.sq + ""


Analizlgot.FG1.Subtotal flexSTSum, 0, 13, , RGB(150, 250, 200), vbBlack, True, "ВСЕГО"
Analizlgot.FG1.Subtotal flexSTSum, 0, 8, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 9, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 11, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 12, , RGB(150, 250, 200), vbBlack, True

Analizlgot.FG1.Subtotal flexSTSum, 1, 8, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 1, 9, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 1, 11, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 1, 12, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 1, 13, , RGB(150, 250, 200), vbBlack, True


End If




Analizlgot.Show

Unload Me
Unload RepLgota
Unload Reports
Exit Sub

Unload Me

End Sub

Private Sub Command3_Click()
If Combo2.Text = "Все" Then
MsgBox "Выбери категорию"
Combo2.SetFocus
Label2.ForeColor = vbRed
Label2.FontBold = True
Exit Sub
End If

If Combo1.Text = "Площади" Then
Analizlgot.Titl = "РАСЧЕТ" + vbNewLine + "   на возмещение разницы в тарифах по жилищно-коммунальным услугам льготным категориям граждан" + vbNewLine + "  за " + MonthName(Month(MainForm.DR), False) + " " + Str(Year(MainForm.DR))
Analizlgot.G = 8
Reports.sq = "SELECT tmp_lgota.NAME_KLS AS Наименование, tmp_lgota.Procent AS [Процент льгот], tmp_lgota.Use, Adding.Tarif AS Тариф, round(Sum(tmp_lgota.PloLG),2) AS [Лиг пло], round(([Adding]![Tarif]*[Лиг пло]*[tmp_lgota]![Procent]/100),2) AS [К воз-ию], Count(tmp_lgota.UniKOd) AS [Кол-во льгот], Adding.ispr FROM Adding INNER JOIN tmp_lgota ON Adding.Key = tmp_lgota.UniKOd GROUP BY tmp_lgota.NAME_KLS, tmp_lgota.Procent, tmp_lgota.Use, Adding.Tarif, Adding.ispr, Adding.NameKat, tmp_lgota.Prim HAVING (((Adding.ispr)=0) AND ((Adding.NameKat)='" + Combo2.Text + "') AND ((tmp_lgota.Prim)=1)) ORDER BY tmp_lgota.NAME_KLS, Sum(tmp_lgota.PloLG)"


Analizlgot.FG1.Subtotal flexSTSum, 0, 5, , RGB(150, 250, 200), vbBlack, True, "ВСЕГО"
Analizlgot.FG1.Subtotal flexSTSum, 0, 7, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 6, , RGB(150, 250, 200), vbBlack, True

Analizlgot.FG1.Subtotal flexSTSum, 1, 7, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 1, 6, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 1, 5, , RGB(150, 250, 200), vbBlack, True

'Analizlgot.FG1.Subtotal flexSTSum, 1, 13, , RGB(150, 250, 200), vbBlack, True


End If

If Combo1.Text = "Количества жильцов" Then
Analizlgot.Titl = "РАСЧЕТ" + vbNewLine + "   на возмещение разницы в тарифах по жилищно-коммунальным услугам льготным категориям граждан" + vbNewLine + "  за " + MonthName(Month(MainForm.DR), False) + " " + Str(Year(MainForm.DR))
Analizlgot.G = 8

Reports.sq = "SELECT tmp_lgota.NAME_KLS AS Наименование, tmp_lgota.Procent AS [Процент льгот], tmp_lgota.Use, Adding.Tarif AS Тариф, Sum([tmp_lgota]![Prim1]) AS [Кол во лиг жильцов], Round(Sum(([Adding]![Tarif]*[tmp_lgota]![Prim1]*[tmp_lgota]![Procent]/100)),2) AS [К воз-ию], Count(tmp_lgota.UniKOd) AS [Кол-во льгот], Adding.ispr FROM Adding INNER JOIN tmp_lgota ON Adding.Key = tmp_lgota.UniKOd GROUP BY tmp_lgota.NAME_KLS, tmp_lgota.Procent, tmp_lgota.Use, Adding.Tarif, Adding.ispr, Adding.NameKat, tmp_lgota.Prim HAVING (((Adding.ispr)=0) AND ((Adding.NameKat)='" + Combo2.Text + "') AND ((tmp_lgota.Prim)=1)) ORDER BY tmp_lgota.NAME_KLS"

'Analizlgot.FG1.Subtotal flexSTSum, 0, 7, , RGB(150, 250, 200), vbBlack, True
'Analizlgot.FG1.Subtotal flexSTSum, 0, 6, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 7, , RGB(150, 250, 200), vbBlack, True, "И ТОГО:"
Analizlgot.FG1.Subtotal flexSTSum, 0, 6, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 5, , RGB(150, 250, 200), vbBlack, True

End If


Analizlgot.Show
Unload Me
Unload RepLgota
Unload Reports
Exit Sub

Unload Me

End Sub

Private Sub Command4_Click()
If Combo2.Text = "Все" Then
MsgBox "Выбери категорию"
Combo2.SetFocus
Label2.ForeColor = vbRed
Label2.FontBold = True
Exit Sub
End If

If Combo1.Text = "Площади" Then
Analizlgot.Titl = "РАСЧЕТ" + vbNewLine + "   на возмещение разницы в тарифах по жилищно-коммунальным услугам льготным категориям граждан" + vbNewLine + "  за " + MonthName(Month(MainForm.DR), False) + " " + Str(Year(MainForm.DR))
Analizlgot.G = 8

Reports.sq = "SELECT tmp_lgota.NAME_KLS AS Наименование, tmp_lgota.Procent AS [Процент льгот], tmp_lgota.Use, Adding.Tarif AS Тариф, round(tmp_lgota.PloLG,2) AS [Лиг пло], Round(Sum(([Adding]![Tarif]*[Лиг пло]*[tmp_lgota]![Procent]/100)),2) AS [К воз-ию], Count(tmp_lgota.UniKOd) AS [Кол-во льгот], Adding.ispr FROM Adding INNER JOIN tmp_lgota ON Adding.Key = tmp_lgota.UniKOd GROUP BY tmp_lgota.NAME_KLS, tmp_lgota.Procent, tmp_lgota.Use, Adding.Tarif, Adding.ispr, tmp_lgota.PloLG, Adding.NameKat, tmp_lgota.Prim HAVING (((Adding.ispr)=0) AND ((Adding.NameKat)='" + Combo2.Text + "') AND ((tmp_lgota.Prim)=1)) ORDER BY tmp_lgota.NAME_KLS, tmp_lgota.PloLG"

Analizlgot.FG1.Subtotal flexSTSum, 0, 7, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 6, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 1, 7, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 1, 6, , RGB(150, 250, 200), vbBlack, True

End If


If Combo1.Text = "Количества жильцов" Then
Analizlgot.Titl = "РАСЧЕТ" + vbNewLine + "   на возмещение разницы в тарифах по жилищно-коммунальным услугам льготным категориям граждан" + vbNewLine + "  за " + MonthName(Month(MainForm.DR), False) + " " + Str(Year(MainForm.DR))
Analizlgot.G = 8

Reports.sq = "SELECT tmp_lgota.NAME_KLS AS Наименование, tmp_lgota.Procent AS [Процент льгот], tmp_lgota.Use, Adding.Tarif AS Тариф, [tmp_lgota]![Prim1] AS [Кол во лиг жильцов], Round(Sum(([Adding]![Tarif]*[tmp_lgota]![Prim1]*[tmp_lgota]![Procent]/100)),2) AS [К воз-ию], Count(tmp_lgota.UniKOd) AS [Кол-во льгот], Adding.ispr FROM Adding INNER JOIN tmp_lgota ON Adding.Key = tmp_lgota.UniKOd GROUP BY tmp_lgota.NAME_KLS, tmp_lgota.Procent, tmp_lgota.Use, Adding.Tarif, [tmp_lgota]![Prim1], Adding.ispr, Adding.NameKat, tmp_lgota.Prim HAVING (((Adding.ispr)=0) AND ((Adding.NameKat)='" + Combo2.Text + "') AND ((tmp_lgota.Prim)=1)) ORDER BY tmp_lgota.NAME_KLS"

Analizlgot.FG1.Subtotal flexSTSum, 0, 7, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 6, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 1, 7, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 1, 6, , RGB(150, 250, 200), vbBlack, True

End If





Analizlgot.Show
Unload Me
Unload RepLgota
Unload Reports
End Sub

Private Sub Command5_Click()
If Combo2.Text = "Все" Then
MsgBox "Выбери категорию"
Combo2.SetFocus
Label2.ForeColor = vbRed
Label2.FontBold = True


Exit Sub
End If

Analizlgot.Titl = "РАСЧЕТ" + vbNewLine + "   на возмещение разницы в тарифах по жилищно-коммунальным услугам льготным категориям граждан" + vbNewLine + "  за " + MonthName(Month(MainForm.DR), False) + " " + Str(Year(MainForm.DR))
Analizlgot.G = 10
'Reports.sq = "SELECT tmp_lgota.NAME_KLS AS Наименование, tmp_lgota.Procent AS [Размер льгот], Count(tmp_lgota.UniKOd) AS [Кол-во льгот], Adding.Propis AS [Кол-во чл сем], Sum(tmp_lgota.PloLG) AS [Лиг площадь], Adding.ObPl AS [Общ пл], Adding.SummaI AS Начислено, Adding.SummaBl AS [Без льгот], [Без льгот]-[Начислено] AS [К возмещению] FROM Adding INNER JOIN tmp_lgota ON Adding.Key = tmp_lgota.UniKOd GROUP BY tmp_lgota.NAME_KLS, tmp_lgota.Procent, Adding.Propis, Adding.ObPl, Adding.SummaI, Adding.SummaBl, Adding.NameKat, tmp_lgota.Prim HAVING (((Adding.NameKat)=" + Chr(34) + Combo2.Text + Chr(34) + ") AND ((tmp_lgota.Prim)=1))"

'Reports.sq = "SELECT tmp_lgota.NAME_KLS AS Наименование, tmp_lgota.Procent AS [Размер льгот], Count(tmp_lgota.UniKOd) AS [Кол-во льгот], Adding.Propis AS [Кол-во чл сем], Sum(tmp_lgota.PloLG) AS [Лиг площадь], Adding.ObPl AS [Общ пл], Adding.SummaI AS Начислено, Adding.SummaBl AS [Без льгот], ([Adding]![Tarif]*[tmp_lgota]![Procent]*[tmp_lgota]![PloLG])/100 AS [К возмещениюию] FROM Adding INNER JOIN tmp_lgota ON Adding.Key = tmp_lgota.UniKOd GROUP BY tmp_lgota.NAME_KLS, tmp_lgota.Procent, Adding.Propis, Adding.ObPl, Adding.SummaI, Adding.SummaBl, ([Adding]![Tarif]*[tmp_lgota]![Procent]*[tmp_lgota]![PloLG])/100, Adding.NameKat, tmp_lgota.Prim HAVING (((Adding.NameKat)='" + Combo2.Text + "') AND ((tmp_lgota.Prim)=1))"

If Combo1.Text = "Площади" Then
Reports.sq = "SELECT tmp_lgota.NAME_KLS AS Наименование, tmp_lgota.Procent AS [Размер льгот], Adding.Propis AS [Кол-во чл сем], Sum(tmp_lgota.PloLG) AS [Лиг площадь], Adding.ObPl AS [Общ пл], Adding.SummaI AS Начислено, Adding.SummaBl AS [Без льгот], Sum(([Adding]![Tarif]*[tmp_lgota]![Procent]*[tmp_lgota]![PloLG])/100) AS [К возмещениюию],  Count(tmp_lgota.UniKOd) AS [Кол-во льгот], Adding.ispr FROM Adding INNER JOIN tmp_lgota ON Adding.Key = tmp_lgota.UniKOd GROUP BY tmp_lgota.NAME_KLS, tmp_lgota.Procent, Adding.Propis, Adding.ObPl, Adding.SummaI, Adding.SummaBl, Adding.ispr, Adding.NameKat, tmp_lgota.Prim HAVING (((Adding.ispr)=0) AND ((Adding.NameKat)='" + Combo2.Text + "') AND ((tmp_lgota.Prim)=1))"
End If


If Combo1.Text = "Количества жильцов" Then
Reports.sq = "SELECT tmp_lgota.NAME_KLS AS Наименование, tmp_lgota.Procent AS [Размер льгот], Adding.Propis AS [Кол-во чл сем], Sum(tmp_lgota.Prim1) AS [Кол лиг жильцов], Adding.ObPl AS [Общ пл], Adding.SummaI AS Начислено, Adding.SummaBl AS [Без льгот], Sum(([Adding]![Tarif]*[tmp_lgota]![Procent]*[tmp_lgota]![Prim1])/100) AS [К возмещениюию],  Count(tmp_lgota.UniKOd) AS [Кол-во льгот], Adding.ispr FROM Adding INNER JOIN tmp_lgota ON Adding.Key = tmp_lgota.UniKOd GROUP BY tmp_lgota.NAME_KLS, tmp_lgota.Procent, Adding.Propis, Adding.ObPl, Adding.SummaI, Adding.SummaBl, Adding.ispr, Adding.NameKat, tmp_lgota.Prim HAVING (((Adding.ispr)=0) AND ((Adding.NameKat)='" + Combo2.Text + "') AND ((tmp_lgota.Prim)=1))"
End If



Analizlgot.Об 2

'Analizlgot.FG1.Subtotal flexSTSum, 0, 8, , RGB(150, 250, 200), vbBlack, True
'Analizlgot.FG1.Subtotal flexSTSum, 1, 8, , RGB(150, 250, 200), vbBlack, True
'Analizlgot.FG1.Subtotal flexSTClear, 2, 7, , RGB(150, 250, 100), vbBlack, True

Analizlgot.Show




Unload Me
Unload RepLgota
Unload Reports
Exit Sub
'Unload RepLgota


End Sub

Private Sub Command6_Click()

If Combo2.Text = "Все" Then
MsgBox "Выбери категорию"
Combo2.SetFocus
Label2.ForeColor = vbRed
Label2.FontBold = True
Exit Sub
End If



Analizlgot.Titl = "Необходимо проверить нижеперечисленные счета на " + vbNewLine + "правильность расчета льгот, а так же лиготируемых площадей" + vbNewLine + "  за " + MonthName(Month(MainForm.DR), False) + " " + Str(Year(MainForm.DR))
Analizlgot.G = 7


If Combo1.Text = "Площади" Then
Reports.sq = "SELECT TEST_02.№, TEST_02.Фамилия, TEST_02.Имя, TEST_02.Отчество, TEST_02.Начислено, TEST_02.[Без льгот], Test_03.[Sum-К воз-ию], TEST_02.[К возмещению], Round([TEST_02]![К возмещению]-[TEST_03]![Sum-К воз-ию],2) AS Отклонение, TEST_02.NameKat FROM TEST_02 INNER JOIN Test_03 ON TEST_02.№ = Test_03.№ WHERE (((Round([TEST_02]![К возмещению]-[TEST_03]![Sum-К воз-ию],2))<-0.01 Or (Round([TEST_02]![К возмещению]-[TEST_03]![Sum-К воз-ию],2))>0.01) AND ((TEST_02.NameKat)='" + Combo2.Text + "')) ORDER BY Round([TEST_02]![К возмещению]-[TEST_03]![Sum-К воз-ию],2)"
Analizlgot.FG1.Subtotal flexSTSum, 0, 6, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 5, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 4, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 2, , RGB(150, 250, 200), vbBlack, True
End If

'If Combo1.Text = "Количества жильцов" Then
'Reports.sq = "SELECT [02].№, [02].Фамилия, [02].Имя, [02].Отчество, Round([02]![К возмещению]-[03]![Sum-К воз-ию],2) AS Расхождение, [02].Тариф, [02].Начислено, [02].[Без льгот], [02].[К возмещению], [02].NameKat FROM 02 INNER JOIN 03 ON [02].№ = [03].№ Where (((Round([02]![К возмещению] - [03]![Sum-К воз-ию], 2)) < -0.01 Or (Round([02]![К возмещению] - [03]![Sum-К воз-ию], 2)) > 0.01) And (([02].NameKat) = '" + Combo2.Text + "')) ORDER BY Round([02]![К возмещению]-[03]![Sum-К воз-ию],2)"
'End If


'Analizlgot.Об 2
Analizlgot.Show

Unload Me
Unload RepLgota
Unload Reports
End Sub

Private Sub Command7_Click()
If Combo2.Text = "Все" Then
MsgBox "Выбери категорию"
Combo2.SetFocus
Label2.ForeColor = vbRed
Label2.FontBold = True
Exit Sub
End If



Analizlgot.Titl = "Необходимо проверить нижеперечисленные счета на " + vbNewLine + "правильность расчета льгот, а так же лиготируемых площадей" + vbNewLine + "  за " + MonthName(Month(MainForm.DR), False) + " " + Str(Year(MainForm.DR))
Analizlgot.G = 12


'If Combo1.Text = "Площади" Then
Reports.sq = "SELECT TEST_04.№, TEST_04.Фамилия, TEST_04.Имя, TEST_04.Отчество, TEST_04.[Количество лиг жильцов], TEST_04.Тариф, Round(TEST_04![К возмещению]-Test_05![Sum-К воз-ию],2) AS Расхождение, TEST_04.Начислено, TEST_04.[Без льгот], TEST_04.[К возмещению], TEST_05.[Sum-К воз-ию], TEST_04.NameKat, TEST_04.[Количество лиг жильцов] FROM TEST_04 INNER JOIN TEST_05 ON TEST_04.№ = TEST_05.№ WHERE (((TEST_04.NameKat)='вывоз мусора') AND ((Round([TEST_04]![К возмещению]-[Test_05]![Sum-К воз-ию],2))<-0.01 Or (Round([TEST_04]![К возмещению]-[Test_05]![Sum-К воз-ию],2))>0.01)) ORDER BY Round(TEST_04![К возмещению]-Test_05![Sum-К воз-ию],2)"

Analizlgot.FG1.Subtotal flexSTSum, 0, 9, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 5, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 7, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 8, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 10, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 11, , RGB(150, 250, 200), vbBlack, True
'Analizlgot.FG1.Subtotal flexSTSum, 0, 12, , RGB(150, 250, 200), vbBlack, True
'End If


'Analizlgot.Об 2
Analizlgot.Show

Unload Me
Unload RepLgota
Unload Reports
End Sub

Private Sub Command8_Click()
If Combo2.Text = "Все" Then
MsgBox "Выбери категорию"
Combo2.SetFocus
Label2.ForeColor = vbRed
Label2.FontBold = True
Exit Sub
End If

Analizlgot.Titl = "Отчет о фактических расходах, связанных с оказанием мер социальной поддержки " + vbNewLine + "отдельных категорий граждан в части оплаты жилищно-коммунальных услуг за " + MainForm.Label8 + " г."

If Combo1.Text = "Площади" Then

If MsgBox("Сворачивать лиготируемую площадь не равную 18,21 и 33 кв.м. в ПРОЧИЕ ?", vbYesNo) = vbYes Then

Analizlgot.G = 12
Reports.sq = "SELECT LGTip.Name AS [Тип льготы], TMP_Lgota.NAME_KLS AS Льгота, TMP_Lgota.Use AS [Способ применения], TMP_Lgota.Procent AS [Размер льгот], TMP_Lgota.tarif, IIf([PloLG]<>18,IIf([PloLG]<>21,IIf([PloLG]<>33,0,TMP_Lgota!PloLG),TMP_Lgota!PloLG),TMP_Lgota!PloLG) AS [Лиготируемая площадь], Sum(TMP_Lgota.Prop) AS [Всего прописано], Count([TMP_Lgota]![Key]) AS Получатели, [Количество льгот]-[Получатели] AS [Члены семей], Sum(TMP_Lgota!Koll) AS [Количество льгот], Round(Sum(([TMP_Lgota]![Procent]*[TMP_Lgota]![PloLG]/100)*[TMP_Lgota]![tarif]),2) AS [К возмещению] FROM Adding INNER JOIN ((TMP_Lgota LEFT JOIN KLS_PRIV ON TMP_Lgota.KodKls = KLS_PRIV.N_KLS) LEFT JOIN LGTip ON KLS_PRIV.Tip = LGTip.Tip) ON Adding.Key = TMP_Lgota.UniKOd Where (((Adding.NameKat) = '" + Combo2.Text + "') And ((TMP_Lgota.Prim) > 0)) GROUP BY LGTip.Name, TMP_Lgota.NAME_KLS, TMP_Lgota.Use, TMP_Lgota.Procent, TMP_Lgota.tarif, IIf([PloLG]<>18,IIf([PloLG]<>21,IIf([PloLG]<>33,0,TMP_Lgota!PloLG),TMP_Lgota!PloLG),TMP_Lgota!PloLG)"

Analizlgot.FG1.MergeCells = flexMergeFree



Analizlgot.FG1.Subtotal flexSTSum, 0, 9, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 7, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 8, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 10, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 11, , RGB(150, 250, 200), vbBlack, True

Analizlgot.FG1.Subtotal flexSTSum, 2, 9, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 2, 7, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 2, 8, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 2, 10, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 2, 11, , RGB(150, 250, 200), vbBlack, True

Else

Analizlgot.G = 12
Reports.sq = "SELECT LGTip.Name AS [Тип льготы], TMP_Lgota.NAME_KLS AS Льгота, TMP_Lgota.Use AS [Способ применения], TMP_Lgota.Procent AS [Размер льгот], TMP_Lgota.tarif AS Тариф, [TMP_Lgota]![PloLG] AS [Лиготируемая площадь], Sum(TMP_Lgota.Prop) AS [Всего прописано], Count(TMP_Lgota!Key) AS Получатели, [Количество льгот]-[Получатели] AS [Члены семей], Sum(TMP_Lgota!Koll) AS [Количество льгот], Round(Sum((TMP_Lgota!Procent*TMP_Lgota!PloLG/100)*TMP_Lgota!tarif),2) AS [К возмещению] FROM Adding INNER JOIN ((TMP_Lgota LEFT JOIN KLS_PRIV ON TMP_Lgota.KodKls = KLS_PRIV.N_KLS) LEFT JOIN LGTip ON KLS_PRIV.Tip = LGTip.Tip) ON Adding.Key = TMP_Lgota.UniKOd Where (((Adding.NameKat) = '" + Combo2.Text + "') And ((TMP_Lgota.Prim) > 0)) GROUP BY LGTip.Name, TMP_Lgota.NAME_KLS, TMP_Lgota.Use, TMP_Lgota.Procent, TMP_Lgota.tarif, [TMP_Lgota]![PloLG]"

Analizlgot.FG1.MergeCells = flexMergeFree



Analizlgot.FG1.Subtotal flexSTSum, 0, 9, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 7, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 8, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 10, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 11, , RGB(150, 250, 200), vbBlack, True

Analizlgot.FG1.Subtotal flexSTSum, 2, 9, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 2, 7, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 2, 8, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 2, 10, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 2, 11, , RGB(150, 250, 200), vbBlack, True



End If


Analizlgot.Show

Unload RepLgota
Unload Reports
Unload Me




End If

If Combo1.Text = "Количества жильцов" Then


If MsgBox("Развернутый по количеству прописанных?", vbYesNo) = vbYes Then


Analizlgot.G = 11
Reports.sq = "SELECT LGTip.Name AS [Тип льготы], TMP_Lgota.NAME_KLS AS Льгота, TMP_Lgota.Use AS [Способ применения], TMP_Lgota.Procent AS [Размер льгот], TMP_Lgota.tarif, TMP_Lgota.Prop AS [Всего прописано], Count(TMP_Lgota!Key) AS Получатели, [Количество льгот]-[Получатели] AS [Члены семей], Sum(TMP_Lgota!Koll) AS [Количество льгот], Round(Sum(([TMP_Lgota]![Procent]*[TMP_Lgota]![Koll]/100)*[TMP_Lgota]![tarif]),2) AS [К возмещению] FROM Adding INNER JOIN ((TMP_Lgota LEFT JOIN KLS_PRIV ON TMP_Lgota.KodKls = KLS_PRIV.N_KLS) LEFT JOIN LGTip ON KLS_PRIV.Tip = LGTip.Tip) ON Adding.Key = TMP_Lgota.UniKOd Where (((Adding.NameKat) = '" + Combo2.Text + "') And ((TMP_Lgota.Prim) > 0)) GROUP BY LGTip.Name, TMP_Lgota.NAME_KLS, TMP_Lgota.Use, TMP_Lgota.Procent, TMP_Lgota.tarif, TMP_Lgota.Prop"
Analizlgot.FG1.MergeCells = flexMergeFree


Analizlgot.FG1.Subtotal flexSTSum, 0, 6, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 9, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 7, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 8, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 10, , RGB(150, 250, 200), vbBlack, True


Analizlgot.FG1.Subtotal flexSTSum, 2, 6, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 2, 9, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 2, 7, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 2, 8, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 2, 10, , RGB(150, 250, 200), vbBlack, True


Else

Analizlgot.G = 10
Reports.sq = "SELECT LGTip.Name AS [Тип льготы], TMP_Lgota.NAME_KLS AS Льгота, TMP_Lgota.Use AS [Способ применения], TMP_Lgota.Procent AS [Размер льгот], TMP_Lgota.tarif, Count(TMP_Lgota!Key) AS Получатели, [Количество льгот]-[Получатели] AS [Члены семей], Sum(TMP_Lgota!Koll) AS [Количество льгот], Round(Sum((TMP_Lgota!Procent*TMP_Lgota!Koll/100)*TMP_Lgota!tarif),2) AS [К возмещению] FROM Adding INNER JOIN ((TMP_Lgota LEFT JOIN KLS_PRIV ON TMP_Lgota.KodKls = KLS_PRIV.N_KLS) LEFT JOIN LGTip ON KLS_PRIV.Tip = LGTip.Tip) ON Adding.Key = TMP_Lgota.UniKOd Where (((Adding.NameKat) = '" + Combo2.Text + "') And ((TMP_Lgota.Prim) > 0)) GROUP BY LGTip.Name, TMP_Lgota.NAME_KLS, TMP_Lgota.Use, TMP_Lgota.Procent, TMP_Lgota.tarif"
Analizlgot.FG1.MergeCells = flexMergeFree


Analizlgot.FG1.Subtotal flexSTSum, 0, 6, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 9, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 7, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 0, 8, , RGB(150, 250, 200), vbBlack, True


Analizlgot.FG1.Subtotal flexSTSum, 2, 6, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 2, 9, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 2, 7, , RGB(150, 250, 200), vbBlack, True
Analizlgot.FG1.Subtotal flexSTSum, 2, 8, , RGB(150, 250, 200), vbBlack, True




End If





Analizlgot.Show

Unload Me
Unload RepLgota
Unload Reports



End If

End Sub

Private Sub Form_Load()
Dim cnParam As ADODB.Connection
Dim rsVrem As ADODB.Recordset

MakeWindow Me, True

Коннект "kvartplata.amd"

'Set cnParam = New ADODB.Connection
 ' cnParam.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Persist Security Info=True"
 ' cnParam.Open "data/Kvartplata.mdb"
    
Set rsVrem = New ADODB.Recordset
Set rsVrem.ActiveConnection = Mconn
 
'Расчет от
Combo1.Text = "Площади"
Combo1.AddItem "Площади"
Combo1.AddItem "Количества жильцов"


'Категория расчета


rsVrem.Open ("SELECT Kategor.Name_Kategor FROM Kategor")
rsVrem.MoveFirst
Combo2.AddItem "Все"
Do While Not rsVrem.EOF
Combo2.AddItem rsVrem.Fields("Name_Kategor")
rsVrem.MoveNext
Loop
rsVrem.Close

'Адрес
Адрес = "Все"
rsVrem.Open ("SELECT KLS_PODR.NAIM_KLS, KLS_PODR.Num FROM KLS_PODR")
rsVrem.MoveFirst
Combo3.AddItem "Все"
Do While Not rsVrem.EOF
Combo3.AddItem rsVrem.Fields("NAIM_KLS")
rsVrem.MoveNext
Loop
rsVrem.Close

lblTitle.Caption = "Параметры отчета"
Set cnParam = Nothing
Set rsVrem = Nothing
End Sub

Private Sub imgTitleHelp_Click()
Unload Me
End Sub
